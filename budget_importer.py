"""
Budget CSV Importer
====================
Standardizes transactions from Wings Visa Platinum and Apple Card CSVs
and appends them to a Google Sheets budget tracker.

Columns in master sheet: Cost | Location | Category | Date | Account

USAGE:
  1. Drop CSV files into the inbox/ directory.
  2. Run: python budget_importer.py
  3. Processed files are moved to archive/ automatically.

File type is detected automatically from the filename:
  - Contains "apple"  → Apple Card
  - Contains "visa"   → Wings Visa Platinum
  - Contains "debit"  → Wings Debit

SETUP (one-time):
  pip install pandas gspread google-auth

  Then follow the Google Sheets API setup instructions printed below.
"""

import re
import sys
import pandas as pd
from datetime import datetime
from pathlib import Path

# ──────────────────────────────────────────────
# CONFIGURATION — edit these
# ──────────────────────────────────────────────

GOOGLE_SHEET_NAME    = "Budget 2026"       # Exact name of your Google Sheet
WORKSHEET_NAME       = "Transactions"      # Tab name for expenses
INCOME_WORKSHEET_NAME = "Income"           # Tab name for income (debit credits)
CREDENTIALS_FILE     = "credentials.json" # Path to your Google service account JSON

# Directories for the drop-and-run workflow
INBOX_DIR   = Path("inbox")
ARCHIVE_DIR = Path("archive")

# Set to True to do a dry run (prints rows, does not write to sheet)
DRY_RUN = False

# Supported file types and their parser functions
FILE_TYPES = ["wings_visa", "apple_card", "wings_debit", "amex"]

# ──────────────────────────────────────────────
# CATEGORY MAPPING
# ──────────────────────────────────────────────

# Apple Card category → your categories
APPLE_CATEGORY_MAP = {
    "restaurants":        "Dining",
    "grocery":            "Grocery",
    "gas":                "Gas",
    "utilities":          "Bill",
    "entertainment":      "Entertainment",
    "transportation":     "Travel",
    "govt-services-parking": "Car",
    "other":              "Others",
    # Skip these row types entirely
    "payment":            None,
    "installment":        "Bill",
}

# Merchant keyword overrides (checked after category mapping, case-insensitive)
# More specific rules first. First match wins.
MERCHANT_OVERRIDES = {
    # Cat
    "petsmart":           "Cat",
    "petco":              "Cat",
    "chewy":              "Cat",

    # Gaming
    "steam":              "Gaming",
    "microsoft*pc game":  "Gaming",
    "microsoft":          "Gaming",   # catches Microsoft/Xbox Game Pass
    "xbox":               "Gaming",
    "playstation":        "Gaming",
    "playstat":           "Gaming",   # catches STORE.PLAYSTAT short form
    "nintendo":           "Gaming",

    # Bars
    "casual animal brew": "Bars",
    "john's big deck":    "Bars",
    "session taco":       "Dining",   # restaurant, not bar
    "king g":             "Bars",

    # Entertainment
    "amc":                "Entertainment",
    "nelson-atkins":      "Entertainment",

    # Travel
    "uber":               "Travel",
    "lyft":               "Travel",
    "airbnb":             "Travel",
    "united":             "Travel",   # United Airlines
    "southwest":          "Travel",   # Southwest Airlines
    "delta":              "Travel",
    "american airlines":  "Travel",
    "avis":               "Travel",   # car rental
    "hertz":              "Travel",
    "enterprise rent":    "Travel",
    "agent fee":          "Travel",   # travel agency fees

    # Shopping
    "amazon":             "Shopping",
    "walmart":            "Shopping",
    "wal-mart":           "Shopping",
    "micro center":       "Shopping",
    "shopdap":            "Shopping",

    # Car
    "molle automotive":   "Car",
    "carwash":            "Car",
    "motor vehicle":      "Car",

    # Bill / Subscriptions
    "netflix":            "Bill",
    "disney":             "Bill",
    "apple.com/bill":     "Bill",
    "apple services":     "Bill",     # catches Apple Card's cleaned merchant name
    "amazon prime":       "Bill",
    "wikimedia":          "Bill",
    "claude.ai":          "Bill",
    "caribou coffee":     "Dining",
    "meshuggah bagels":   "Dining",   # bagel cafe, not grocery

    # Grocery clarifications
    "hy-vee":             "Grocery",  # Hy-Vee grocery stores (not gas)
}

# AMEX category string → your categories (most-specific prefixes first)
AMEX_CATEGORY_MAP = {
    "merchandise & supplies-groceries": "Grocery",
    "business services-insurance":      "Bill",
    "other-government":                 "Others",
    "automotive-automotive repair":     "Car",
    "restaurant":                       "Dining",
    "fees & adjustments":               "Bill",
    "merchandise & supplies":           "Shopping",
    "transportation":                   "Gas",
    "travel":                           "Travel",
    "business services":                "Others",
    "entertainment":                    "Entertainment",
    "automotive":                       "Gas",
    "personal":                         "Others",
    "healthcare":                       "Others",
    "other":                            "Others",
}

# For Visa: keyword → category (merchant descriptions are messy)
VISA_KEYWORD_CATEGORIES = {
    "mcdonald":           "Dining",
    "culver":             "Dining",
    "chick-fil":          "Dining",
    "popeyes":            "Dining",
    "sonic":              "Dining",
    "bc bistro":          "Dining",
    "bamboo penny":       "Dining",
    "jerusalem cafe":     "Dining",
    "kc craft ramen":     "Dining",
    "session taco":       "Dining",
    "epic vision":        "Others",   # eye care
    "wal-mart":           "Grocery",
    "wm supercenter":     "Grocery",
    "city of kansas":     "Bill",
    "ysi*maa":            "Bill",     # insurance/savings — adjust if needed
    "carwash":            "Car",
    "molle automotive":   "Car",
    "amc":                "Entertainment",
    "nelson-atkins":      "Entertainment",
    "shopdap":            "Shopping",
    "payment":            None,       # skip credit card payments
}

# Debit: descriptions that are credit card payments or internal transfers — skip to avoid
# double-counting (the transactions are already captured on the respective card's CSV)
DEBIT_SKIP_KEYWORDS = [
    "applecard gsbank",     # Apple Card payment
    "amex",                 # AMEX payment
    "visa payment",         # Wings Visa payment
    "digital transfer",     # internal Wings account-to-account transfer
    "ian seal",             # self-initiated external fund transfer
]

# Debit: keyword → category for expense transactions
DEBIT_KEYWORD_CATEGORIES = {
    "dept revenue":   "Bill",       # state/federal tax payments
    "dept of rev":    "Bill",
    "apple gs savings": "Savings",  # Apple savings account transfer
    "wealthfront":    "Savings",    # investment/brokerage transfer
    "mcdonald":       "Dining",
    "culver":         "Dining",
    "chick-fil":      "Dining",
    "popeyes":        "Dining",
    "sonic":          "Dining",
    "amazon":         "Shopping",
    "walmart":        "Shopping",
    "wal-mart":       "Shopping",
    "carwash":        "Car",
}

# Debit: keyword → income sub-category for credit transactions
DEBIT_INCOME_CATEGORIES = {
    "payroll":        "Payroll",
    "dividend":       "Dividend",
    "tax refund":     "Tax Refund",
    "refund":         "Tax Refund",
    "mobile deposit": "Deposit",
}

# ──────────────────────────────────────────────
# PARSING FUNCTIONS
# ──────────────────────────────────────────────

def clean_amount(val):
    """Convert amount strings like '-$9.81' or '9.81' to a positive float."""
    val = str(val).replace("$", "").replace(",", "").strip()
    return round(abs(float(val)), 2)


def parse_date(val):
    """Normalize various date formats to MM/DD/YYYY."""
    for fmt in ("%m/%d/%y", "%m/%d/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(str(val).strip(), fmt).strftime("%m/%d/%Y")
        except ValueError:
            continue
    return str(val).strip()


def apply_merchant_override(merchant_str):
    """Check merchant string against override keywords. Returns category or None."""
    m = merchant_str.lower()
    for keyword, category in MERCHANT_OVERRIDES.items():
        if keyword in m:
            return category
    return None


def parse_apple_card(filepath):
    """Parse Apple Card CSV into standardized rows."""
    df = pd.read_csv(filepath)
    rows = []

    for _, row in df.iterrows():
        apple_cat = str(row.get("Category", "")).strip().lower()

        # Skip payments
        if apple_cat in ("payment"):
            continue

        merchant = str(row.get("Merchant", "")).strip()
        amount   = clean_amount(row.get("Amount (USD)", 0))
        date     = parse_date(row.get("Transaction Date", ""))

        # Determine category: merchant override takes priority
        override = apply_merchant_override(merchant)
        if override:
            category = override
        else:
            category = APPLE_CATEGORY_MAP.get(apple_cat, "Others")

        # Special case: Hy-Vee gas vs grocery
        # Apple tags Hy-Vee fuel pumps as "Gas" — keep that, but Hy-Vee grocery → Grocery
        if "hy-vee" in merchant.lower() and apple_cat == "gas":
            category = "Gas"

        rows.append({
            "Cost":     amount,
            "Location": merchant,
            "Category": category,
            "Date":     date,
            "Account":  "Apple Card",
        })

    return rows


def parse_wings_visa(filepath):
    """Parse Wings Visa Platinum CSV into standardized rows."""
    df = pd.read_csv(filepath)
    rows = []

    for _, row in df.iterrows():
        description = str(row.get("Description", "")).strip()
        amount_raw  = str(row.get("Amount", "")).strip()

        # Skip payments (positive amounts or 'PAYMENT' in description)
        if "payment" in description.lower():
            continue
        # Also skip rows with no dollar sign (blank/balance rows)
        if "$" not in amount_raw:
            continue

        amount = clean_amount(amount_raw)
        date   = parse_date(row.get("Date", ""))

        # Extract a cleaner merchant name from the messy description
        # Format: "MERCHANT NAME      CITY  ST REF #..."
        merchant_raw = re.split(r'\s{2,}', description)[0].strip()
        merchant_raw = re.sub(r'\s+REF\s*#.*', '', merchant_raw, flags=re.IGNORECASE).strip()
        # Remove "PRE-AUTH " prefix
        merchant_raw = re.sub(r'^PRE-AUTH\s+', '', merchant_raw, flags=re.IGNORECASE).strip()
        # Remove TST* prefix (Toast POS)
        merchant_raw = re.sub(r'^TST\*\s*', '', merchant_raw, flags=re.IGNORECASE).strip()
        # Title case
        location = merchant_raw.title()

        # Determine category from keyword list
        desc_lower = description.lower()
        category = "Others"
        for keyword, cat in VISA_KEYWORD_CATEGORIES.items():
            if keyword in desc_lower:
                if cat is None:
                    category = None
                    break
                category = cat
                break

        if category is None:
            continue  # skip this row (e.g. payments)

        # Also check merchant overrides
        override = apply_merchant_override(description)
        if override:
            category = override

        rows.append({
            "Cost":     amount,
            "Location": location,
            "Category": category,
            "Date":     date,
            "Account":  "Wings Visa",
        })

    return rows


def parse_wings_debit(filepath):
    """Parse Wings Checking/Debit CSV into standardized rows.

    Returns two lists: (expense_rows, income_rows).
    Expense rows go to the Transactions worksheet; income rows go to the Income worksheet.
    Transfers between accounts (credit card payments, savings moves, etc.) are skipped.
    """
    df = pd.read_csv(filepath)
    expense_rows = []
    income_rows  = []

    for _, row in df.iterrows():
        description = str(row.get("Description", "")).strip()
        amount_raw  = str(row.get("Amount", "")).strip()
        date        = parse_date(row.get("Date", ""))

        if not amount_raw or amount_raw == "nan":
            continue

        desc_lower = description.lower()

        # Skip transfers and inter-account payments
        if any(kw in desc_lower for kw in DEBIT_SKIP_KEYWORDS):
            continue

        # Parse signed amount: positive = credit (income), negative = debit (expense)
        amount_str = amount_raw.replace("$", "").replace(",", "").strip()
        is_income  = not amount_str.startswith("-")
        amount     = round(abs(float(amount_str)), 2)

        # Clean merchant name: split on 2+ spaces, take first part, strip trailing junk
        merchant_raw = re.split(r'\s{2,}', description)[0].strip()
        merchant_raw = re.sub(r'\s*[\d#\-]+\s*$', '', merchant_raw).strip()
        location     = merchant_raw.title() if merchant_raw else description.title()

        if is_income:
            category = "Income"
            for keyword, cat in DEBIT_INCOME_CATEGORIES.items():
                if keyword in desc_lower:
                    category = cat
                    break
            income_rows.append({
                "Cost":     amount,
                "Location": location,
                "Category": category,
                "Date":     date,
                "Account":  "Wings Debit",
            })
        else:
            category = "Others"
            for keyword, cat in DEBIT_KEYWORD_CATEGORIES.items():
                if keyword in desc_lower:
                    category = cat
                    break
            override = apply_merchant_override(description)
            if override:
                category = override
            expense_rows.append({
                "Cost":     amount,
                "Location": location,
                "Category": category,
                "Date":     date,
                "Account":  "Wings Debit",
            })

    return expense_rows, income_rows


def parse_amex(filepath):
    """Parse American Express CSV into standardized rows.
    Handles both the full export (11 columns incl. Category) and the simplified
    3-column export (Date, Description, Amount).
    """
    df = pd.read_csv(filepath)
    rows = []

    for _, row in df.iterrows():
        description = str(row.get("Description", "")).strip()
        amount_raw  = str(row.get("Amount", "")).strip()
        date        = parse_date(row.get("Date", ""))

        if not amount_raw:
            continue

        # Skip payments (negative amounts or "payment" in description)
        if "payment" in description.lower():
            continue

        amount = clean_amount(amount_raw)

        # AMEX pads merchant name to ~20 chars then appends CITY  ST
        merchant_raw = description[:20].strip()
        # Strip trailing store numbers / reference codes
        merchant_raw = re.sub(r'\s*[\d#\-]+\s*$', '', merchant_raw).strip()
        # Remove PRE-AUTH and AplPay (Apple Pay) prefixes
        merchant_raw = re.sub(r'^PRE-AUTH\s+', '', merchant_raw, flags=re.IGNORECASE).strip()
        merchant_raw = re.sub(r'^AplPay\s+',   '', merchant_raw, flags=re.IGNORECASE).strip()
        raw = merchant_raw if merchant_raw else description
        location = re.sub(r"'([A-Z])", lambda m: "'" + m.group(1).lower(), raw.title())

        # Map AMEX category if the full export format is present
        amex_cat = str(row.get("Category", "")).strip().lower()
        category = "Others"
        for key, cat in AMEX_CATEGORY_MAP.items():
            if amex_cat.startswith(key):
                category = cat
                break

        # Merchant overrides take final priority
        override = apply_merchant_override(description)
        if override:
            category = override

        rows.append({
            "Cost":     amount,
            "Location": location,
            "Category": category,
            "Date":     date,
            "Account":  "Amex",
        })

    return rows


# ──────────────────────────────────────────────
# PARSER REGISTRY
# ──────────────────────────────────────────────

PARSERS = {
    "apple_card":  parse_apple_card,
    "wings_visa":  parse_wings_visa,
    "wings_debit": parse_wings_debit,
    "amex":        parse_amex,
}

# ──────────────────────────────────────────────
# GOOGLE SHEETS IMPORT
# ──────────────────────────────────────────────

def get_existing_rows(worksheet):
    """Return a set of (date, location, cost) tuples already in the sheet."""
    existing = worksheet.get_all_values()
    seen = set()
    for row in existing[1:]:  # skip header
        if len(row) >= 4:
            try:
                # Normalize cost to float string so "150" and "150.0" compare equal
                cost_key = str(round(float(row[0].replace(",", "")), 2))
            except (ValueError, AttributeError):
                cost_key = row[0]
            seen.add((row[3], row[1], cost_key))
    return seen


def append_to_sheet(rows, sheet_name, worksheet_name, credentials_file):
    """Append new rows to the Google Sheet, skipping duplicates."""
    import gspread
    from google.oauth2.service_account import Credentials

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds     = Credentials.from_service_account_file(credentials_file, scopes=scopes)
    client    = gspread.authorize(creds)
    sheet     = client.open(sheet_name)
    worksheet = sheet.worksheet(worksheet_name)

    existing  = get_existing_rows(worksheet)
    new_rows  = []

    for r in rows:
        key = (r["Date"], r["Location"], str(round(float(r["Cost"]), 2)))
        if key not in existing:
            new_rows.append([r["Cost"], r["Location"], r["Category"], r["Date"], r["Account"]])

    if new_rows:
        worksheet.append_rows(new_rows, value_input_option="USER_ENTERED")
        print(f"✅ Appended {len(new_rows)} new rows to '{worksheet_name}'.")
    else:
        print("ℹ️  No new rows to add (all already exist in sheet).")

    return len(new_rows)


# ──────────────────────────────────────────────
# MAIN
# ──────────────────────────────────────────────

def detect_file_type(filename):
    """Infer file type from filename. Returns a key from PARSERS or None."""
    name = filename.lower()
    if "apple" in name:
        return "apple_card"
    if "visa" in name or "platinum" in name:
        return "wings_visa"
    if "debit" in name or "checking" in name:
        return "wings_debit"
    if "amex" in name:
        return "amex"
    return None


def archive_file(filepath):
    """Move a file to ARCHIVE_DIR, adding a timestamp suffix on name collision."""
    dest = ARCHIVE_DIR / filepath.name
    if dest.exists():
        stem   = filepath.stem
        suffix = filepath.suffix
        ts     = datetime.now().strftime("%Y%m%d_%H%M%S")
        dest   = ARCHIVE_DIR / f"{stem}_{ts}{suffix}"
    filepath.rename(dest)
    print(f"   📦 Archived → archive/{dest.name}")


def main():
    INBOX_DIR.mkdir(exist_ok=True)
    ARCHIVE_DIR.mkdir(exist_ok=True)

    csv_files = sorted(INBOX_DIR.glob("*.csv"))
    if not csv_files:
        print("📭 No CSV files found in inbox/ — nothing to do.")
        return

    expense_rows = []
    income_rows  = []
    processed    = []
    skipped      = []

    for filepath in csv_files:
        file_type = detect_file_type(filepath.name)
        if file_type is None:
            print(f"⚠️  Could not detect type for '{filepath.name}' — skipping")
            skipped.append(filepath)
            continue

        print(f"📂 Parsing {file_type}: {filepath.name}")
        result = PARSERS[file_type](filepath)

        # Wings Debit returns (expense_rows, income_rows); all others return a flat list
        if isinstance(result, tuple):
            expenses, income = result
            print(f"   → {len(expenses)} expense(s), {len(income)} income row(s) found")
            expense_rows.extend(expenses)
            income_rows.extend(income)
        else:
            print(f"   → {len(result)} transactions found")
            expense_rows.extend(result)

        processed.append(filepath)

    if not expense_rows and not income_rows:
        print("ℹ️  No transactions found across inbox files.")
        for fp in processed:
            archive_file(fp)
        return

    sort_key = lambda r: datetime.strptime(r["Date"], "%m/%d/%Y")
    expense_rows.sort(key=sort_key)
    income_rows.sort(key=sort_key)

    if DRY_RUN:
        pd.set_option("display.max_rows", None)
        pd.set_option("display.max_colwidth", 40)
        if expense_rows:
            print(f"\n🔍 DRY RUN — '{WORKSHEET_NAME}' rows ({len(expense_rows)}):\n")
            print(pd.DataFrame(expense_rows, columns=["Cost", "Location", "Category", "Date", "Account"]).to_string(index=False))
        if income_rows:
            print(f"\n🔍 DRY RUN — '{INCOME_WORKSHEET_NAME}' rows ({len(income_rows)}):\n")
            print(pd.DataFrame(income_rows, columns=["Cost", "Location", "Category", "Date", "Account"]).to_string(index=False))
    else:
        if not Path(CREDENTIALS_FILE).exists():
            print(f"\n❌ credentials.json not found at '{CREDENTIALS_FILE}'.")
            print("   See the GOOGLE SHEETS SETUP section in the script comments.")
            sys.exit(1)

        print(f"\n📊 Connecting to Google Sheet: '{GOOGLE_SHEET_NAME}'")
        if expense_rows:
            print(f"   → Writing to '{WORKSHEET_NAME}'")
            append_to_sheet(expense_rows, GOOGLE_SHEET_NAME, WORKSHEET_NAME, CREDENTIALS_FILE)
        if income_rows:
            print(f"   → Writing to '{INCOME_WORKSHEET_NAME}'")
            append_to_sheet(income_rows, GOOGLE_SHEET_NAME, INCOME_WORKSHEET_NAME, CREDENTIALS_FILE)

    print()
    for fp in processed:
        archive_file(fp)
    if skipped:
        print(f"\n⚠️  Skipped {len(skipped)} unrecognized file(s): {[f.name for f in skipped]}")


if __name__ == "__main__":
    main()


# ══════════════════════════════════════════════════════════════
# GOOGLE SHEETS API SETUP (one-time, ~5 minutes)
# ══════════════════════════════════════════════════════════════
#
# 1. Go to https://console.cloud.google.com/
# 2. Create a new project (or use an existing one)
# 3. Enable these two APIs:
#    - Google Sheets API
#    - Google Drive API
#
# 4. Create a Service Account:
#    IAM & Admin → Service Accounts → Create Service Account
#    Give it any name, skip optional steps, click Done
#
# 5. Create a JSON key:
#    Click the service account → Keys tab → Add Key → Create new key → JSON
#    Save the downloaded file as "credentials.json" in the same folder as this script
#
# 6. Share your Google Sheet with the service account:
#    Open the JSON file, copy the "client_email" value (looks like xxx@xxx.iam.gserviceaccount.com)
#    Open your Google Sheet → Share → paste that email → Editor → Done
#
# 7. Install dependencies:
#    pip install pandas gspread google-auth
#
# 8. Run:
#    python budget_importer.py <file.csv> <file_type>
#
# ══════════════════════════════════════════════════════════════
