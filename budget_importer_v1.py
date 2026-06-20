"""
Budget CSV Importer
====================
Standardizes transactions from Wings Visa Platinum and Apple Card CSVs
and appends them to a Google Sheets budget tracker.

Columns in master sheet: Cost | Location | Category | Date | Account

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

GOOGLE_SHEET_NAME = "Budget 2026"          # Exact name of your Google Sheet
WORKSHEET_NAME    = "Transactions"             # Tab name inside the sheet
CREDENTIALS_FILE  = "credentials.json"  # Path to your Google service account JSON

# Set to True to do a dry run (prints rows, does not write to sheet)
DRY_RUN = False

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
    "caribou coffee":     "Dining",
    "meshuggah bagels":   "Dining",   # bagel cafe, not grocery

    # Grocery clarifications
    "hy-vee":             "Grocery",  # Hy-Vee grocery stores (not gas)
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


def parse_visa(filepath):
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


# ──────────────────────────────────────────────
# GOOGLE SHEETS IMPORT
# ──────────────────────────────────────────────

def get_existing_rows(worksheet):
    """Return a set of (date, location, cost) tuples already in the sheet."""
    existing = worksheet.get_all_values()
    seen = set()
    for row in existing[1:]:  # skip header
        if len(row) >= 3:
            seen.add((row[3], row[1], row[0]))  # date, location, cost
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
        key = (r["Date"], r["Location"], str(r["Cost"]))
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

def main():
    # Accept file paths as CLI args, or use defaults
    apple_file = sys.argv[1] if len(sys.argv) > 1 else None
    visa_file  = sys.argv[2] if len(sys.argv) > 2 else None

    if not apple_file or not visa_file:
        print("Usage: python budget_importer.py <apple_card.csv> <visa.csv>")
        print("\nExample:")
        print("  python budget_importer.py Apple_Card_Jan.csv Visa_Jan.csv")
        sys.exit(1)

    print(f"📂 Parsing Apple Card: {apple_file}")
    apple_rows = parse_apple_card(apple_file)
    print(f"   → {len(apple_rows)} transactions found")

    print(f"📂 Parsing Visa: {visa_file}")
    visa_rows = parse_visa(visa_file)
    print(f"   → {len(visa_rows)} transactions found")

    all_rows = apple_rows + visa_rows
    all_rows.sort(key=lambda r: datetime.strptime(r["Date"], "%m/%d/%Y"))

    if DRY_RUN:
        print("\n🔍 DRY RUN — rows that would be written:\n")
        df = pd.DataFrame(all_rows, columns=["Cost", "Location", "Category", "Date", "Account"])
        pd.set_option("display.max_rows", None)
        pd.set_option("display.max_colwidth", 40)
        print(df.to_string(index=False))
        print(f"\nTotal: {len(all_rows)} rows")
        return

    if not Path(CREDENTIALS_FILE).exists():
        print(f"\n❌ credentials.json not found at '{CREDENTIALS_FILE}'.")
        print("   See the GOOGLE SHEETS SETUP section in the script comments.")
        sys.exit(1)

    print(f"\n📊 Connecting to Google Sheet: '{GOOGLE_SHEET_NAME}' → '{WORKSHEET_NAME}'")
    append_to_sheet(all_rows, GOOGLE_SHEET_NAME, WORKSHEET_NAME, CREDENTIALS_FILE)


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
#    python budget_importer.py Apple_Card_Jan.csv Visa_Jan.csv
#
# ══════════════════════════════════════════════════════════════
