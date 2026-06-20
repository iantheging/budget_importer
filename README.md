# Budget Importer

Parses transaction CSVs from multiple accounts and appends them to a Google Sheets budget tracker.

## Supported Accounts

| Account | Filename keyword |
|---|---|
| Apple Card | `apple` |
| Wings Visa / Platinum | `visa` or `platinum` |
| Wings Checking / Debit | `checking` or `debit` |
| American Express | `amex` |

## Setup

1. Install dependencies:
   ```
   pip install pandas gspread google-auth
   ```

2. Set up a Google Cloud service account with access to the Sheets and Drive APIs, download the JSON key, and save it as `credentials.json` in the project directory.

3. Share your Google Sheet with the service account's `client_email`.

4. Update the config at the top of `budget_importer.py`:
   ```python
   GOOGLE_SHEET_NAME     = "Your Sheet Name"
   WORKSHEET_NAME        = "Transactions"
   INCOME_WORKSHEET_NAME = "Income"
   ```

## Usage

1. Drop one or more CSV exports into the `inbox/` directory. Name each file so it contains the account keyword (e.g. `amex june 2026.csv`).
2. Run:
   ```
   python budget_importer.py
   ```
3. Processed files are moved to `archive/` automatically.

Expenses go to the `Transactions` tab. Income entries from the checking account go to the `Income` tab. Transfers between accounts are skipped to avoid double-counting.

Set `DRY_RUN = True` in the config to preview rows without writing to the sheet.
