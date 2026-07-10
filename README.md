# Raman Financial Services Accounting Tools

A Streamlit website for bank-statement extraction, annual transaction workbooks,
payroll records, and draft compiled financial statements.

The authenticated workspace uses a responsive dashboard, persistent tool
navigation, and separate pages for each accounting, mortgage, and real estate
workflow.

The root URL is a public home page. Selecting **Open secure tools** changes to
the protected workspace route, where the configured `APP_PASSWORD` is required
before any financial tool can be opened.

## Run locally

```powershell
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

## Deploy

Deploy the included Dockerfile on a persistent container host. Docling is too large for ordinary serverless functions such as Vercel Functions.

Set the host environment variable `APP_PASSWORD` to protect the website before processing financial documents.

For real bank statements, review the host's privacy, storage, access-control and data-retention policies. Streamlit Community Cloud is useful for testing, but a private authenticated host is more appropriate for client financial documents.

## Supported banks

- BMO PDF via Docling, or Docling JSON
- CIBC PDF
- RBC PDF
- RBC Avion Visa Business credit-card PDF
- Tangerine PDF
- TD and unfamiliar banks through a generic Docling table/OCR fallback

Known layouts use tuned parsers first. Generic results are marked for review unless opening and closing balances reconcile.

## Excel output

The visible Transactions sheet contains Date, Description, Amount, Category and Calculated Balance. Deposits are positive and withdrawals are negative. A hidden Extraction Data sheet preserves source values for validation and annual merging.

For credit cards, payments/refunds are positive and purchases/fees/interest are negative. The calculated outstanding balance uses the credit-card balance direction.

## Google Sheets connection

Completed extraction results can be sent directly to a private Google Sheet:

1. Download `google_sheets_connector.gs` from the website.
2. Create an Apps Script project and paste the connector into `Code.gs`.
3. Add the Script Property `RFS_SHARED_SECRET` with a long private value.
4. Deploy the project as a Web App that executes as the sheet owner.
5. In the website, enter the Web App `/exec` URL, Google Sheet URL, destination
   tab and the same private secret.

The connector appends transactions in one batch and maintains a hidden
`_RFS Upload Log` tab. Its batch identifier prevents the same statement from
being appended twice. Formula-like text is escaped before it is written.

## Payroll CRA PDOC helper for Google Sheets

The Payroll page includes a download for `payroll_pdoc_google_sheets.gs`.
Paste that file into a Google Sheets Apps Script project attached to your
payroll workbook.

In Google Sheets:

1. Open **Extensions > Apps Script**.
2. Paste `payroll_pdoc_google_sheets.gs` into `Code.gs`.
3. Save and reload the spreadsheet.
4. Use **Payroll PDOC > Setup PDOC sheets**.
5. Use **Payroll PDOC > Paste PDOC result text**.
6. Copy text from the CRA PDOC result PDF, paste it into the dialog, and import.
7. Select the imported PDOC row and use **Apply selected PDOC row to Payroll**.

The helper updates the matching row in the `Payroll` tab by `employee_id` and
`pay_date`. It writes the CRA PDOC CPP, CPP2, EI, federal tax, provincial tax,
total deductions, and net pay into the payroll row.

## Compiled financial statements

The Financial statements page accepts searchable T2 Schedule 100 and Schedule
125 PDFs. It extracts GIFI codes and amounts, checks that the balance sheet
balances and the income statement reconciles, and creates:

- a draft Compilation Engagement Report package in PDF format
- an editable Word version for practitioner review
- a preview of all extracted Schedule 100 and Schedule 125 rows

Generated packages are drafts. A qualified practitioner must approve the basis
of accounting, report wording, classifications, report date, and signature
before issuance.

## Mortgage qualification

The Maximum mortgage page estimates Canadian mortgage capacity under the GDS
ratio. It includes property taxes, heating, 50% of condominium fees, and the
mortgage stress-test calculation.

## Real estate investment agent

The Real estate agent page accepts a Realtor.ca map search URL or a saved
listing CSV. It ranks listings using financing costs and optional comparable
sales, rental, transit, school, and development data, then exports an Excel
workbook.
