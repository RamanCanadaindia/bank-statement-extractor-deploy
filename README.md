# Bank Statement Extractor Website

A Streamlit website for converting Canadian bank statements into Excel and merging monthly workbooks into an annual file.

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
