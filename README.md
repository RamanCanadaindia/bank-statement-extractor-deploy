# Bank Statement Extractor Website

A Streamlit website for converting BMO, CIBC, RBC and Tangerine statements into Excel and merging monthly workbooks into an annual file.

## Run locally

```powershell
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

## Deploy

Push this folder to a private GitHub repository, then deploy `app.py` with Streamlit Community Cloud or another Python host.

For real bank statements, review the host's privacy, storage, access-control and data-retention policies. Streamlit Community Cloud is useful for testing, but a private authenticated host is more appropriate for client financial documents.

## Supported banks

- BMO PDF via Docling, or Docling JSON
- CIBC PDF
- RBC PDF
- Tangerine PDF
- TD detection only; a sample is still required to tune extraction

## Excel output

The visible Transactions sheet contains Date, Description, Debit, Credit, Category and Calculated Balance. A hidden Extraction Data sheet preserves source values for annual merging.

