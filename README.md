# LNG IFRS Exposure Builder

## Run locally

```bash
pip install -r requirements.txt
streamlit run ifrs_exposure_app.py
```

Then upload the source workbook in the Streamlit page.

## Default logic

- Reads the `Raw Data` worksheet.
- Uses `PnL Year` to build IFRS 2026 to IFRS 2031 columns.
- Net Gas includes natural gas and LNG products, including `Natural Gas & LNG` and `Unsold House Curves`, but excludes HH.
- Net Oil includes oil products.
- HH is shown separately.
- EUA and freight rows are excluded.
- `Both` ViewGroup rows are included in both Discounted and Undiscounted views.
- Discounting impact is Discounted minus Undiscounted.
- Default divisor is 1,000,000 for TBTU / MM bbl output and can be changed in the sidebar.

The Excel download contains Summary, Product Detail, Classification Audit, and Filtered Raw Data sheets.
