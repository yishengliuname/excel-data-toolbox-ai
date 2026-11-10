# Excel Data Toolbox AI v9.9.2

This release includes the demo and documentation changes from the v9.9.1 tag, with a corrected Windows build for a clean checkout.

## Try it without an API key

- **Order reconciliation** creates two fictional tables, flags short payouts, unmatched order IDs, and duplicate rows, and exports a workbook with evidence.
- **Sales demo** selects the fictional sales table and prefills a supported monthly and regional analysis request.
- **Windows release checks** launch the packaged application, run both demos, and reopen their generated workbooks before publication.

## Download

- Windows x64. Extract `Excel-Data-Toolbox-AI-Windows-x64.zip` and run `BiaogeKuaichuAI.exe`.
- Sample output. Open `Synthetic-Order-Reconciliation.xlsx` to inspect the actual demo result without installing the app.
- Python. The wheel and source distribution are included for developers.

The release also includes the shared business validation and task isolation changes from `main`, a stricter allowlist for web assets, and more reliable Excel export checks for numeric and boolean values.

[Demo guide](https://github.com/yishengliuname/excel-data-toolbox-ai/blob/main/docs/DEMO.md) · [Help improve the project](https://github.com/yishengliuname/excel-data-toolbox-ai/blob/main/CONTRIBUTING.md)
