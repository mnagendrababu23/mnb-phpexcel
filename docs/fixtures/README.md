# Real XLSX Compatibility Fixtures

Place real files from spreadsheet applications here before publishing a compatibility report.
Fixture binaries are intentionally not included in the release ZIP.

Recommended folders:

```text
docs/fixtures/microsoft-excel/simple.xlsx
docs/fixtures/microsoft-excel/formulas.xlsx
docs/fixtures/microsoft-excel/styles.xlsx
docs/fixtures/microsoft-excel/comments-hyperlinks.xlsx
docs/fixtures/libreoffice-calc/simple.xlsx
docs/fixtures/google-sheets-export/simple.xlsx
docs/fixtures/wps-office/simple.xlsx
```

Run:

```bash
php tools/compatibility-fixture-check.php docs/fixtures
```

Each fixture should be manually opened in Microsoft Excel or LibreOffice to confirm there is no repair warning.

