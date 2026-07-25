# Modular release checklist

1. Run `composer test` in PHP 8.1, 8.2, 8.3, 8.4, and the newest supported CI runtime.
2. Include `ext-zip`, `ext-xmlreader`, `ext-pdo_sqlite`, and `ext-mbstring` in at least one full integration job.
3. Run `composer test:modules`.
4. Inspect `dist/module-manifest.json` and archive hashes.
5. Tag all split repositories with the same version.
6. Publish core before dependent format packages, then database/application/all.
7. Verify clean Composer installs for: core+CSV, core+JSON, core+XML, core+XLSX, ODS, XLS, and all.
8. Run real fixtures from Excel, LibreOffice, Google Sheets export, and WPS Office.
9. Confirm the full monolith cannot be installed alongside split modules because of `replace`/`conflict` rules.
10. Publish deprecations and compatibility boundaries before changing legacy array option behavior.
