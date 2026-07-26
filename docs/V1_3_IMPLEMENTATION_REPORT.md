# MNB PHPExcel v1.3.0 Implementation Report

## Objective

Promote all capabilities previously classified as partial/basic into explicit production APIs while preserving the modular/lightweight package model and v1.2 compatibility.

## Implemented capability map

| Former partial capability | v1.3 implementation |
|---|---|
| Direct cell values | `cell`, `cells`, `rangeValues`, and format shortcuts |
| Calculated formula values | Optional `FormulaEvaluatorInterface` and PhpSpreadsheet adapter |
| Formatting reads | Complete `XlsxStyleMap` and cell/range style APIs |
| Image reads | Image inventory, bytes option, anchor metadata, safe extraction |
| Rich text | Typed rich-text/run objects and cell APIs |
| Header detection | Scored semantic detector integrated into `ReadSession` |
| Templates | Styled import-template factory with validation/comments/examples |
| Charts | Native OOXML chart/drawing generation for seven chart families |
| Freeze panes | Arbitrary row/column split and top-left cell |
| Filters | Explicit range plus five filter families |
| Conditional formatting | Native OOXML rules and differential styles |
| Pivot tables | Template preservation, workbook cache relationships, source rebinding, refresh-on-open |

## Compatibility design

- New constructor fields are appended with defaults.
- Existing reader options remain valid.
- Existing facade and split-package entry points remain available.
- PhpSpreadsheet remains optional for native XLSX users.
- Generated module packages use synchronized `^1.3` dependencies.

## Verification

Completed on July 25, 2026:

- PHP syntax validation passed for 196 files.
- All 46 smoke-test scripts passed, including the v1.3 full-capability regression suite.
- All 10 modular Composer packages were generated with synchronized `^1.3` dependencies.
- Isolated package-install combinations passed for every generated package.
- Release readiness passed all 15 checks with no warnings or failures.
- OOXML worksheet, style, chart, drawing, relationship, and pivot-cache fragments were parsed and validated.
- Real XLSX packages were generated through a temporary ZIP compatibility harness, passed package integrity validation and `unzip -t`, and opened successfully in headless LibreOffice.
- A workbook containing all seven chart families opened successfully in headless LibreOffice.
- Embedded-image generation, inventory, anchor resolution, byte retrieval, and file extraction passed an end-to-end package test.
- Namespace-prefix regression tests cover workbook relationships, cells, formulas, rich text, styles, and drawing anchors.

## Environment boundary

This container does not provide native PHP `ext-zip`, `ext-xmlreader`, PDO SQLite, or Composer. Tests that explicitly require those native extensions correctly reported skips. The release CI matrix must still execute the same suite on supported PHP versions with the native extensions installed. The temporary ZIP harness and LibreOffice checks validate generated package structure, but they do not replace native-extension CI.
