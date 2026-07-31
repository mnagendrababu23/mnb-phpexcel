# MNB PHPExcel: Monolith vs Standalone XLSX

**Comparison date:** 31 July 2026  
**Compared repositories:**

- `mnb/mnb-phpexcel` (monolithic package)
- `mnb/mnb-phpexcel-xlsx` plus its required `mnb/mnb-phpexcel-core`

## Fix applied

The monolithic package's `src/Support/Xml/XmlReader.php` has been synchronized with the corrected standalone Core implementation.

The old monolithic adapter called `syncNative()` immediately after native `XMLReader::open()` or `XMLReader::XML()`. Native XMLReader is not positioned on a node until `read()` succeeds, so strict libxml builds—such as the reported XAMPP setup—throw:

```text
Failed to read property due to libxml error
```

The corrected implementation now:

1. resets public adapter state after `open()` and `XML()`;
2. waits for `read()` before synchronizing node properties;
3. resets state at end-of-stream;
4. reads element-only properties only on element nodes;
5. wraps synchronization failures in a contextual `RuntimeException`.

A new monolithic smoke test, `tests/XmlReaderNativeInitializationSmokeTest.php`, reproduces strict libxml behavior with a stand-in native reader and confirms the fix without requiring the XMLReader extension on the test machine.

## Verification

- Monolithic syntax check: **256 PHP files passed**.
- Monolithic smoke suite: **52/52 test files passed**.
- New strict XMLReader regression test: **passed**.
- Generated XLSX fixture read through monolithic `MnbExcel::read()`: **passed**.
- The same fixture read through standalone `Format\Xlsx::read()`: **passed**.
- Both produced the same normalized rows.

## Are the two packages functionally identical?

**No.** Their common XLSX row-reading path is largely shared, and the user's reading chain now behaves the same, but their complete public APIs and some policies differ.

### Feature matrix

| Area | Monolithic `mnb-phpexcel` | Standalone Core + XLSX | Same? |
|---|---|---|---|
| XLSX read by sheet name/index | Supported | Supported | Yes after XML fix |
| Header-row mapping | Supported | Supported | Yes |
| Streaming row iteration | Supported | Supported | Mostly yes |
| XLSX write/styles/formulas/images/comments | Supported | Supported | Broadly yes |
| Encrypted XLSX and protection | Supported | Supported | Broadly yes |
| Native XMLReader initialization | **Fixed in this archive** | Already fixed in Core `v2.0.1+` | Yes after fix |
| Lightweight file/sheet/row metadata facade | `Xlsx::fileInfo()`, `sheetsInfo()`, `sheetInfo()`, `rowCount()`, `rowCounts()` exist | Underlying `XlsxQuickInfo` exists, but these five `Format\Xlsx` methods are absent | No |
| Worksheet existence helpers | Not currently exposed | `hasSheet()`, `sheetExists()`, `sheetIfExists()` | No |
| Active worksheet helpers | Not currently exposed | `activeSheetInfo()`, `activeSheetName()`, `activeSheetIndex()`, `activeSheet()`, `sheetOrActive()` and aliases | No |
| Empty worksheet helpers | Not currently exposed | `hasRows()`, `isEmpty()`, `assertHasRows()`, `requireRows()` | No |
| Worksheet-specific exceptions | Generic errors | `SheetSelectionException` and `EmptyWorksheetException` with dedicated error codes | No |
| Optional global error renderer | Not present | `MnbExcelErrorHandler` present in Core | No |
| Formula evaluator policy | Native evaluator with optional PhpSpreadsheet compatibility fallback when available | Native evaluator only | No |
| Import-template implementation | Uses monolithic `WorkbookBuilder` | Uses standalone `XlsxImportTemplateFactory` | Same feature, different architecture |
| CSV/JSON/XML/ODS/XLS/database/jobs/HTTP/mail | Included | Not part of the XLSX package | Intentionally no |
| Installation together | Monolith replaces split packages | Standalone packages conflict with monolith | Must choose one |

## Static source comparison

### Core dependency

- Standalone Core PHP source files: **51**
- Files sharing the same relative path with the monolith: **42**
- Standalone-only Core files: **9**
- Shared-path files with different contents before/after current package evolution: **6**

Standalone-only Core capabilities include:

- reader capability interfaces;
- active worksheet support;
- sheet-existence helpers;
- empty-sheet exceptions;
- worksheet-selection exceptions;
- optional developer error rendering.

### XLSX package

- Standalone XLSX PHP source files: **39**
- Files sharing the same relative path with the monolith: **37**
- Standalone-only XLSX files: **2**
- Shared-path files with different contents: **8**

The two standalone-only XLSX implementation classes are:

- `Support/XlsxCompatibilityFixtureFactory.php`
- `Template/XlsxImportTemplateFactory.php`

## Public API differences

### Available only in the monolithic `Format\Xlsx` facade

```text
fileInfo()
sheetsInfo()
sheetInfo()
rowCount()
rowCounts()
```

The standalone XLSX repository contains `XlsxQuickInfo`, so these methods can be restored as simple delegations. Their absence is API drift rather than a missing engine.

### Available only in the latest standalone `ReadSession`

```text
hasSheet()
sheetExists()
sheetIfExists()
sheetOrActive()
selectSheetOrActive()
activeSheetInfo()
activeSheetName()
activeSheetIndex()
activeSheet()
useActiveSheet()
hasRows()
isEmpty()
assertHasRows()
requireRows()
```

These are useful but are not required for the user's existing `sheet(...)->withHeaderRow(...)->rows()` chain.

## Why the same user code differed

The public chain was equivalent, but the installed internal dependency versions were not:

```text
Monolith
  MnbExcel/Format\Xlsx
    -> monolithic ReadSession
    -> monolithic XlsxReader
    -> old monolithic Support\Xml\XmlReader

Standalone
  Format\Xlsx
    -> standalone Core ReadSession
    -> standalone XlsxReader
    -> corrected Core Support\Xml\XmlReader
```

Thus the error appeared only when native XMLReader was selected by the monolithic installation.

## Recommended synchronization work

1. Release the XMLReader fix in the monolithic package as a new corrective version; do not move an existing tag.
2. Backport the standalone Core worksheet utilities and exceptions into the monolith while retaining its all-in-one services.
3. Restore the five quick-info facade methods in standalone `Format\Xlsx` and the Application facade.
4. Generate Core/XLSX source subsets from one authoritative manifest rather than manually copying fixes.
5. Add a package-parity CI job that runs the same workbook fixtures through both installation modes and compares normalized rows, metadata, exceptions, and public method availability.
6. Resolve the broader version identity issue separately: the supplied monolith is tagged `v2.0.0` but still reports runtime version `1.7.0`, while standalone Core/XLSX report 2.0.x.

## User code after the fix

Both forms remain valid when installed separately.

```php
use Mnb\PHPExcel\Format\Xlsx;

$rows = Xlsx::read(__DIR__ . '/excels/excel1.xlsx')
    ->sheet('ALL_PARAMETERS')
    ->withHeaderRow(1)
    ->rows();
```

```php
use Mnb\PHPExcel\MnbExcel;

$rows = MnbExcel::read(__DIR__ . '/excels/excel1.xlsx')
    ->sheet('ALL_PARAMETERS')
    ->withHeaderRow(1)
    ->rows();
```

Do not require the monolithic and standalone XLSX packages in the same Composer project; their Composer metadata intentionally marks them as conflicting alternatives.
