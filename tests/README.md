# Smoke tests

These are standalone smoke tests, not a PHPUnit suite — they exist because
`composer test:smoke` (via `tools/run-smoke-tests.php`) already expects to
find `tests/*SmokeTest.php` and previously found none, so `composer test`
failed out of the box on a clean checkout.

Each `*SmokeTest.php` file:

- requires `tests/bootstrap.php` (which loads the Composer autoloader),
- exercises one real, observable behavior through the public `MnbExcel`
  facade (round-tripping a file, catching a blocked formula, validating an
  array, etc.), and
- exits with a non-zero status and a message on `STDERR` if any assertion
  fails.

## Running

```bash
composer test:smoke      # just the smoke tests
composer test            # syntax check + smoke tests
php tests/ArrayToXlsxSmokeTest.php   # run a single file directly
```

## Coverage in this initial pass

- `ArrayToXlsxSmokeTest.php` — array → XLSX → array round trip, plus
  save-time integrity validation.
- `CsvSmokeTest.php` — array → CSV → array round trip, text-column
  preservation, and CSV/formula-injection scanning.
- `JsonXmlSmokeTest.php` — array → JSON/XML string and saved-file round
  trips.
- `ValidationSmokeTest.php` — `validateArray()` pass/fail behavior and
  building a failed-rows report workbook.
- `SecurityFormulaGuardSmokeTest.php` — `FormulaGuard` risk detection and
  the writer's default "safe" formula policy actually blocking a save.
- `LargeWriterSmokeTest.php` — streaming writer → streaming reader round
  trip on a small generator (a correctness smoke test, **not** a
  performance benchmark — see `tools/benchmark-large-writer.php` and
  `docs/benchmarks/README.md` for real large-scale numbers).
- `EnvironmentCheckSmokeTest.php` — `environmentCheck()`,
  `releaseReadiness()`, and `version()` sanity checks.

## What's intentionally not here yet

This pass closes the "there are no shipped tests at all" gap. It does not
yet cover: large SQL import/export, the plugin/event/transformer registries,
`.env`/DSN database config resolution, or XLSX metadata (comments,
hyperlinks, rich text). Those are natural next additions — see the main
`ROADMAP.md` for the broader test/CI plan this supports.
