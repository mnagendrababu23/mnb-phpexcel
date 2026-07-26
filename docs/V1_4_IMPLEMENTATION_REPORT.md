# MNB PHPExcel v1.4 Implementation Report

## Objective

Promote twelve generic import examples into explicit library APIs without duplicating or replacing the stable streaming and SQL import engines.

## Architecture

- `DomainImportType`: typed built-in domain names and aliases.
- `DomainImportPreset`: canonical fields, aliases, rules, defaults, unique keys, template metadata, and cross-field validators.
- `DomainImportRegistry`: instance-scoped preset registry.
- `DomainImporter`: format-neutral row streaming, mapping, normalization, validation, duplicate handling, batching, transactions, failed-row output, and progress.
- `MnbExcel`: schemas, previews, templates, generic domain import, and one method per built-in domain.

## Supported domains

Users, products, orders, inventory, students, attendance, marks, contacts, locations, blog posts, media paths, and categories.

## Developer controls

Map and alias overrides, validation rules, defaults, normalizers, transformations, unique keys, source/database duplicate policies, strictness controls, batching, transactions, callbacks, dry-run, failed-row CSV, and bounded error/sample collection.

## Modular packaging

Domain code belongs to `mnb/mnb-phpexcel-database`; facade code belongs to `mnb/mnb-phpexcel-application`. All internal package constraints are synchronized at `^1.4`.

## Compatibility

All v1.3 APIs remain available. Presets are configurable and do not assume a framework, ORM, queue, HTTP stack, mailer, storage system, or fixed table schema.

## Verification

- PHP syntax validation: 201 files passed.
- Smoke suite: 47 scripts passed.
- Modular build: 10 packages generated and isolated combinations passed.
- Release readiness: 15 checks passed, 0 warnings, 0 failures.
- Source ZIP and every module ZIP passed archive-integrity checks.
- The v1.3-to-v1.4 patch applies cleanly to the clean v1.3 source.

Extension-dependent XLSX/ODS/XML and SQLite integration cases remain skipped in this container because `ext-zip`, `ext-xmlreader`, and PDO SQLite are unavailable; these paths should also run in full-extension CI before publication.
