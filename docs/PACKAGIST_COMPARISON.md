# Packagist-ready Comparison Guide

## When to use MNB PHPExcel

Use MNB PHPExcel when you need an application-ready Excel workflow:

- Small/normal XLSX reports with safety and integrity checks
- Very large XLSX import without loading the full workbook into memory
- Streaming database import with validation, failed-row CSV, and checkpoint/resume
- `.env`, PHP config file, constants, or existing PDO database connections
- Upload safety checks and admin dashboard status responses
- Plugin-level custom validators, row transformers, events, and import profiles
- Large XLSX generator/PDO cursor export and CSV ZIP fallback

## When to use another library

- Use PhpSpreadsheet for deep workbook manipulation, advanced spreadsheet formats, and formula/calculation-heavy workflows.
- Use OpenSpout when you only need simple proven streaming read/write and do not need the application import workflow.
- Use PHP_XLSXWriter when you only need lightweight XLSX writing.
- Use Laravel Excel/FastExcel when the project is Laravel-only and needs Laravel-native queues, collections, and Eloquent workflows.

## Honest positioning

MNB PHPExcel is not trying to replace every advanced spreadsheet engine. It is positioned as a framework-neutral PHP Excel application toolkit: upload → preflight → method decision → streaming import/export → validation → database → dashboard/status → recovery.

