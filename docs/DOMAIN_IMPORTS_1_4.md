# Typed Domain Imports — v1.4

MNB PHPExcel v1.4 promotes twelve common imports to discoverable, domain-specific APIs while retaining the generic streaming reader, validator, transformation pipeline, failed-row exporter, and PDO batch importer underneath.

## Package ownership

- `mnb/mnb-phpexcel-database`: domain presets, `DomainImporter`, validation, duplicate handling, and PDO execution.
- `mnb/mnb-phpexcel-application`: static `MnbExcel` convenience methods and template builder.
- `mnb/mnb-phpexcel-all`: installs the complete family.

No new split repository is required.

## Built-in domains

| Domain | Method | Default table | Default unique key |
|---|---|---|---|
| Users | `importUsers()` | `users` | `email` |
| Products | `importProducts()` | `products` | `sku` |
| Orders | `importOrders()` | `orders` | `order_number` |
| Inventory | `importInventory()` | `inventory` | `sku`, `warehouse` |
| Students | `importStudents()` | `students` | `student_id` |
| Attendance | `importAttendance()` | `attendance` | identity + date |
| Marks | `importMarks()` | `marks` | student + subject + exam |
| Contacts | `importContacts()` | `contacts` | `email` |
| Locations | `importLocations()` | `locations` | `code` |
| Blog posts | `importBlogPosts()` | `blog_posts` | `slug` |
| Media paths | `importImagesWithPaths()` / `importMedia()` | `media` | configurable |
| Categories | `importCategories()` | `categories` | `slug` |

## Preview before writing

```php
use Mnb\PHPExcel\MnbExcel;

$preview = MnbExcel::previewDomainImport('products', 'products.xlsx', [
    'sheet' => 'Products',
    'limit' => 25,
]);
```

Preview performs header mapping, normalization, validation, cross-field rules, and source duplicate detection without opening a database connection.

## Import and upsert

```php
$result = MnbExcel::importProducts('products.xlsx', $pdo, 'products', [
    'batch_size' => 1000,
    'duplicate_strategy' => 'update',
    'unique_by' => ['sku'],
    'failed_rows_csv' => __DIR__ . '/failed-products.csv',
    'progress' => static function (array $state): void {
        echo $state['rows_scanned'] . " rows scanned\n";
    },
]);
```

Database duplicate strategies are `fail`, `skip`, and `update`. Source-file duplicate policies are `error`, `skip`, and `allow`, and work across processing-batch boundaries.

## Flexible maps

Built-in aliases map common headings such as `Product Code`, `Item Code`, and `Stock Keeping Unit` to `sku`. Explicit mappings accept either direction:

```php
$result = MnbExcel::importUsers('users.csv', $pdo, 'users', [
    'map' => [
        'email' => 'Mail Address',
        'Given Name' => 'first_name',
        'last_name' => 'Family Name',
    ],
]);
```

## Rules, defaults, and transformations

```php
$result = MnbExcel::importOrders('orders.xlsx', $pdo, 'orders', [
    'rules' => ['total' => 'required|numeric|min:0'],
    'defaults' => ['currency' => 'USD', 'status' => 'pending'],
    'normalizers' => ['order_number' => 'uppercase', 'total' => 'decimal'],
    'transformers' => [
        static function (array $row): array {
            $row['external_reference'] = 'IMPORT-' . $row['order_number'];
            return $row;
        },
    ],
]);
```

Built-in normalizers: string, email/lowercase, uppercase, integer, decimal/numeric, boolean, date, datetime, slug, raw, or a custom callable.

## Cross-field behavior

- Attendance requires `student_id` or `employee_id`.
- Media requires `path` or `url`.
- Blog/category slugs are generated when blank.
- User/student/contact names can be derived from first and last name.

## Generate templates

```php
MnbExcel::domainImportTemplate('students', [
    'title' => 'Student Admission Import',
    'sample_rows' => 2,
])->save('student-import-template.xlsx');
```

The template uses the same field metadata as the importer: headers, descriptions, examples, required indicators, comments, and applicable data-validation lists.

## Service API

```php
use Mnb\PHPExcel\Import\DomainImporter;

$importer = DomainImporter::create();
$schema = $importer->schema('inventory');
$preview = $importer->preview('inventory', 'stock.csv');
$result = $importer->importInventory('stock.csv', $pdo);
```

## Result information

Results include source format/sheet, source columns, mapping details and confidence, missing required columns, canonical columns, unique keys, scanned/valid/failed/planned/inserted/skipped counts, batches, bounded errors, failed-row CSV, optional sample rows, and elapsed time.

## Boundaries

The presets do not perform application business actions such as welcome emails, inventory reservation, tax calculation, media download, or category-tree reconciliation. Tables must already exist, indexes should match `unique_by`, and the relevant PDO and reader modules must be installed.
