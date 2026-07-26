# XLSX Security — v1.7

MNB PHPExcel v1.7 supports two separate Excel security mechanisms:

1. **Password-to-open encryption** protects the complete OOXML ZIP package inside an OLE Compound File container.
2. **Workbook and worksheet protection** limits editing operations after the workbook is opened.

Protection is not encryption. Use encryption whenever workbook confidentiality matters.

## Runtime requirements

The XLSX security implementation requires:

```text
ext-openssl
ext-iconv
```

The split package declares both extensions:

```bash
composer require mnb/mnb-phpexcel-xlsx
```

## Encryption modes

### Agile

Agile encryption is the default:

```php
MnbExcel::encryptXlsx(
    'report.xlsx',
    'report-secure.xlsx',
    'S3cret!'
);
```

The implementation uses:

```text
AES-256-CBC
SHA-512
100,000 password-hash iterations by default
Encrypted verifier
Encrypted package key
HMAC package-integrity verification
```

The spin count is configurable between 1,000 and 10,000,000:

```php
MnbExcel::encryptXlsx($source, $destination, $password, [
    'mode' => 'agile',
    'spin_count' => 200000,
]);
```

### Standard compatibility mode

Use Standard Encryption for broad compatibility with older viewers:

```php
MnbExcel::encryptXlsx($source, $destination, $password, [
    'mode' => 'standard',
]);
```

Equivalent convenience option:

```php
['compatibility_mode' => true]
```

Standard mode uses AES-128-ECB and SHA-1 because that is the format required by ECMA-376 Standard Encryption. Agile mode should be preferred for new confidential documents.

## Reading encrypted XLSX files

### Typed normal reader

```php
$options = ReaderOptions::defaults()->withPassword('S3cret!');

foreach (Xlsx::read('secure.xlsx', $options)->sheet('Orders')->rows() as $row) {
    // ...
}
```

Array options remain supported:

```php
$session = Xlsx::read('secure.xlsx', [
    'password' => 'S3cret!',
]);
```

### Large reader

```php
MnbExcel::largeRead('secure.xlsx', [
    'password' => 'S3cret!',
    'max_decrypted_bytes' => 2 * 1024 * 1024 * 1024,
])
    ->withHeader()
    ->chunk(1000, $callback);
```

Or fluently:

```php
MnbExcel::largeRead('secure.xlsx')
    ->password('S3cret!')
    ->chunk(1000, $callback);
```

The decrypted temporary OOXML package is removed after the operation.

## Encrypting generated output

```php
MnbExcel::report($rows)
    ->encryptWithPassword('S3cret!', [
        'mode' => 'agile',
    ])
    ->save('secure-report.xlsx');
```

Or through the format writer:

```php
Xlsx::write($rows, 'secure-report.xlsx', [
    'with_header' => true,
    'password' => 'S3cret!',
    'encryption_mode' => 'standard',
]);
```

## Decrypting to a normal XLSX

```php
MnbExcel::decryptXlsx(
    'secure.xlsx',
    'plain.xlsx',
    'S3cret!'
);
```

## Workbook and worksheet protection

```php
MnbExcel::report($rows)
    ->protectWorkbook('EditPassword!', [
        'lock_structure' => true,
        'lock_windows' => false,
    ])
    ->protectSheet('Report', 'EditPassword!', [
        'restrictions' => [
            'formatCells' => true,
            'insertRows' => true,
            'deleteRows' => true,
            'sort' => false,
            'autoFilter' => false,
        ],
    ])
    ->save('edit-protected.xlsx');
```

Protect every worksheet and encrypt the file with one call:

```php
MnbExcel::report($rows)
    ->passwordProtectOutput('S3cret!', [
        'encrypt_file' => true,
        'protect_workbook' => true,
        'protect_sheets' => true,
        'encryption_options' => ['mode' => 'agile'],
    ])
    ->save('fully-protected.xlsx');
```

Inspect protection without returning password verifiers or salts:

```php
$metadata = MnbExcel::xlsxProtection(
    'fully-protected.xlsx',
    1,
    ['password' => 'S3cret!']
);
```

## Upload and integrity validation

Encrypted uploads are fail-closed when no password is supplied:

```php
$result = MnbExcel::validateUpload($_FILES['workbook'], [
    'password' => $submittedPassword,
    'max_decrypted_bytes' => 512 * 1024 * 1024,
    'max_uncompressed_size_mb' => 1024,
]);
```

The validator decrypts to a private temporary file, applies ZIP-bomb, unsafe-path, macro, external-link, and package checks to the inner XLSX, then removes the temporary file.

Integrity validation is also password-aware:

```php
$result = MnbExcel::validateXlsx('secure.xlsx', [
    'password' => 'S3cret!',
]);
```

## Limits and operational guidance

- Office passwords are limited to 255 Unicode code points.
- `max_source_bytes`, `max_decrypted_bytes`, and upload archive limits should be set for untrusted input.
- Agile decryption verifies HMAC integrity before returning workbook bytes.
- Standard Encryption does not provide Agile's corruption-authentication properties; this is a limitation of that legacy standard.
- Decryption currently creates a temporary plaintext OOXML package. Keep the system temporary directory private and encrypted at rest when required by your threat model.
- Do not log passwords or store them in job manifests.
