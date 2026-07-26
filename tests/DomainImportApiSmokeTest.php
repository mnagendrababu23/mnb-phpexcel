<?php

declare(strict_types=1);

require __DIR__ . '/bootstrap.php';

use Mnb\PHPExcel\Domain\DomainImportRegistry;
use Mnb\PHPExcel\Domain\DomainImportType;
use Mnb\PHPExcel\Import\DomainImporter;
use Mnb\PHPExcel\MnbExcel;

smoke_run('built-in domain schemas', static function (): void {
    $schemas = MnbExcel::domainImportSchemas();
    smoke_assert_equals(12, count($schemas), 'Twelve domain import presets should be registered');
    foreach (DomainImportType::cases() as $case) {
        smoke_assert_true(isset($schemas[$case->value]), 'Missing domain schema: ' . $case->value);
        smoke_assert_true(($schemas[$case->value]['columns'] ?? []) !== [], 'Domain schema must expose columns: ' . $case->value);
    }
});

smoke_run('product alias mapping and cross-batch duplicates', static function (): void {
    $dir = smoke_temp_dir('domain_products');
    $path = $dir . '/products.csv';
    file_put_contents($path, "Product Code,Product Name,Selling Price,Quantity,Image URL\nSKU-1,Keyboard,$49.99,10,/images/keyboard.jpg\nSKU-1,Duplicate,20,2,/images/duplicate.jpg\n");
    $result = MnbExcel::previewDomainImport('products', $path, ['format'=>'csv','header_row'=>true,'batch_size'=>1,'limit'=>10]);
    smoke_assert_equals('dry_run_with_errors', $result['status'], 'Duplicate product preview should report errors');
    smoke_assert_equals('sku', $result['mapping']['product_code'] ?? null, 'Product Code should map to sku');
    smoke_assert_equals('price', $result['mapping']['selling_price'] ?? null, 'Selling Price should map to price');
    smoke_assert_equals(1, $result['valid_rows'], 'One product should remain valid');
    smoke_assert_equals(1, $result['failed_rows'], 'Duplicate should fail across separate batches');
    smoke_assert_equals(49.99, $result['sample_rows'][0]['data']['price'] ?? null, 'Currency-like price should normalize to decimal');
});

smoke_run('user normalizers and derived name', static function (): void {
    $dir = smoke_temp_dir('domain_users');
    $path = $dir . '/users.csv';
    file_put_contents($path, "E-mail,First Name,Last Name,Mobile Number\n JANE@EXAMPLE.COM ,Jane,Doe,+1 555 0100\n");
    $result = DomainImporter::create()->preview('users', $path, ['format'=>'csv','header_row'=>true,'limit'=>5]);
    $row = $result['sample_rows'][0]['data'] ?? [];
    smoke_assert_equals('jane@example.com', $row['email'] ?? null, 'Email should normalize to lowercase');
    smoke_assert_equals('Jane Doe', $row['name'] ?? null, 'Name should derive from first and last name');
    smoke_assert_equals('active', $row['status'] ?? null, 'User default status should be applied');
});

smoke_run('attendance and media cross-field validation', static function (): void {
    $dir = smoke_temp_dir('domain_cross_fields');
    $attendance = $dir . '/attendance.csv';
    file_put_contents($attendance, "Date,Status\n2026-01-15,present\n");
    $a = MnbExcel::previewDomainImport('attendance', $attendance, ['format'=>'csv','header_row'=>true,'strict_mapping'=>false]);
    smoke_assert_equals(1, $a['failed_rows'], 'Attendance should require a student or employee ID');
    smoke_assert_contains('student_id or employee_id', implode(' ', $a['errors'][0]['errors'] ?? []), 'Attendance error should be clear');
    $media = $dir . '/media.csv';
    file_put_contents($media, "Name,Alt Text\nphoto.jpg,Example photo\n");
    $m = MnbExcel::previewDomainImport('media', $media, ['format'=>'csv','header_row'=>true,'strict_mapping'=>false]);
    smoke_assert_equals(1, $m['failed_rows'], 'Media should require path or URL');
    smoke_assert_contains('path or url', strtolower(implode(' ', $m['errors'][0]['errors'] ?? [])), 'Media error should be clear');
});

smoke_run('blog and category slug generation', static function (): void {
    $dir = smoke_temp_dir('domain_slugs');
    $blog = $dir . '/blog.csv';
    file_put_contents($blog, "Post Title,Post Content\nHello World,Body text\n");
    $b = MnbExcel::previewDomainImport('blog_posts', $blog, ['format'=>'csv','header_row'=>true]);
    smoke_assert_equals('hello-world', $b['sample_rows'][0]['data']['slug'] ?? null, 'Blog slug should derive from title');
    $category = $dir . '/categories.csv';
    file_put_contents($category, "Category Name\nHome & Garden\n");
    $c = MnbExcel::previewDomainImport('categories', $category, ['format'=>'csv','header_row'=>true]);
    smoke_assert_equals('home-garden', $c['sample_rows'][0]['data']['slug'] ?? null, 'Category slug should derive from name');
});

smoke_run('every built-in template row validates', static function (): void {
    $registry = DomainImportRegistry::withBuiltIns();
    $dir = smoke_temp_dir('domain_templates');
    foreach ($registry->all() as $name => $preset) {
        $path = $dir . '/' . $name . '.csv';
        $h = fopen($path, 'wb');
        if ($h === false) throw new SmokeTestFailure('Unable to create fixture');
        $headers=[]; $examples=[];
        foreach ($preset->templateColumns() as $column) { $headers[]=(string)($column['header']??''); $examples[]=$column['example']??''; }
        fputcsv($h,$headers); fputcsv($h,$examples); fclose($h);
        $result = DomainImporter::create()->preview($name,$path,['format'=>'csv','header_row'=>true,'strict_mapping'=>false]);
        smoke_assert_equals(1,$result['valid_rows'],'Built-in example should validate for '.$name);
        smoke_assert_equals(0,$result['failed_rows'],'Built-in example should have no errors for '.$name);
    }
});

smoke_run('explicit maps accept both directions', static function (): void {
    $dir=smoke_temp_dir('domain_map_direction'); $path=$dir.'/users.csv';
    file_put_contents($path,"Mail Address,Given,Family\nPERSON@EXAMPLE.COM,Pat,Lee\n");
    $result=MnbExcel::previewDomainImport('users',$path,['format'=>'csv','header_row'=>true,'map'=>['email'=>'Mail Address','Given'=>'first_name','last_name'=>'Family']]);
    $row=$result['sample_rows'][0]['data']??[];
    smoke_assert_equals('person@example.com',$row['email']??null,'Canonical-to-source map should work');
    smoke_assert_equals('Pat Lee',$row['name']??null,'Mixed map directions should work');
});

smoke_run('duplicate skip policy reports skipped rows', static function (): void {
    $dir=smoke_temp_dir('domain_duplicate_skip'); $path=$dir.'/products.csv';
    file_put_contents($path,"SKU,Name\nA-1,First\nA-1,Second\n");
    $result=MnbExcel::previewDomainImport('products',$path,['format'=>'csv','header_row'=>true,'batch_size'=>1,'file_duplicate_policy'=>'skip']);
    smoke_assert_equals(1,$result['valid_rows'],'Duplicate skip should keep the first row');
    smoke_assert_equals(0,$result['failed_rows'],'Skipped duplicates should not fail');
    smoke_assert_equals(1,$result['skipped_duplicate_rows'],'Skipped duplicate count should be reported');
});

smoke_run('domain import template', static function (): void {
    $template=MnbExcel::domainImportTemplate('products')->toArray();
    smoke_assert_true(count($template)>=2,'Template should contain header and example rows');
    $found=false;
    foreach ($template as $row) if (is_array($row) && in_array('SKU',$row,true) && in_array('Name',$row,true)) $found=true;
    smoke_assert_true($found,'Product template should include SKU and Name');
});

smoke_run('domain-specific facade methods exist', static function (): void {
    foreach (['importUsers','importProducts','importOrders','importInventory','importStudents','importAttendance','importMarks','importContacts','importLocations','importBlogPosts','importImagesWithPaths','importMedia','importCategories'] as $method) {
        smoke_assert_true(method_exists(MnbExcel::class,$method),'Missing facade method: '.$method);
        smoke_assert_true(method_exists(DomainImporter::class,$method),'Missing service method: '.$method);
    }
});

echo "DomainImportApiSmokeTest passed\n";
