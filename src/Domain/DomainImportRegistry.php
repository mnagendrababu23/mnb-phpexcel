<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Domain;

use Mnb\PHPExcel\Support\ErrorCode;
use Mnb\PHPExcel\Support\MnbExcelException;

final class DomainImportRegistry
{
    /** @var array<string,DomainImportPreset> */
    private array $presets = [];

    public static function withBuiltIns(): self
    {
        $registry = new self();
        foreach (self::builtIns() as $preset) $registry->register($preset);
        return $registry;
    }

    public function register(DomainImportPreset $preset): self
    {
        $this->presets[$preset->type->value] = $preset;
        return $this;
    }

    public function has(DomainImportType|string $type): bool
    {
        try { $resolved = DomainImportType::fromMixed($type); } catch (MnbExcelException) { return false; }
        return isset($this->presets[$resolved->value]);
    }

    public function get(DomainImportType|string $type): DomainImportPreset
    {
        $resolved = DomainImportType::fromMixed($type);
        if (!isset($this->presets[$resolved->value])) {
            throw MnbExcelException::withCode('Domain import preset is not registered: ' . $resolved->value, ErrorCode::VALIDATION_FAILED);
        }
        return $this->presets[$resolved->value];
    }

    /** @return array<string,DomainImportPreset> */
    public function all(): array { return $this->presets; }
    /** @return list<string> */
    public function names(): array { return array_keys($this->presets); }

    /** @return list<DomainImportPreset> */
    private static function builtIns(): array
    {
        return [self::users(), self::products(), self::orders(), self::inventory(), self::students(), self::attendance(), self::marks(), self::contacts(), self::locations(), self::blogPosts(), self::media(), self::categories()];
    }

    private static function users(): DomainImportPreset
    {
        return new DomainImportPreset(DomainImportType::Users, 'users', 'Employee, customer, member, or application-user imports.', [
            'id' => self::f('ID', ['user id','employee id','customer id'], 'nullable|integer', '1001', normalizer:'integer'),
            'email' => self::f('Email', ['email address','e-mail','mail'], 'required|email', 'jane@example.com', 'Primary unique email.', normalizer:'email'),
            'username' => self::f('Username', ['login','login name','user name'], 'nullable|string|max_length:100', 'jane.doe'),
            'name' => self::f('Name', ['full name','display name','employee name','customer name'], 'nullable|string|max_length:255', 'Jane Doe'),
            'first_name' => self::f('First Name', ['firstname','given name','given_name'], 'nullable|string|max_length:120', 'Jane'),
            'last_name' => self::f('Last Name', ['lastname','surname','family name','family_name'], 'nullable|string|max_length:120', 'Doe'),
            'phone' => self::f('Phone', ['mobile','mobile number','telephone','phone number'], 'nullable|phone_basic', '+1 555 0100'),
            'role' => self::f('Role', ['user role','account role'], 'nullable|string|max_length:100', 'customer'),
            'status' => self::f('Status', ['account status','active'], 'nullable|string|max_length:50', 'active', default:'active'),
            'password' => self::f('Password', ['temporary password','temp password'], 'nullable|string|min_length:8', 'Password123'),
            'created_at' => self::f('Created At', ['created','created date','registration date'], 'nullable|date', '2026-01-15', normalizer:'datetime'),
        ], ['email']);
    }

    private static function products(): DomainImportPreset
    {
        return new DomainImportPreset(DomainImportType::Products, 'products', 'Product catalog, pricing, stock, and media-path imports.', [
            'id' => self::f('ID', ['product id'], 'nullable|integer', '501', normalizer:'integer'),
            'sku' => self::f('SKU', ['product sku','item code','product code','stock keeping unit'], 'required|string|max_length:120', 'SKU-1001', 'Unique product code.'),
            'name' => self::f('Name', ['product name','item name','title'], 'required|string|max_length:255', 'Wireless Keyboard'),
            'description' => self::f('Description', ['product description','details'], 'nullable|string', 'Compact wireless keyboard'),
            'category' => self::f('Category', ['category name','product category'], 'nullable|string|max_length:255', 'Accessories'),
            'category_id' => self::f('Category ID', ['category_id'], 'nullable|integer', '12', normalizer:'integer'),
            'price' => self::f('Price', ['unit price','regular price','selling price'], 'nullable|numeric|min:0', '49.99', normalizer:'decimal'),
            'sale_price' => self::f('Sale Price', ['discount price','offer price'], 'nullable|numeric|min:0', '39.99', normalizer:'decimal'),
            'cost' => self::f('Cost', ['cost price','purchase price'], 'nullable|numeric|min:0', '25.00', normalizer:'decimal'),
            'stock' => self::f('Stock', ['quantity','qty','stock quantity','inventory'], 'nullable|integer|min:0', '10', normalizer:'integer'),
            'barcode' => self::f('Barcode', ['upc','ean','gtin'], 'nullable|string|max_length:100', '1234567890123'),
            'image_path' => self::f('Image Path', ['image','image url','image path','product image'], 'nullable|string|max_length:2048', '/images/keyboard.jpg'),
            'status' => self::f('Status', ['product status','active'], 'nullable|string|max_length:50', 'active', default:'active'),
        ], ['sku']);
    }

    private static function orders(): DomainImportPreset
    {
        return new DomainImportPreset(DomainImportType::Orders, 'orders', 'Sales orders, customers, payment, status, and amount imports.', [
            'id' => self::f('ID', ['order id'], 'nullable|integer', '1', normalizer:'integer'),
            'order_number' => self::f('Order Number', ['order no','order #','invoice number','reference'], 'required|string|max_length:100', 'ORD-1001'),
            'customer_id' => self::f('Customer ID', ['user id','client id'], 'nullable|integer', '1001', normalizer:'integer'),
            'customer_email' => self::f('Customer Email', ['email','buyer email'], 'nullable|email', 'buyer@example.com', normalizer:'email'),
            'customer_name' => self::f('Customer Name', ['buyer name','client name'], 'nullable|string|max_length:255', 'Jane Doe'),
            'order_date' => self::f('Order Date', ['date','created at'], 'required|date', '2026-01-15', normalizer:'datetime'),
            'status' => self::f('Status', ['order status'], 'nullable|string|max_length:50', 'pending', default:'pending'),
            'payment_status' => self::f('Payment Status', ['payment','paid status'], 'nullable|string|max_length:50', 'paid'),
            'subtotal' => self::f('Subtotal', ['sub total'], 'nullable|numeric|min:0', '100.00', normalizer:'decimal'),
            'tax' => self::f('Tax', ['tax amount'], 'nullable|numeric|min:0', '8.00', normalizer:'decimal'),
            'shipping' => self::f('Shipping', ['shipping amount','delivery fee'], 'nullable|numeric|min:0', '5.00', normalizer:'decimal'),
            'total' => self::f('Total', ['grand total','order total','amount'], 'required|numeric|min:0', '113.00', normalizer:'decimal'),
            'currency' => self::f('Currency', ['currency code'], 'nullable|string|max_length:10', 'USD', default:'USD', normalizer:'uppercase'),
        ], ['order_number']);
    }

    private static function inventory(): DomainImportPreset
    {
        return new DomainImportPreset(DomainImportType::Inventory, 'inventory', 'Warehouse stock, availability, and reorder imports.', [
            'id' => self::f('ID', ['inventory id'], 'nullable|integer', '1', normalizer:'integer'),
            'sku' => self::f('SKU', ['product code','item code'], 'required|string|max_length:120', 'SKU-1001'),
            'warehouse' => self::f('Warehouse', ['warehouse code','location','store'], 'required|string|max_length:150', 'MAIN'),
            'quantity' => self::f('Quantity', ['qty','stock','on hand'], 'required|integer|min:0', '100', normalizer:'integer'),
            'reserved' => self::f('Reserved', ['reserved quantity','allocated'], 'nullable|integer|min:0', '5', default:0, normalizer:'integer'),
            'available' => self::f('Available', ['available quantity'], 'nullable|integer|min:0', '95', normalizer:'integer'),
            'reorder_level' => self::f('Reorder Level', ['minimum stock','min stock'], 'nullable|integer|min:0', '20', normalizer:'integer'),
            'bin_location' => self::f('Bin Location', ['bin','rack','shelf'], 'nullable|string|max_length:100', 'A-01'),
            'updated_at' => self::f('Updated At', ['stock date','last updated'], 'nullable|date', '2026-01-15', normalizer:'datetime'),
        ], ['sku','warehouse']);
    }

    private static function students(): DomainImportPreset
    {
        return new DomainImportPreset(DomainImportType::Students, 'students', 'Student admissions, identity, class, guardian, and contact imports.', [
            'id' => self::f('ID', ['database id'], 'nullable|integer', '1', normalizer:'integer'),
            'student_id' => self::f('Student ID', ['student number','admission number','roll number'], 'required|string|max_length:100', 'STU-1001'),
            'name' => self::f('Name', ['student name','full name'], 'nullable|string|max_length:255', 'Alex Smith'),
            'first_name' => self::f('First Name', ['given name'], 'nullable|string|max_length:120', 'Alex'),
            'last_name' => self::f('Last Name', ['surname','family name'], 'nullable|string|max_length:120', 'Smith'),
            'email' => self::f('Email', ['student email'], 'nullable|email', 'alex@example.com', normalizer:'email'),
            'phone' => self::f('Phone', ['mobile','contact number'], 'nullable|phone_basic', '+1 555 0101'),
            'date_of_birth' => self::f('Date of Birth', ['dob','birth date'], 'nullable|date', '2010-05-12', normalizer:'date'),
            'class' => self::f('Class', ['grade','standard'], 'nullable|string|max_length:100', '10'),
            'section' => self::f('Section', ['division'], 'nullable|string|max_length:50', 'A'),
            'guardian_name' => self::f('Guardian Name', ['parent name','father name','mother name'], 'nullable|string|max_length:255', 'Sam Smith'),
            'guardian_phone' => self::f('Guardian Phone', ['parent phone'], 'nullable|phone_basic', '+1 555 0102'),
            'admission_date' => self::f('Admission Date', ['joined date','enrollment date'], 'nullable|date', '2026-01-15', normalizer:'date'),
            'status' => self::f('Status', ['student status'], 'nullable|string|max_length:50', 'active', default:'active'),
        ], ['student_id']);
    }

    private static function attendance(): DomainImportPreset
    {
        $identity = static function(array $row): ?string {
            return trim((string)($row['student_id'] ?? '')) === '' && trim((string)($row['employee_id'] ?? '')) === '' ? 'Either student_id or employee_id is required.' : null;
        };
        return new DomainImportPreset(DomainImportType::Attendance, 'attendance', 'Student or employee attendance records.', [
            'student_id' => self::f('Student ID', ['student number','roll number'], 'nullable|string|max_length:100', 'STU-1001'),
            'employee_id' => self::f('Employee ID', ['staff id','user id'], 'nullable|string|max_length:100', ''),
            'date' => self::f('Date', ['attendance date'], 'required|date', '2026-01-15', normalizer:'date'),
            'status' => self::f('Status', ['attendance status','present'], 'required|string|max_length:50', 'present'),
            'check_in' => self::f('Check In', ['in time','arrival'], 'nullable|string|max_length:50', '09:00'),
            'check_out' => self::f('Check Out', ['out time','departure'], 'nullable|string|max_length:50', '17:00'),
            'remarks' => self::f('Remarks', ['notes','comment'], 'nullable|string', ''),
        ], ['student_id','employee_id','date'], [$identity]);
    }

    private static function marks(): DomainImportPreset
    {
        return new DomainImportPreset(DomainImportType::Marks, 'marks', 'Exam marks, grades, subjects, and result imports.', [
            'student_id' => self::f('Student ID', ['student number','roll number'], 'required|string|max_length:100', 'STU-1001'),
            'subject' => self::f('Subject', ['subject name','course'], 'required|string|max_length:150', 'Mathematics'),
            'exam' => self::f('Exam', ['exam name','assessment','term'], 'required|string|max_length:150', 'Midterm'),
            'marks_obtained' => self::f('Marks Obtained', ['marks','score','obtained'], 'required|numeric|min:0', '85', normalizer:'decimal'),
            'maximum_marks' => self::f('Maximum Marks', ['max marks','total marks','out of'], 'required|numeric|min:0', '100', normalizer:'decimal'),
            'grade' => self::f('Grade', ['letter grade'], 'nullable|string|max_length:20', 'A'),
            'result_date' => self::f('Result Date', ['date','exam date'], 'nullable|date', '2026-01-15', normalizer:'date'),
            'remarks' => self::f('Remarks', ['comments','notes'], 'nullable|string', ''),
        ], ['student_id','subject','exam']);
    }

    private static function contacts(): DomainImportPreset
    {
        return new DomainImportPreset(DomainImportType::Contacts, 'contacts', 'CRM contacts, organizations, addresses, and tags.', [
            'id' => self::f('ID', ['contact id'], 'nullable|integer', '1', normalizer:'integer'),
            'email' => self::f('Email', ['email address','e-mail'], 'required|email', 'contact@example.com', normalizer:'email'),
            'name' => self::f('Name', ['contact name','full name'], 'nullable|string|max_length:255', 'Taylor Lee'),
            'first_name' => self::f('First Name', ['given name'], 'nullable|string|max_length:120', 'Taylor'),
            'last_name' => self::f('Last Name', ['surname'], 'nullable|string|max_length:120', 'Lee'),
            'phone' => self::f('Phone', ['mobile','telephone'], 'nullable|phone_basic', '+1 555 0103'),
            'company' => self::f('Company', ['organization','business'], 'nullable|string|max_length:255', 'Example Inc'),
            'job_title' => self::f('Job Title', ['designation','position'], 'nullable|string|max_length:150', 'Manager'),
            'address' => self::f('Address', ['street address'], 'nullable|string', '100 Main Street'),
            'city' => self::f('City', ['town'], 'nullable|string|max_length:150', 'Austin'),
            'state' => self::f('State', ['province','region'], 'nullable|string|max_length:150', 'Texas'),
            'postal_code' => self::f('Postal Code', ['zip','zip code','postcode'], 'nullable|string|max_length:30', '78701'),
            'country' => self::f('Country', ['country code'], 'nullable|string|max_length:100', 'US'),
            'tags' => self::f('Tags', ['labels','groups'], 'nullable|string', 'customer,vip'),
            'notes' => self::f('Notes', ['comments'], 'nullable|string', ''),
        ], ['email']);
    }

    private static function locations(): DomainImportPreset
    {
        return new DomainImportPreset(DomainImportType::Locations, 'locations', 'City, branch, store, warehouse, GPS, and timezone data.', [
            'id' => self::f('ID', ['location id'], 'nullable|integer', '1', normalizer:'integer'),
            'name' => self::f('Name', ['location name','place','branch name'], 'required|string|max_length:255', 'Downtown Store'),
            'code' => self::f('Code', ['location code','branch code'], 'nullable|string|max_length:100', 'DT-01'),
            'address' => self::f('Address', ['street address'], 'nullable|string', '100 Main Street'),
            'city' => self::f('City', ['town'], 'nullable|string|max_length:150', 'Austin'),
            'state' => self::f('State', ['province','region'], 'nullable|string|max_length:150', 'Texas'),
            'postal_code' => self::f('Postal Code', ['zip','zip code','postcode'], 'nullable|string|max_length:30', '78701'),
            'country' => self::f('Country', ['country name','country code'], 'nullable|string|max_length:100', 'US'),
            'latitude' => self::f('Latitude', ['lat'], 'nullable|numeric|min:-90|max:90', '30.2672', normalizer:'decimal'),
            'longitude' => self::f('Longitude', ['lng','lon','long'], 'nullable|numeric|min:-180|max:180', '-97.7431', normalizer:'decimal'),
            'timezone' => self::f('Timezone', ['time zone','tz'], 'nullable|string|max_length:100', 'America/Chicago'),
            'status' => self::f('Status', ['location status','active'], 'nullable|string|max_length:50', 'active', default:'active'),
        ], ['code']);
    }

    private static function blogPosts(): DomainImportPreset
    {
        return new DomainImportPreset(DomainImportType::BlogPosts, 'blog_posts', 'CMS post titles, content, authors, categories, tags, and SEO fields.', [
            'id' => self::f('ID', ['post id','blog id'], 'nullable|integer', '1', normalizer:'integer'),
            'title' => self::f('Title', ['post title','blog title'], 'required|string|max_length:255', 'Product Launch Update'),
            'slug' => self::f('Slug', ['url slug','permalink'], 'nullable|string|max_length:255', 'product-launch-update', 'Generated from title when blank.', normalizer:'slug'),
            'content' => self::f('Content', ['body','post content','article'], 'required|string', 'Full post content'),
            'excerpt' => self::f('Excerpt', ['summary','short description'], 'nullable|string', 'A short summary'),
            'status' => self::f('Status', ['post status'], 'nullable|string|max_length:50', 'draft', default:'draft'),
            'author' => self::f('Author', ['author name','writer'], 'nullable|string|max_length:255', 'Jane Doe'),
            'author_id' => self::f('Author ID', ['user id','writer id'], 'nullable|integer', '1001', normalizer:'integer'),
            'published_at' => self::f('Published At', ['publish date','published date','date'], 'nullable|date', '2026-01-15', normalizer:'datetime'),
            'category' => self::f('Category', ['category name'], 'nullable|string|max_length:255', 'News'),
            'tags' => self::f('Tags', ['post tags','labels'], 'nullable|string', 'launch,news'),
            'featured_image' => self::f('Featured Image', ['image','image path','featured image path'], 'nullable|string|max_length:2048', '/images/launch.jpg'),
            'meta_title' => self::f('Meta Title', ['seo title'], 'nullable|string|max_length:255', 'Product Launch Update'),
            'meta_description' => self::f('Meta Description', ['seo description'], 'nullable|string|max_length:500', 'Latest product launch news.'),
        ], ['slug']);
    }

    private static function media(): DomainImportPreset
    {
        $pathOrUrl = static function(array $row): ?string {
            return trim((string)($row['path'] ?? '')) === '' && trim((string)($row['url'] ?? '')) === '' ? 'Either path or url is required.' : null;
        };
        return new DomainImportPreset(DomainImportType::Media, 'media', 'Image and media-library path, URL, ownership, and metadata imports.', [
            'id' => self::f('ID', ['media id','image id'], 'nullable|integer', '1', normalizer:'integer'),
            'name' => self::f('Name', ['file name','image name'], 'nullable|string|max_length:255', 'keyboard.jpg'),
            'path' => self::f('Path', ['image path','file path','media path','local path'], 'nullable|string|max_length:2048', '/uploads/keyboard.jpg'),
            'url' => self::f('URL', ['image url','media url','file url'], 'nullable|url|max_length:2048', 'https://example.com/uploads/keyboard.jpg'),
            'alt_text' => self::f('Alt Text', ['alt','alternative text'], 'nullable|string|max_length:500', 'Wireless keyboard'),
            'title' => self::f('Title', ['media title','caption'], 'nullable|string|max_length:255', 'Wireless Keyboard'),
            'mime_type' => self::f('MIME Type', ['mime','content type','file type'], 'nullable|string|max_length:150', 'image/jpeg'),
            'size' => self::f('Size', ['file size','size bytes'], 'nullable|integer|min:0', '204800', normalizer:'integer'),
            'category' => self::f('Category', ['media category','folder'], 'nullable|string|max_length:255', 'products'),
            'entity_type' => self::f('Entity Type', ['model','owner type'], 'nullable|string|max_length:150', 'product'),
            'entity_id' => self::f('Entity ID', ['owner id','record id'], 'nullable|string|max_length:150', '501'),
        ], [], [$pathOrUrl]);
    }

    private static function categories(): DomainImportPreset
    {
        return new DomainImportPreset(DomainImportType::Categories, 'categories', 'Hierarchical product, content, or application categories.', [
            'id' => self::f('ID', ['category id'], 'nullable|integer', '1', normalizer:'integer'),
            'name' => self::f('Name', ['category name','title'], 'required|string|max_length:255', 'Accessories'),
            'slug' => self::f('Slug', ['category slug','url slug'], 'nullable|string|max_length:255', 'accessories', 'Generated from name when blank.', normalizer:'slug'),
            'parent_id' => self::f('Parent ID', ['parent category id','parent_id'], 'nullable|integer', '', normalizer:'integer'),
            'parent_slug' => self::f('Parent Slug', ['parent category','parent'], 'nullable|string|max_length:255', ''),
            'description' => self::f('Description', ['category description','details'], 'nullable|string', 'Product accessories'),
            'status' => self::f('Status', ['category status','active'], 'nullable|string|max_length:50', 'active', default:'active'),
            'sort_order' => self::f('Sort Order', ['position','display order','order'], 'nullable|integer|min:0', '10', default:0, normalizer:'integer'),
            'image_path' => self::f('Image Path', ['image','category image','image url'], 'nullable|string|max_length:2048', '/images/accessories.jpg'),
        ], ['slug']);
    }

    /** @return array<string,mixed> */
    private static function f(string $header, array $aliases = [], string $rule = '', mixed $example = '', string $description = '', mixed $default = null, string $normalizer = 'string'): array
    {
        $field = compact('header','aliases','rule','example','description','normalizer');
        if (func_num_args() >= 6 && $default !== null) $field['default'] = $default;
        return $field;
    }
}
