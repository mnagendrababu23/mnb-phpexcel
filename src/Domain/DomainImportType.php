<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Domain;

use Mnb\PHPExcel\Support\ErrorCode;
use Mnb\PHPExcel\Support\MnbExcelException;

enum DomainImportType: string
{
    case Users = 'users';
    case Products = 'products';
    case Orders = 'orders';
    case Inventory = 'inventory';
    case Students = 'students';
    case Attendance = 'attendance';
    case Marks = 'marks';
    case Contacts = 'contacts';
    case Locations = 'locations';
    case BlogPosts = 'blog_posts';
    case Media = 'media';
    case Categories = 'categories';

    public static function fromMixed(self|string $type): self
    {
        if ($type instanceof self) {
            return $type;
        }
        $normalized = strtolower(trim($type));
        $normalized = str_replace(['-', ' '], '_', $normalized);
        $aliases = [
            'user' => self::Users,
            'product' => self::Products,
            'order' => self::Orders,
            'stock' => self::Inventory,
            'inventories' => self::Inventory,
            'student' => self::Students,
            'attendances' => self::Attendance,
            'mark' => self::Marks,
            'results' => self::Marks,
            'contact' => self::Contacts,
            'location' => self::Locations,
            'blog' => self::BlogPosts,
            'post' => self::BlogPosts,
            'posts' => self::BlogPosts,
            'blog_post' => self::BlogPosts,
            'image' => self::Media,
            'images' => self::Media,
            'image_paths' => self::Media,
            'media_paths' => self::Media,
            'category' => self::Categories,
        ];
        if (isset($aliases[$normalized])) {
            return $aliases[$normalized];
        }
        $resolved = self::tryFrom($normalized);
        if ($resolved === null) {
            throw MnbExcelException::withCode(
                'Unsupported domain import type: ' . $type,
                ErrorCode::VALIDATION_FAILED,
                ['domain' => $type, 'supported' => array_map(static fn(self $case): string => $case->value, self::cases())]
            );
        }
        return $resolved;
    }
}
