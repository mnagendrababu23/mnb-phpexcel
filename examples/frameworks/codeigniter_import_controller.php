<?php

declare(strict_types=1);

use Mnb\PHPExcel\MnbExcel;

final class ExcelImportController
{
    public function upload(): void
    {
        $file = $_FILES['excel'] ?? null;
        $upload = MnbExcel::validateUpload($file ?? [], [
            'allowed_extensions' => ['xlsx', 'csv'],
            'max_size_mb' => 100,
        ]);

        if (!$upload['valid']) {
            header('Content-Type: application/json');
            echo json_encode($upload);
            return;
        }

        $path = MnbExcel::storagePath('upload_path', basename((string) $file['name']));
        move_uploaded_file((string) $file['tmp_name'], $path);

        $result = MnbExcel::largeImportToSql($path, APPPATH . 'Config/database.php', 'students', [
            'with_header' => true,
            'chunk_size' => 500,
            'resume' => true,
            'duplicate_strategy' => 'skip',
            'failed_rows_format' => 'human',
        ]);

        header('Content-Type: application/json');
        echo json_encode($result);
    }

    public function status(string $manifestFile): void
    {
        header('Content-Type: application/json');
        echo json_encode(MnbExcel::importDashboard(MnbExcel::storagePath('manifest_path', $manifestFile)));
    }
}
