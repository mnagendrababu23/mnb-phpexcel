<?php

declare(strict_types=1);

require __DIR__ . '/../../tests/bootstrap.php';

use Mnb\PHPExcel\MnbExcel;

$manifest = $_GET['manifest'] ?? __DIR__ . '/../../storage/import-manifest.json';
$dashboard = MnbExcel::importDashboard((string) $manifest, [
    'download_base_url' => '/admin/imports/downloads',
]);

header('Content-Type: text/html; charset=utf-8');
?>
<!doctype html>
<html lang="en">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>MNB PHPExcel Import Dashboard</title>
    <style>
        body { font-family: system-ui, -apple-system, Segoe UI, sans-serif; margin: 32px; color: #172033; }
        .wrap { max-width: 860px; margin: 0 auto; }
        .bar { height: 14px; background: #e7ebf3; overflow: hidden; }
        .bar span { display: block; height: 100%; background: #172033; width: <?= (float) ($dashboard['progress_percent'] ?? 0) ?>%; }
        .grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 16px; margin: 24px 0; }
        .metric { border-top: 1px solid #d7deea; padding-top: 12px; }
        .metric strong { display: block; font-size: 26px; }
        .actions a, .actions button { padding: 10px 14px; border: 1px solid #172033; background: white; color: #172033; text-decoration: none; margin-right: 8px; }
    </style>
</head>
<body>
<div class="wrap">
    <h1>Import status</h1>
    <p><?= htmlspecialchars((string) ($dashboard['message'] ?? 'No status available.')) ?></p>
    <div class="bar"><span></span></div>
    <p><?= htmlspecialchars((string) ($dashboard['progress_percent'] ?? 0)) ?>% complete</p>
    <div class="grid">
        <div class="metric"><strong><?= (int) ($dashboard['rows_scanned'] ?? 0) ?></strong>Rows scanned</div>
        <div class="metric"><strong><?= (int) ($dashboard['inserted_rows'] ?? 0) ?></strong>Inserted</div>
        <div class="metric"><strong><?= (int) ($dashboard['failed_rows'] ?? 0) ?></strong>Failed</div>
        <div class="metric"><strong><?= htmlspecialchars((string) ($dashboard['estimated_remaining_seconds'] ?? '—')) ?></strong>Seconds left</div>
    </div>
    <div class="actions">
        <?php if (!empty($dashboard['resume']['enabled'])): ?>
            <button data-manifest="<?= htmlspecialchars((string) ($dashboard['resume']['manifest_path'] ?? '')) ?>">Resume import</button>
        <?php endif; ?>
        <?php if (!empty($dashboard['failed_rows_download_url'])): ?>
            <a href="<?= htmlspecialchars((string) $dashboard['failed_rows_download_url']) ?>">Download failed rows</a>
        <?php endif; ?>
    </div>
</div>
</body>
</html>
