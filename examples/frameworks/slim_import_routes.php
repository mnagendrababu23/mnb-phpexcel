<?php

declare(strict_types=1);

use Mnb\PHPExcel\MnbExcel;
use Psr\Http\Message\ResponseInterface as Response;
use Psr\Http\Message\ServerRequestInterface as Request;

// Example only: wire these closures into your Slim app.

$uploadImport = function (Request $request, Response $response): Response {
    $files = $request->getUploadedFiles();
    $excel = $files['excel'] ?? null;

    if ($excel === null) {
        $response->getBody()->write(json_encode(['status' => 'error', 'message' => 'Excel file is required.']));
        return $response->withHeader('Content-Type', 'application/json')->withStatus(422);
    }

    $target = MnbExcel::storagePath('upload_path', uniqid('upload-', true) . '.xlsx');
    $excel->moveTo($target);

    $upload = MnbExcel::validateUpload($target, ['max_size_mb' => 100]);
    if (!$upload['valid']) {
        $response->getBody()->write(json_encode($upload));
        return $response->withHeader('Content-Type', 'application/json')->withStatus(422);
    }

    $plan = MnbExcel::autoImportPlan($target, ['server' => 'shared', 'memory_limit' => '256M']);
    $result = MnbExcel::largeImportToSql($target, __DIR__ . '/../../.env', 'students', [
        'with_header' => true,
        'chunk_size' => $plan['chunk_size'],
        'resume' => true,
        'time_budget_seconds' => 25,
    ]);

    $response->getBody()->write(json_encode(['plan' => $plan, 'result' => $result]));
    return $response->withHeader('Content-Type', 'application/json');
};

$status = function (Request $request, Response $response, array $args): Response {
    $manifest = MnbExcel::storagePath('manifest_path', $args['job'] . '.json');
    $response->getBody()->write(json_encode(MnbExcel::importDashboard($manifest, [
        'download_base_url' => '/admin/imports/downloads',
    ])));
    return $response->withHeader('Content-Type', 'application/json');
};
