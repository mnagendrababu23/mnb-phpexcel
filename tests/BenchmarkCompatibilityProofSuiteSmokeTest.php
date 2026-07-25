<?php

declare(strict_types=1);

require __DIR__ . '/bootstrap.php';

use Mnb\PHPExcel\MnbExcel;

$plan = MnbExcel::benchmarkPlan(['rows' => [100000, 500000, 1000000], 'columns' => 12]);
assert($plan['status'] === 'ready');
assert(count($plan['datasets']) === 3);
assert(isset($plan['libraries']['mnb_phpexcel_large_writer']));
assert(isset($plan['libraries']['phpoffice_phpspreadsheet']));
assert(in_array('peak_memory_mb', $plan['report_columns'], true));

$requirements = MnbExcel::compatibilityFixtureRequirements();
assert($requirements['status'] === 'ready');
assert(isset($requirements['required_fixture_groups']['microsoft_excel']));
assert(in_array('manual open without repair warning', $requirements['checks'], true));

$fixtureReport = MnbExcel::verifyCompatibilityFixtures(__DIR__ . '/../docs/fixtures');
assert(in_array($fixtureReport['status'], ['passed', 'warning'], true));
assert(isset($fixtureReport['requirements']));

$db = MnbExcel::databaseIntegrationCheck(['drivers' => ['mysql', 'pgsql']]);
assert(in_array($db['status'], ['available', 'skipped'], true));
assert(isset($db['plan']['drivers']['mysql']));

$capabilities = MnbExcel::advancedWorkbookCapabilities();
assert($capabilities['status'] === 'ready');
assert(isset($capabilities['capabilities']['deep_cell_level_editing']));
assert(isset($capabilities['capabilities']['formula_calculation_engine']));

// Keep internal benchmark executable for tiny dry smoke without requiring huge runtime.
if (extension_loaded('zip')) {
    $path = sys_get_temp_dir() . DIRECTORY_SEPARATOR . 'mnb-proof-benchmark-smoke.xlsx';
    @unlink($path);
    $bench = MnbExcel::benchmarkLargeWriter(5, 3, ['path' => $path]);
    assert($bench['status'] === 'completed');
    assert($bench['rows'] === 5);
    assert(is_file($path));
    @unlink($path);
}

echo "BenchmarkCompatibilityProofSuiteSmokeTest passed\n";
