<?php
declare(strict_types=1);
require __DIR__.'/bootstrap.php';
use Mnb\PHPExcel\Cloud\Http\HttpClient;
$h=new HttpClient();
$ref=new ReflectionClass($h);
$accepted=$ref->getMethod('accepted');$accepted->setAccessible(true);
if(!$accepted->invoke($h,308,[308]))throw new RuntimeException('308 not accepted');
if(!$accepted->invoke($h,202,[]))throw new RuntimeException('202 not accepted');
$retry=$ref->getMethod('retryable');$retry->setAccessible(true);
if(!$retry->invoke($h,429)||!$retry->invoke($h,503)||$retry->invoke($h,308))throw new RuntimeException('retry classification');
$cp=sys_get_temp_dir().'/mnb-checkpoint-'.bin2hex(random_bytes(4)).'.json';
\Mnb\PHPExcel\Cloud\UploadCheckpoint::save($cp,['offset'=>123]);
if((\Mnb\PHPExcel\Cloud\UploadCheckpoint::load($cp)['offset']??null)!==123)throw new RuntimeException('checkpoint');
\Mnb\PHPExcel\Cloud\UploadCheckpoint::clear($cp);
echo "CloudRuntimeHardeningSmokeTest passed\n";
