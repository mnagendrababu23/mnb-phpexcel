<?php
declare(strict_types=1);
require __DIR__.'/bootstrap.php';
use Mnb\PHPExcel\MnbExcel;
use Mnb\PHPExcel\Cloud\CloudManager;

$src=tempnam(sys_get_temp_dir(),'mnbcloudsrc');$dst=tempnam(sys_get_temp_dir(),'mnbclouddst');@unlink($dst);
file_put_contents($src,"a,b\n1,2\n");
$m=new CloudManager();
$m=$m->account('tenant-a',['provider'=>'local','access_token'=>'not-used']);
$f=$m->upload($src,['path'=>$dst]);
if($f->provider!=='local'||$f->size!==filesize($src))throw new RuntimeException('local upload');
$copy=$dst.'.copy';$m->download($dst,$copy);
if(file_get_contents($copy)!==file_get_contents($src))throw new RuntimeException('local download');
MnbExcel::setCloudManager($m);
if(MnbExcel::cloud()->file($dst)->name!==basename($dst))throw new RuntimeException('unified cloud facade');
@unlink($src);@unlink($dst);@unlink($copy);
echo "CloudStorageEngineSmokeTest passed\n";
