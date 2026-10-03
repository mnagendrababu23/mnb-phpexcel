<?php
declare(strict_types=1);
require __DIR__.'/bootstrap.php';
use Mnb\PHPExcel\Cloud\{CloudManager,ResumableSession};
$src=tempnam(sys_get_temp_dir(),'mnbdeep').'.csv';file_put_contents($src,"name,amount\nA,10\nB,20\n");
$m=(new CloudManager())->account('local',['provider'=>'local']);
$copy=$src.'.remote';$m->upload($src,['path'=>$copy]);
$r=$m->open($copy);$rows=iterator_to_array($r->rows());if(count($rows)<2)throw new RuntimeException('remote rows');
$meta=$r->meta()->quick()->read();if($meta->format()!=='csv')throw new RuntimeException('remote meta');
$tmp=$r->localPath();$r->close();if(is_file($tmp))throw new RuntimeException('cleanup');
$cp=$src.'.checkpoint';$s=new ResumableSession('google-drive',$src,'https://example.invalid/session',123,999,$cp);$s->save();$x=ResumableSession::restore($cp);if(!$x||$x->offset!==123)throw new RuntimeException('session restore');$x->clear();
@unlink($src);@unlink($copy);echo "CloudDeepRuntimeSmokeTest passed\n";
