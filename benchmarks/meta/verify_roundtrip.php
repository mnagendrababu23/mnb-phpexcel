<?php
declare(strict_types=1);
require dirname(__DIR__,2).'/tests/bootstrap.php';
use Mnb\PHPExcel\MnbExcel;

$dir=$argv[1]??(__DIR__.'/fixtures/generated'); $results=[];
foreach(glob(rtrim($dir,'/\\').'/*.{xlsx,xls}',GLOB_BRACE)?:[] as $src){
 if(str_starts_with(basename($src),'corrupt_'))continue;
 $dst=sys_get_temp_dir().'/mnb-meta-rt-'.bin2hex(random_bytes(5)).'.'.pathinfo($src,PATHINFO_EXTENSION);
 try{
  $before=MnbExcel::meta($src)->quick()->read()->toArray();
  MnbExcel::meta($src)->update(['document'=>['title'=>'MNB Round Trip','creator'=>'MNB Benchmark']])->save($dst);
  $after=MnbExcel::meta($dst)->quick()->read()->toArray();
  $ok=(($after['document']['title']??null)==='MNB Round Trip') && (($after['document']['creator']??null)==='MNB Benchmark');
  $results[]=['file'=>basename($src),'ok'=>$ok,'before_title'=>$before['document']['title']??null,'after_title'=>$after['document']['title']??null];
 }catch(Throwable $e){$results[]=['file'=>basename($src),'ok'=>false,'error'=>$e->getMessage()];}
 @unlink($dst);
}
echo json_encode($results,JSON_PRETTY_PRINT|JSON_UNESCAPED_SLASHES)."\n";
if(array_filter($results,fn($r)=>!($r['ok']??false)))exit(1);
