<?php
declare(strict_types=1);

/**
 * Optional PhpSpreadsheet comparison.
 * Run from a project where vendor/phpoffice/phpspreadsheet is installed.
 */
$autoload=$argv[1]??dirname(__DIR__,2).'/vendor/autoload.php';
if(!is_file($autoload)){fwrite(STDERR,"PhpSpreadsheet autoload not found: $autoload\n");exit(2);}
require $autoload;
if(!class_exists(\PhpOffice\PhpSpreadsheet\IOFactory::class)){fwrite(STDERR,"PhpSpreadsheet is not installed.\n");exit(2);}

$dir=$argv[2]??(__DIR__.'/fixtures/generated');
$out=$argv[3]??(__DIR__.'/results'); if(!is_dir($out))mkdir($out,0777,true);
$rows=[];
foreach(glob(rtrim($dir,'/\\').'/*')?:[] as $file){
 if(!is_file($file)||!preg_match('/\.(xlsx|xls|csv)$/i',$file)||str_starts_with(basename($file),'corrupt_'))continue;
 gc_collect_cycles();$p0=memory_get_peak_usage(true);$t=hrtime(true);
 try{
  $reader=\PhpOffice\PhpSpreadsheet\IOFactory::createReaderForFile($file);
  $reader->setReadDataOnly(true);
  $ss=$reader->load($file);$props=$ss->getProperties();
  $data=['title'=>$props->getTitle(),'creator'=>$props->getCreator(),'sheet_count'=>$ss->getSheetCount()];
  $ok=true;$err=null;$ss->disconnectWorksheets();unset($ss);
 }catch(Throwable $e){$ok=false;$err=$e->getMessage();$data=[];}
 $rows[]=['file'=>basename($file),'engine'=>'phpspreadsheet','ok'=>$ok,'ms'=>round((hrtime(true)-$t)/1e6,3),
 'peak_bytes'=>max(0,memory_get_peak_usage(true)-$p0),'title'=>$data['title']??null,'creator'=>$data['creator']??null,
 'sheet_count'=>$data['sheet_count']??null,'error'=>$err];
}
$p="$out/phpspreadsheet-comparison-".date('Ymd-His').".json";
file_put_contents($p,json_encode($rows,JSON_PRETTY_PRINT|JSON_UNESCAPED_SLASHES));echo "$p\n";
