<?php
declare(strict_types=1);
require dirname(__DIR__,2) . '/tests/bootstrap.php';

use Mnb\PHPExcel\MnbExcel;

$fixtureDir=$argv[1]??(__DIR__.'/fixtures/generated');
$out=$argv[2]??(__DIR__.'/results');
if(!is_dir($out)) mkdir($out,0777,true);
$files=array_values(array_filter(glob(rtrim($fixtureDir,'/\\').'/*')?:[], 'is_file'));
$rows=[];

function bench(callable $fn): array {
 gc_collect_cycles(); $m0=memory_get_usage(true); $p0=memory_get_peak_usage(true); $t=hrtime(true);
 try { $v=$fn(); $ok=true; $err=null; } catch(Throwable $e){$v=null;$ok=false;$err=get_class($e).': '.$e->getMessage();}
 return ['ok'=>$ok,'ms'=>(hrtime(true)-$t)/1e6,'memory_delta'=>max(0,memory_get_usage(true)-$m0),
   'peak_delta'=>max(0,memory_get_peak_usage(true)-$p0),'error'=>$err,'value'=>$v];
}
foreach($files as $file){
 $ext=strtolower(pathinfo($file,PATHINFO_EXTENSION));
 $r=bench(fn()=>MnbExcel::meta($file)->quick()->read()->toArray());
 $rows[]=['file'=>basename($file),'ext'=>$ext,'size'=>filesize($file),'engine'=>'mnb','profile'=>'quick',
   'ok'=>$r['ok'],'ms'=>round($r['ms'],3),'peak_bytes'=>$r['peak_delta'],'error'=>$r['error'],
   'format'=>$r['ok']?($r['value']['format']??null):null,'sheet_count'=>$r['ok']?($r['value']['workbook']['sheet_count']??null):null];
 if($r['ok'] && !str_starts_with(basename($file),'corrupt_')){
   $f=bench(fn()=>MnbExcel::meta($file)->full()->read()->toArray());
   $rows[]=['file'=>basename($file),'ext'=>$ext,'size'=>filesize($file),'engine'=>'mnb','profile'=>'full',
    'ok'=>$f['ok'],'ms'=>round($f['ms'],3),'peak_bytes'=>$f['peak_delta'],'error'=>$f['error'],
    'format'=>$f['ok']?($f['value']['format']??null):null,'sheet_count'=>$f['ok']?($f['value']['workbook']['sheet_count']??null):null];
 }
}
$stamp=date('Ymd-His'); $json="$out/meta-benchmark-$stamp.json"; $csv="$out/meta-benchmark-$stamp.csv";
file_put_contents($json,json_encode(['generated_at'=>date(DATE_ATOM),'fixture_dir'=>$fixtureDir,'results'=>$rows],JSON_PRETTY_PRINT|JSON_UNESCAPED_SLASHES));
$h=fopen($csv,'wb'); fputcsv($h,array_keys($rows[0]??['file'=>'']));
foreach($rows as $row) fputcsv($h,$row); fclose($h);
echo "$json\n$csv\n";
