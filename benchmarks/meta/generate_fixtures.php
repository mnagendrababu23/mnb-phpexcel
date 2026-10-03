<?php
declare(strict_types=1);

require dirname(__DIR__,2) . '/tests/bootstrap.php';

use Mnb\PHPExcel\Core\WorkbookBuilder;

$out=__DIR__.'/fixtures/generated';
if(!is_dir($out)) mkdir($out,0777,true);

$csvCases=[
 ['comma_utf8',",","\n",false],
 ['semicolon_utf8',";","\n",false],
 ['tab_utf8',"\t","\n",false],
 ['pipe_utf8',"|","\n",false],
 ['comma_crlf',",","\r\n",false],
 ['comma_bom',",","\n",true],
];
foreach($csvCases as [$name,$d,$eol,$bom]){
 $p="$out/$name.csv"; $h=fopen($p,'wb'); if($bom) fwrite($h,"\xEF\xBB\xBF");
 fputcsv($h,['name','amount','note'],$d,'"','');
 for($i=1;$i<=50;$i++) fputcsv($h,["User $i",$i*10,$i%7===0?"line1\nline2":"ok"],$d,'"','');
 fclose($h);
}
for($n=1;$n<=70;$n++){
 $p=sprintf("%s/csv_%03d.csv",$out,$n); $h=fopen($p,'wb');
 fputcsv($h,['id','name','value'],',','"','');
 for($i=1;$i<=($n*3+10);$i++) fputcsv($h,[$i,"Name $i",$i/3],',','"','');
 fclose($h);
}

$rows=[['Name','Amount','Date'],['Alice',1200,'2026-01-01'],['Bob',2500,'2026-01-02']];
for($n=1;$n<=30;$n++){
 try { WorkbookBuilder::fromArray($rows)->save(sprintf("%s/xlsx_%03d.xlsx",$out,$n)); } catch(Throwable $e) { break; }
}
for($n=1;$n<=30;$n++){
 try { WorkbookBuilder::fromArray($rows)->save(sprintf("%s/xls_%03d.xls",$out,$n)); } catch(Throwable $e) { break; }
}

file_put_contents("$out/corrupt_truncated.xlsx","PK\x03\x04broken");
file_put_contents("$out/corrupt_fake.xls","\xD0\xCF\x11\xE0broken");
file_put_contents("$out/corrupt_binary.csv","\0\1\2,\xff\n");
echo "Generated ".count(glob("$out/*"))." fixtures.\n";
