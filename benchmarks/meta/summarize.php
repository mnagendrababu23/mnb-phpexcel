<?php
declare(strict_types=1);
$file=$argv[1]??null;if(!$file||!is_file($file)){fwrite(STDERR,"Usage: php summarize.php benchmark.json\n");exit(2);}
$d=json_decode(file_get_contents($file),true,512,JSON_THROW_ON_ERROR);$r=$d['results']??[];
$groups=[];
foreach($r as $x){$k=$x['ext'].'/'.$x['profile'];$groups[$k][]=$x;}
echo "# Metadata Benchmark Summary\n\n";
foreach($groups as $k=>$items){
 $ok=array_values(array_filter($items,fn($x)=>$x['ok']));
 $avg=$ok?array_sum(array_column($ok,'ms'))/count($ok):0;
 $maxPeak=$ok?max(array_column($ok,'peak_bytes')):0;
 printf("- %s: %d/%d successful, avg %.3f ms, max peak delta %.2f MB\n",$k,count($ok),count($items),$avg,$maxPeak/1048576);
}
