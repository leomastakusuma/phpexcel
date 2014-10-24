<?php 
$dir = "/var/www/phpexcel/reporting/GOSGA/BO/2014/";
    
include "updatetest/php_excel.php";
include "updatetest/koneksi.php";

$querybo = mysql_query("SELECT distinct bo FROM ga_webmember");
echo '<pre>';
while($bo      = mysql_fetch_array($querybo)){
  $tes[]=$bo['bo'];
}

echo '<pre>';
$year = 2014;
$yearsnow=  date('Y');
for($years=$year; $years<=$yearsnow;$years++)
{
    for($bulan=1;$bulan<=12;$bulan++)
    {
        $waktu = $years.'-'.'0'.$bulan;
        $waktu2 = $years.'0'.$bulan;
        
        foreach($tes as $row)
        {   $namafile = $dir.$row.'/'."garuda-gos-".$row.'-'.$waktu2.'.xls';
            if(!file_exists($dir.'/'.$row))
            {         
                mkdir($dir.'/'.$row.'/', 0777,true);
            }
            if(!file_exists($namafile)){
                phpexcel($waktu,$row,$namafile); 
            }
        }
         
    }

}

    




?>