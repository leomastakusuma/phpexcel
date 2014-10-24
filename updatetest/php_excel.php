<?php
/**
 * PHPExcel
 *
 * Copyright (C) 2006 - 2014 PHPExcel
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
 *
 * @category   PHPExcel
 * @package    PHPExcel
 * @copyright  Copyright (c) 2006 - 2014 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version    1.8.0, 2014-03-02
 */

/** Error reporting */

error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);
date_default_timezone_set('Europe/London');

$dbhost = "localhost";
$dbuser = "root";
$dbpass = "root";
$dbname = "dev_garudacos";

define('EOL',(PHP_SAPI == 'cli') ? PHP_EOL : '<br />');

/** Include PHPExcel */
require_once 'PHPExcel_1.8.0_pdf/Classes/PHPExcel.php';
//require_once dirname(__FILE__) . '/../Classes/PHPExcel.php';

function phpexcel($waktu,$bo,$dir) {

// Create new PHPExcel object
echo date('H:i:s') , " Create new PHPExcel object" , EOL;
$objPHPExcel = new PHPExcel();

// Set document properties
echo date('H:i:s') , " Set document properties" , EOL;
$objPHPExcel->getProperties()->setCreator("Maarten Balliauw")
							 ->setLastModifiedBy("Maarten Balliauw")
							 ->setTitle("PHPExcel Test Document")
							 ->setSubject("PHPExcel Test Document")
							 ->setDescription("Test document for PHPExcel, generated using PHP classes.")
							 ->setKeywords("office PHPExcel php")
							 ->setCategory("Test result file");


// Add some data
echo date('H:i:s') , " Add some data" , EOL;
//$objPHPExcel->setActiveSheetIndex(0)
//            ->setCellValue('A1', 'Hello')
//            ->setCellValue('B2', 'world!')
//            ->setCellValue('C1', 'Hello')
//            ->setCellValue('D2', 'world!');

// Miscellaneous glyphs, UTF-8
//$objPHPExcel->setActiveSheetIndex(0)
//            ->setCellValue('A4', 'file'.PATHINFO_BASENAME)
//            ->setCellValue('A5', 'éàèùâêîôûëïüÿäöüç');


//$objPHPExcel->getActiveSheet()->setCellValue('A8',"Hello\nWorld");
//$objPHPExcel->getActiveSheet()->getRowDimension(8)->setRowHeight(-1);
//$objPHPExcel->getActiveSheet()->getStyle('A8')->getAlignment()->setWrapText(true);

// $dbc = mysql_connect( $dbhost , $dbuser , $dbpass ) or die( mysql_error() );
// mysql_select_db( $dbname );
$query = mysql_query("SELECT order_id, orders_n.userid, g.bo, userbca,
         bankname, invoiceno, pnr, departuredate, paymtd, nocc, paymentcode, approvalcode,
         transdate, transtime, fareklas, tourcode, vat, yq, yi, iwjr, alltax, agtkom, netamount, 
         amount, orders_n.komdom, orders_n.komint, orders_n.pkp, bsr, descr , intdescr , addinfo , 
         validate , status, orders_n.usersts, datetkp, noref, recon, recondesc, history, recon_solve, 
         recondesc_history, logid, trip_cat, createby, createon, modifby, 
         modifon FROM orders_n, ga_webmember g 
         WHERE orders_n.userid = g.userid and  DATE_FORMAT(orders_n.transtime,'%Y-%m')='$waktu'
         and (status='tkt' or status ='tkx') and g.bo='$bo' and g.userid!=''");

$objPHPExcel->setActiveSheetIndex(0)
    ->setCellValue('A3', 'Order_id')
    ->setCellValue('B3', 'Bo')
    ->setCellValue('C3', 'userbca')
    ->setCellValue('D3', 'bankname')
    ->setCellValue('E3', 'invoiceno')
    ->setCellValue('F3', 'pnr')
    ->setCellValue('G3', 'departuredate')
    ->setCellValue('H3', 'paymtd')
    ->setCellValue('I3', 'nocc')
    ->setCellValue('J3', 'paymentcode')
    ->setCellValue('K3', 'approvalcode')
    ->setCellValue('L3', 'transdate')
    ->setCellValue('M3', 'transtime')
    ->setCellValue('N3', 'fareklas')
    ->setCellValue('O3', 'tourcode')
    ->setCellValue('P3', 'vat')
    ->setCellValue('Q3', 'yq')
    ->setCellValue('R3', 'yi')
    ->setCellValue('S3', 'iwjr')
    ->setCellValue('T3', 'alltax')
    ->setCellValue('U3', 'agtkom')
    ->setCellValue('V3', 'netamount')
    ->setCellValue('W3', 'amount')
    ->setCellValue('X3', 'komdom')
    ->setCellValue('Y3', 'komint')
    ->setCellValue('Z3', 'pkp')
    ->setCellValue('AA3', 'bsr')
    ->setCellValue('AB3', 'descr')
    ->setCellValue('AC3', 'intdescr')
    ->setCellValue('AD3', 'addinfo')
    ->setCellValue('AE3', 'validate')
    ->setCellValue('AF3', 'status')
    ->setCellValue('AG3', 'usersts')
    ->setCellValue('AH3', 'datetkp')
    ->setCellValue('AI3', 'noref')
    ->setCellValue('AJ3', 'recon')
    ->setCellValue('AK3', 'recondesc')
    ->setCellValue('AL3', 'history')
    ->setCellValue('AM3', 'recon_solve')
    ->setCellValue('AN3', 'recondesc_history')
    ->setCellValue('AO3', 'logid')
    ->setCellValue('AP3', 'trip_cat')
    ->setCellValue('AQ3', 'createby')
    ->setCellValue('AR3', 'createon')
    ->setCellValue('AS3', 'modifby')
    ->setCellValue('AT3', 'modifon');

$no=1;
$i = 4;
while($row=  mysql_fetch_array($query)){
   
    $objPHPExcel->setActiveSheetIndex(0)
        ->setCellValue('A'.$i, $no)
        ->setCellValue('B'.$i, $row[2])
        ->setCellValue('C'.$i, $row[3])
        ->setCellValue('D'.$i, $row[4])
        ->setCellValue('E'.$i, $row[5])
        ->setCellValue('F'.$i, $row[6])
        ->setCellValue('G'.$i, $row[7])
        ->setCellValue('H'.$i, $row[8])
        ->setCellValue('I'.$i, $row[9])
        ->setCellValue('J'.$i, $row[10])
        ->setCellValue('K'.$i, $row[11])
        ->setCellValue('L'.$i, $row[12])
        ->setCellValue('M'.$i, $row[13])
        ->setCellValue('N'.$i, $row[14])
        ->setCellValue('O'.$i, $row[15])
        ->setCellValue('P'.$i, $row[16])
        ->setCellValue('Q'.$i, $row[17])
        ->setCellValue('R'.$i, $row[18])
        ->setCellValue('S'.$i, $row[19])
        ->setCellValue('T'.$i, $row[20])
        ->setCellValue('U'.$i, $row[21])
        ->setCellValue('V'.$i, $row[22])
        ->setCellValue('W'.$i, $row[23])
        ->setCellValue('X'.$i, $row[24])
        ->setCellValue('Y'.$i, $row[25])
        ->setCellValue('Z'.$i, $row[26])
        ->setCellValue('AA'.$i, $row[27])
        ->setCellValue('AB'.$i, $row[28])
        ->setCellValue('AC'.$i, $row[29])
        ->setCellValue('AD'.$i, $row[30])
        ->setCellValue('AE'.$i, $row[31])
        ->setCellValue('AF'.$i, $row[32])
        ->setCellValue('AG'.$i, $row[33])
        ->setCellValue('AH'.$i, $row[34])
        ->setCellValue('AI'.$i, $row[35])
        ->setCellValue('AJ'.$i, $row[36])
        ->setCellValue('AK'.$i, $row[37])
        ->setCellValue('AL'.$i, $row[38])
        ->setCellValue('AM'.$i, $row[39])
        ->setCellValue('AN'.$i, $row[40])
        ->setCellValue('AO'.$i, $row[41])
        ->setCellValue('AP'.$i, $row[42])
        ->setCellValue('AQ'.$i, $row[43])
        ->setCellValue('AR'.$i, $row[44])
        ->setCellValue('AS'.$i, $row[45])
        ->setCellValue('AT'.$i, $row[46]);
    
         
           $no++;
           $i++;
    
}

// Rename worksheet
echo date('H:i:s') , " Rename worksheet" , EOL;
$objPHPExcel->getActiveSheet()->setTitle('Simple');


// Set active sheet index to the first sheet, so Excel opens this as the first sheet
$objPHPExcel->setActiveSheetIndex(0);


// Save Excel 95 file
echo date('H:i:s') , " Write to Excel5 format" , EOL;
$callStartTime = microtime(true);

$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
$objWriter->save(str_replace('.php', '.xls', $dir));//'/var/www/test/xls/php_excel.xls'));
$callEndTime = microtime(true);
$callTime = $callEndTime - $callStartTime;

echo date('H:i:s') , " File written to " , str_replace('.php', '.xls', pathinfo(__FILE__, PATHINFO_BASENAME)) , EOL;
echo 'Call time to write Workbook was ' , sprintf('%.4f',$callTime) , " seconds" , EOL;
// Echo memory usage
echo date('H:i:s') , ' Current memory usage: ' , (memory_get_usage(true) / 1024 / 1024) , " MB" , EOL;


// Echo memory peak usage
echo date('H:i:s') , " Peak memory usage: " , (memory_get_peak_usage(true) / 1024 / 1024) , " MB" , EOL;

// Echo done
echo date('H:i:s') , " Done writing files" , EOL;
echo 'Files have been created in ' , getcwd() , EOL;
}
