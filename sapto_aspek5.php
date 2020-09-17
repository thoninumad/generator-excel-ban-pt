<?php

/*

ASPEK 5: KEUANGAN

*/


require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

/* if ($argc < 2 )
{
    exit( "Usage: sapto_aspek5.php <id prodi sesuai di database>\n" );
} */

$styleBorder = [
    'borders' => [
        'allBorders' => ['borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN, 'color' => ['argb' => '000000']]
    ],
];

$styleCenter = [
    'alignment' => [
		'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER, 
		'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER
	]
];

$styleYellow = [
	'fill' => [
        'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
        'color' => ['rgb' => 'FFFF00'],
    ]
];

$styleGreen = [
	'fill' => [
        'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
        'color' => ['rgb' => 'D6E3BC'],
    ]
];

$styleBold = [
	'font' => [
		'bold' => true,
	]
];

// $nama_prodi = $argv[1];
$nama_prodi = 62;

$serverName = "10.199.16.69";
$connectionInfo = array( "Database"=>"its-report", "UID"=>"sa", "PWD"=>"Akreditasi2019!");
$conn = sqlsrv_connect( $serverName, $connectionInfo );
if( $conn === false ) {
    die( print_r( sqlsrv_errors(), true));
}


/**
 * Tabel 4 Penggunaan Dana
 */

$sql_penggunaan_dana = "SELECT id, prodi_id, tahun, jenis_penggunaan, pengguna_dana, total FROM akreditasi.sapto_penggunaan_dana WHERE prodi_id = '".$nama_prodi."' ORDER BY jenis_penggunaan";
$stmt = sqlsrv_query( $conn, $sql_penggunaan_dana );
if( $stmt === false) {
    die( print_r( sqlsrv_errors(), true) );
}

$data_penggunaan_dana1 = [];

while( $row = sqlsrv_fetch_array( $stmt, SQLSRV_FETCH_NUMERIC) ) {
	 $data_array_penggunaan_dana[] = array(
		$row[0], $row[1], $row[2], $row[3], $row[4], $row[5]
	);
}

$data_penggunaan_dana1 = array_merge($data_penggunaan_dana1, $data_array_penggunaan_dana); 

sqlsrv_free_stmt($stmt);

$spreadsheet_penggunaan_dana1 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./blank/sapto_penggunaan_dana.xlsx');
$worksheet_penggunaan_dana1 = $spreadsheet_penggunaan_dana1->getActiveSheet();

$worksheet_penggunaan_dana1->fromArray($data_penggunaan_dana1, NULL, 'A2');

$writer_penggunaan_dana1 = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_penggunaan_dana1, 'Xlsx');
$writer_penggunaan_dana1->save('./raw/sapto_penggunaan_dana.xlsx');

$spreadsheet_penggunaan_dana1->disconnectWorksheets();
unset($spreadsheet_penggunaan_dana1);

$spreadsheet_penggunaan_dana = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_penggunaan_dana.xlsx');
$worksheet_penggunaan_dana = $spreadsheet_penggunaan_dana->getActiveSheet();

$worksheet_penggunaan_dana->insertNewColumnBefore('F', 8);

$worksheet_penggunaan_dana->getCell('F1')->setValue('TS-2 UPPS');
$worksheet_penggunaan_dana->getCell('G1')->setValue('TS-1 UPPS');
$worksheet_penggunaan_dana->getCell('H1')->setValue('TS UPPS');
$worksheet_penggunaan_dana->getCell('I1')->setValue('Rata-rata UPPS');
$worksheet_penggunaan_dana->getCell('J1')->setValue('TS-2 PS');
$worksheet_penggunaan_dana->getCell('K1')->setValue('TS-1 PS');
$worksheet_penggunaan_dana->getCell('L1')->setValue('TS PS');
$worksheet_penggunaan_dana->getCell('M1')->setValue('Rata-rata PS');

$highestRow_penggunaan_dana = $worksheet_penggunaan_dana->getHighestRow();

$penggunaan_dana_ts = intval(date("Y"));
$penggunaan_dana_ts1 = intval(date("Y", strtotime("-1 year")));
$penggunaan_dana_ts2 = intval(date("Y", strtotime("-2 year")));

$worksheet_penggunaan_dana->setAutoFilter('B1:N'.$highestRow_penggunaan_dana);
$autoFilter_penggunaan_dana = $worksheet_penggunaan_dana->getAutoFilter();
$columnFilter_penggunaan_dana = $autoFilter_penggunaan_dana->getColumn('C');
$columnFilter_penggunaan_dana->setFilterType(
    \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_FILTER
);
$columnFilter_penggunaan_dana->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        $penggunaan_dana_ts2
    );
$columnFilter_penggunaan_dana->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        $penggunaan_dana_ts1
    );
$columnFilter_penggunaan_dana->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        $penggunaan_dana_ts
    );

$autoFilter_penggunaan_dana->showHideRows();

for($row = 2;$row <= $highestRow_penggunaan_dana; $row++) {
	$worksheet_penggunaan_dana->setCellValue('F'.$row, '=IF(AND(C'.$row.'='.$penggunaan_dana_ts2.',E'.$row.'="Unit Pengelola Program Studi"),N'.$row.',0)');
	$worksheet_penggunaan_dana->getCell('F'.$row)->getStyle()->setQuotePrefix(true);
	
	$worksheet_penggunaan_dana->setCellValue('G'.$row, '=IF(AND(C'.$row.'='.$penggunaan_dana_ts1.',E'.$row.'="Unit Pengelola Program Studi"),N'.$row.',0)');
	$worksheet_penggunaan_dana->getCell('G'.$row)->getStyle()->setQuotePrefix(true);
	
	$worksheet_penggunaan_dana->setCellValue('H'.$row, '=IF(AND(C'.$row.'='.$penggunaan_dana_ts.',E'.$row.'="Unit Pengelola Program Studi"),N'.$row.',0)');
	$worksheet_penggunaan_dana->getCell('H'.$row)->getStyle()->setQuotePrefix(true);
	
	$worksheet_penggunaan_dana->setCellValue('J'.$row, '=IF(AND(C'.$row.'='.$penggunaan_dana_ts2.',E'.$row.'="Pengelola Program Studi"),N'.$row.',0)');
	$worksheet_penggunaan_dana->getCell('J'.$row)->getStyle()->setQuotePrefix(true);
	
	$worksheet_penggunaan_dana->setCellValue('K'.$row, '=IF(AND(C'.$row.'='.$penggunaan_dana_ts1.',E'.$row.'="Pengelola Program Studi"),N'.$row.',0)');
	$worksheet_penggunaan_dana->getCell('K'.$row)->getStyle()->setQuotePrefix(true);
	
	$worksheet_penggunaan_dana->setCellValue('L'.$row, '=IF(AND(C'.$row.'='.$penggunaan_dana_ts.',E'.$row.'="Pengelola Program Studi"),N'.$row.',0)');
	$worksheet_penggunaan_dana->getCell('L'.$row)->getStyle()->setQuotePrefix(true);
}

$writer_penggunaan_dana = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_penggunaan_dana, 'Xls');
$writer_penggunaan_dana->save('./formatted/sapto_penggunaan_dana (F).xls');

$spreadsheet_penggunaan_dana->disconnectWorksheets();
unset($spreadsheet_penggunaan_dana);

// Load Format Baru
$spreadsheet_penggunaan_dana2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_penggunaan_dana (F).xls');
$worksheet_penggunaan_dana2 = $spreadsheet_penggunaan_dana2->getActiveSheet();

// Formasi Array SAPTO
$array_penggunaan_dana = $worksheet_penggunaan_dana2->toArray();
$data_penggunaan_dana = [];

foreach($worksheet_penggunaan_dana2->getRowIterator() as $row_id => $row) {
    if($worksheet_penggunaan_dana2->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['jenis_penggunaan'] = $array_penggunaan_dana[$row_id-1][3];
            $item['upps_ts2'] = $array_penggunaan_dana[$row_id-1][5];
			$item['upps_ts1'] = $array_penggunaan_dana[$row_id-1][6];
			$item['upps_ts'] = $array_penggunaan_dana[$row_id-1][7];
			$item['upps_rata'] = $array_penggunaan_dana[$row_id-1][8];
			$item['ps_ts2'] = $array_penggunaan_dana[$row_id-1][9];
			$item['ps_ts1'] = $array_penggunaan_dana[$row_id-1][10];
			$item['ps_ts'] = $array_penggunaan_dana[$row_id-1][11];
			$item['ps_rata'] = $array_penggunaan_dana[$row_id-1][12];
            $data_penggunaan_dana[] = $item;
        }
    }
}

$worksheet_penggunaan_dana3 = new \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet($spreadsheet_penggunaan_dana2, 'Sheet 2');
$spreadsheet_penggunaan_dana2->addSheet($worksheet_penggunaan_dana3);

$worksheet_penggunaan_dana3 = $spreadsheet_penggunaan_dana2->getSheetByName('Sheet 2');
$worksheet_penggunaan_dana3->fromArray($data_penggunaan_dana, NULL, 'A1');

$highestRow_penggunaan_dana3 = $worksheet_penggunaan_dana3->getHighestRow();

$row_jumlah = -1;

for($group = 1;$group <= 10; $group++) {

	// $ts1_upps = 0; $ts_upps = 0; $ts2_ps = 0; $ts1_ps = 0; $ts_ps = 0;
	
	$row_jumlah += 2;
	$baris_awal = $row_jumlah;
	$ts2_upps = $worksheet_penggunaan_dana3->getCell('B'.($row_jumlah))->getValue();
	$ts1_upps = $worksheet_penggunaan_dana3->getCell('C'.($row_jumlah))->getValue();
	$ts_upps = $worksheet_penggunaan_dana3->getCell('D'.($row_jumlah))->getValue();
	$ts2_ps = $worksheet_penggunaan_dana3->getCell('F'.($row_jumlah))->getValue();
	$ts1_ps = $worksheet_penggunaan_dana3->getCell('G'.($row_jumlah))->getValue();
	$ts_ps = $worksheet_penggunaan_dana3->getCell('H'.($row_jumlah))->getValue();	
	
	for($row = $row_jumlah;$row <= ($highestRow_penggunaan_dana3+1); $row++) {
		if($worksheet_penggunaan_dana3->getCell('A'.$row)->getValue() == $worksheet_penggunaan_dana3->getCell('A'.($row+1))->getValue()) {
			$ts2_upps += $worksheet_penggunaan_dana3->getCell('B'.($row+1))->getValue();
			$ts1_upps += $worksheet_penggunaan_dana3->getCell('C'.($row+1))->getValue();
			$ts_upps += $worksheet_penggunaan_dana3->getCell('D'.($row+1))->getValue();
	
			$ts2_ps += $worksheet_penggunaan_dana3->getCell('F'.($row+1))->getValue();
			$ts1_ps += $worksheet_penggunaan_dana3->getCell('G'.($row+1))->getValue();
			$ts_ps += $worksheet_penggunaan_dana3->getCell('H'.($row+1))->getValue();
			
			$row_jumlah++;
		} else {
			break;
		}
	}
	
	$worksheet_penggunaan_dana3->insertNewRowBefore(($row_jumlah+1), 1);
	$worksheet_penggunaan_dana3->setCellValue('B'.($row_jumlah+1), $ts2_upps);
	$worksheet_penggunaan_dana3->setCellValue('C'.($row_jumlah+1), $ts1_upps);
	$worksheet_penggunaan_dana3->setCellValue('D'.($row_jumlah+1), $ts_upps);
	$worksheet_penggunaan_dana3->setCellValue('F'.($row_jumlah+1), $ts2_ps);
	$worksheet_penggunaan_dana3->setCellValue('G'.($row_jumlah+1), $ts1_ps);	
	$worksheet_penggunaan_dana3->setCellValue('H'.($row_jumlah+1), $ts_ps);
	
	${"penggunaan_dana".$group} = $worksheet_penggunaan_dana3->rangeToArray('B'.($row_jumlah+1).':I'.($row_jumlah+1), NULL, TRUE, TRUE, TRUE);
}

$writer_penggunaan_dana2 = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_penggunaan_dana2, 'Xls');
$writer_penggunaan_dana2->save('./formatted/sapto_penggunaan_dana (F).xls');

$spreadsheet_penggunaan_dana2->disconnectWorksheets();
unset($spreadsheet_penggunaan_dana2);


/**
 * Set ke Template SAPTO
 */

$spreadsheet_aps = \PhpOffice\PhpSpreadsheet\IOFactory::load('./result/sapto_aps9 (F).xlsx');


// Penggunaan Dana
$worksheet_aps20 = $spreadsheet_aps->getSheetByName('4');
$worksheet_aps20->fromArray($penggunaan_dana1, NULL, 'C7');
$worksheet_aps20->fromArray($penggunaan_dana10, NULL, 'C8');
$worksheet_aps20->fromArray($penggunaan_dana6, NULL, 'C9');
$worksheet_aps20->fromArray($penggunaan_dana7, NULL, 'C10');
$worksheet_aps20->fromArray($penggunaan_dana5, NULL, 'C11');
$worksheet_aps20->fromArray($penggunaan_dana8, NULL, 'C13');
$worksheet_aps20->fromArray($penggunaan_dana9, NULL, 'C14');
$worksheet_aps20->fromArray($penggunaan_dana4, NULL, 'C16');
$worksheet_aps20->fromArray($penggunaan_dana3, NULL, 'C17');
$worksheet_aps20->fromArray($penggunaan_dana2, NULL, 'C18');

for($row = 7; $row <= 19; $row++) {
	if($worksheet_aps20->getCell('C'.$row)->getValue() == "" && $worksheet_aps20->getCell('D'.$row)->getValue() == "" && 
	$worksheet_aps20->getCell('E'.$row)->getValue() == "") {
		$worksheet_aps20->setCellValue('F'.$row, 0);
	}
	
	if($worksheet_aps20->getCell('G'.$row)->getValue() == "" && $worksheet_aps20->getCell('H'.$row)->getValue() == "" && 
	$worksheet_aps20->getCell('I'.$row)->getValue() == "") {
		$worksheet_aps20->setCellValue('J'.$row, 0);
	}
}


$writer_aps = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_aps, 'Xlsx');
$writer_aps->save('./result/sapto_aps9 (F).xlsx');

$spreadsheet_aps->disconnectWorksheets();
unset($spreadsheet_aps);

?>