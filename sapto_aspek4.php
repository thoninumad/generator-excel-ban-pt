<?php

/*

ASPEK 4: KINERJA DOSEN

*/

require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

/* if ($argc < 2 )
{
    exit( "Usage: sapto_aspek4.php <id prodi sesuai di database>\n" );
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
$nama_prodi = 15;

$serverName = "10.199.16.69";
$connectionInfo = array( "Database"=>"its-report", "UID"=>"sa", "PWD"=>"Akreditasi2019!");
$conn = sqlsrv_connect( $serverName, $connectionInfo );
if( $conn === false ) {
    die( print_r( sqlsrv_errors(), true));
}


/**
 * Tabel 3.b.1 Pengakuan/Rekognisi Dosen
 */

$sql_rekognisi = "SELECT id, nama_dosen, bidang_keahlian, nama_reward, tingkat_penghargaan, tahun, id_prodi FROM akreditasi.sapto_rekognisi_dosen WHERE id_prodi = '".$nama_prodi."'";
$stmt = sqlsrv_query( $conn, $sql_rekognisi );
if( $stmt === false) {
    die( print_r( sqlsrv_errors(), true) );
}

$data_rekognisi1 = [];

while( $row = sqlsrv_fetch_array( $stmt, SQLSRV_FETCH_NUMERIC) ) {
	 $data_array_rekognisi[] = array(
		$row[0], $row[1], $row[2], $row[3], $row[4], $row[5], $row[6]
	);
}

$data_rekognisi1 = array_merge($data_rekognisi1, $data_array_rekognisi); 

sqlsrv_free_stmt( $stmt);

$spreadsheet_rekognisi1 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./blank/sapto_rekognisi_dosen.xlsx');
$worksheet_rekognisi1 = $spreadsheet_rekognisi1->getActiveSheet();

$worksheet_rekognisi1->fromArray($data_rekognisi1, NULL, 'A2');

$writer_rekognisi1 = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_rekognisi1, 'Xlsx');
$writer_rekognisi1->save('./raw/sapto_rekognisi_dosen.xlsx');

$spreadsheet_rekognisi1->disconnectWorksheets();
unset($spreadsheet_rekognisi1);

$spreadsheet_rekognisi = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_rekognisi_dosen.xlsx');

$worksheet_rekognisi = $spreadsheet_rekognisi->getActiveSheet();;

$worksheet_rekognisi->insertNewColumnBefore('F', 3);

$worksheet_rekognisi->getCell('F1')->setValue('Wilayah');
$worksheet_rekognisi->getCell('G1')->setValue('Nasional');
$worksheet_rekognisi->getCell('H1')->setValue('Internasional');

$highestRow_rekognisi = $worksheet_rekognisi->getHighestRow();

for($row = 2;$row <= $highestRow_rekognisi; $row++) {
	$worksheet_rekognisi->setCellValue('F'.$row, '=IF(E'.$row.'="Lokal";"V";"")');
	$worksheet_rekognisi->getCell('F'.$row)->getStyle()->setQuotePrefix(true);
	
	$worksheet_rekognisi->setCellValue('G'.$row, '=IF(E'.$row.'="Nasional";"V";"")');
	$worksheet_rekognisi->getCell('G'.$row)->getStyle()->setQuotePrefix(true);
	
	$worksheet_rekognisi->setCellValue('H'.$row, '=IF(E'.$row.'="Internasional";"V";"")');
	$worksheet_rekognisi->getCell('H'.$row)->getStyle()->setQuotePrefix(true);
}

$writer_rekognisi = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_rekognisi, 'Xls');
$writer_rekognisi->save('./formatted/sapto_rekognisi_dosen (F).xls');

$spreadsheet_rekognisi->disconnectWorksheets();
unset($spreadsheet_rekognisi);

// Load Format Baru
$spreadsheet_rekognisi2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_rekognisi_dosen (F).xls');
$worksheet_rekognisi2 = $spreadsheet_rekognisi2->getActiveSheet();

// Formasi Array SAPTO
$array_rekognisi = $worksheet_rekognisi2->toArray();
$data_rekognisi = [];

foreach($worksheet_rekognisi2->getRowIterator() as $row_id => $row) {
    if($worksheet_rekognisi2->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['nama_dosen'] = $array_rekognisi[$row_id-1][1];
            $item['bidang_keahlian'] = $array_rekognisi[$row_id-1][2];
			$item['rekognisi'] = $array_rekognisi[$row_id-1][3];
			$item['wilayah'] = $array_rekognisi[$row_id-1][5];
			$item['nasional'] = $array_rekognisi[$row_id-1][6];
			$item['internasional'] = $array_rekognisi[$row_id-1][7];
			$item['tahun'] = $array_rekognisi[$row_id-1][8];
            $data_rekognisi[] = $item;
        }
    }
}

$spreadsheet_rekognisi2->disconnectWorksheets();
unset($spreadsheet_rekognisi2);


/**
 * Tabel 3.b.2 Penelitian DTPS
 */

$sql_penelitian_dtps = "SELECT id, sumber_dana, jml_penelitian, tahun, id_prodi FROM akreditasi.sapto_penelitian_dtps WHERE id_prodi = '".$nama_prodi."' ORDER BY sumber_dana";
$stmt = sqlsrv_query( $conn, $sql_penelitian_dtps );
if( $stmt === false) {
    die( print_r( sqlsrv_errors(), true) );
}

$data_penelitian_dtps1 = [];

while( $row = sqlsrv_fetch_array( $stmt, SQLSRV_FETCH_NUMERIC) ) {
	 $data_array_penelitian_dtps[] = array(
		$row[0], $row[1], $row[2], $row[3], $row[4]
	);
}

$data_penelitian_dtps1 = array_merge($data_penelitian_dtps1, $data_array_penelitian_dtps); 

sqlsrv_free_stmt( $stmt);

$spreadsheet_penelitian_dtps1 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./blank/sapto_penelitian_dtps.xlsx');
$worksheet_penelitian_dtps1 = $spreadsheet_penelitian_dtps1->getActiveSheet();

$worksheet_penelitian_dtps1->fromArray($data_penelitian_dtps1, NULL, 'A2');

$writer_penelitian_dtps1 = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_penelitian_dtps1, 'Xlsx');
$writer_penelitian_dtps1->save('./raw/sapto_penelitian_dtps.xlsx');

$spreadsheet_penelitian_dtps1->disconnectWorksheets();
unset($spreadsheet_penelitian_dtps1);

$spreadsheet_penelitian_dtps = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_penelitian_dtps.xlsx');
$worksheet_penelitian_dtps = $spreadsheet_penelitian_dtps->getActiveSheet();

$worksheet_penelitian_dtps->insertNewColumnBefore('D', 4);

$worksheet_penelitian_dtps->getCell('D1')->setValue('Jumlah TS-2');
$worksheet_penelitian_dtps->getCell('E1')->setValue('Jumlah TS-1');
$worksheet_penelitian_dtps->getCell('F1')->setValue('Jumlah TS');
$worksheet_penelitian_dtps->getCell('G1')->setValue('Jumlah');

$highestRow_penelitian_dtps = $worksheet_penelitian_dtps->getHighestRow();

$penelitian_dtps_ts = intval(date("Y"));
$penelitian_dtps_ts1 = intval(date("Y", strtotime("-1 year")));
$penelitian_dtps_ts2 = intval(date("Y", strtotime("-2 year")));

$worksheet_penelitian_dtps->setAutoFilter('B1:H'.$highestRow_penelitian_dtps);
$autoFilter_penelitian_dtps = $worksheet_penelitian_dtps->getAutoFilter();
$columnFilter_penelitian_dtps = $autoFilter_penelitian_dtps->getColumn('G');
$columnFilter_penelitian_dtps->setFilterType(
    \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_FILTER
);
$columnFilter_penelitian_dtps->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        $penelitian_dtps_ts2
    );
$columnFilter_penelitian_dtps->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        $penelitian_dtps_ts1
    );
$columnFilter_penelitian_dtps->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        $penelitian_dtps_ts
    );

$autoFilter_penelitian_dtps->showHideRows();

for($row = 2;$row <= $highestRow_penelitian_dtps; $row++) {
	$worksheet_penelitian_dtps->setCellValue('D'.$row, '=IF(G'.$row.'='.$penelitian_dtps_ts2.',C'.$row.',0)');
	$worksheet_penelitian_dtps->getCell('D'.$row)->getStyle()->setQuotePrefix(true);
	
	$worksheet_penelitian_dtps->setCellValue('E'.$row, '=IF(G'.$row.'='.$penelitian_dtps_ts1.',C'.$row.',0)');
	$worksheet_penelitian_dtps->getCell('E'.$row)->getStyle()->setQuotePrefix(true);
	
	$worksheet_penelitian_dtps->setCellValue('F'.$row, '=IF(G'.$row.'='.$penelitian_dtps_ts.',C'.$row.',0)');
	$worksheet_penelitian_dtps->getCell('F'.$row)->getStyle()->setQuotePrefix(true);
}

$writer_penelitian_dtps = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_penelitian_dtps, 'Xls');
$writer_penelitian_dtps->save('./formatted/sapto_penelitian_dtps (F).xls');

$spreadsheet_penelitian_dtps->disconnectWorksheets();
unset($spreadsheet_penelitian_dtps);

// Load Format Baru
$spreadsheet_penelitian_dtps2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_penelitian_dtps (F).xls');
$worksheet_penelitian_dtps2 = $spreadsheet_penelitian_dtps2->getActiveSheet();

// Formasi Array SAPTO
$array_penelitian_dtps = $worksheet_penelitian_dtps2->toArray();
$data_penelitian_dtps = [];

foreach($worksheet_penelitian_dtps2->getRowIterator() as $row_id => $row) {
    if($worksheet_penelitian_dtps2->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['sumber_dana'] = $array_penelitian_dtps[$row_id-1][1];
            $item['judul_ts2'] = $array_penelitian_dtps[$row_id-1][3];
			$item['judul_ts1'] = $array_penelitian_dtps[$row_id-1][4];
			$item['judul_ts'] = $array_penelitian_dtps[$row_id-1][5];
			$item['jumlah'] = $array_penelitian_dtps[$row_id-1][6];
            $data_penelitian_dtps[] = $item;
        }
    }
}

$worksheet_penelitian_dtps3 = new \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet($spreadsheet_penelitian_dtps2, 'Sheet 2');
$spreadsheet_penelitian_dtps2->addSheet($worksheet_penelitian_dtps3);

$worksheet_penelitian_dtps3 = $spreadsheet_penelitian_dtps2->getSheetByName('Sheet 2');
$worksheet_penelitian_dtps3->fromArray($data_penelitian_dtps, NULL, 'A1');

$highestRow_penelitian_dtps3 = $worksheet_penelitian_dtps3->getHighestRow();

$row_jumlah = -1;

for($group = 1;$group <= 3; $group++) {
	
	$row_jumlah += 2;
	$baris_awal = $row_jumlah;
	$judul_penelitian_ts2 = $worksheet_penelitian_dtps3->getCell('B'.($row_jumlah))->getValue();
	$judul_penelitian_ts1 = $worksheet_penelitian_dtps3->getCell('C'.($row_jumlah))->getValue();
	$judul_penelitian_ts = $worksheet_penelitian_dtps3->getCell('D'.($row_jumlah))->getValue();	
	
	for($row = $row_jumlah;$row <= ($highestRow_penelitian_dtps3+1); $row++) {
		if($worksheet_penelitian_dtps3->getCell('A'.$row)->getValue() == $worksheet_penelitian_dtps3->getCell('A'.($row+1))->getValue()) {
			$judul_penelitian_ts2 += $worksheet_penelitian_dtps3->getCell('B'.($row+1))->getValue();
			$judul_penelitian_ts1 += $worksheet_penelitian_dtps3->getCell('C'.($row+1))->getValue();
			$judul_penelitian_ts += $worksheet_penelitian_dtps3->getCell('D'.($row+1))->getValue();
			
			$row_jumlah++;
		} else {
			break;
		}
	}
	
	$worksheet_penelitian_dtps3->insertNewRowBefore(($row_jumlah+1), 1);
	$worksheet_penelitian_dtps3->setCellValue('B'.($row_jumlah+1), $judul_penelitian_ts2);
	$worksheet_penelitian_dtps3->setCellValue('C'.($row_jumlah+1), $judul_penelitian_ts1);
	$worksheet_penelitian_dtps3->setCellValue('D'.($row_jumlah+1), $judul_penelitian_ts);
	
	${"penelitian_dtps".$group} = $worksheet_penelitian_dtps3->rangeToArray('B'.($row_jumlah+1).':D'.($row_jumlah+1), NULL, TRUE, TRUE, TRUE);
}

// Menggabungkan jumlah penelitian bersumber DANA LOKAL ITS dan MANDIRI
$penelitian_dtps_lokal = [];
$penelitian_dtps_lokal[0] = $penelitian_dtps1[0] + $penelitian_dtps3[0]
$penelitian_dtps_lokal[1] = $penelitian_dtps1[1] + $penelitian_dtps3[1]
$penelitian_dtps_lokal[2] = $penelitian_dtps1[2] + $penelitian_dtps3[2]


$writer_penelitian_dtps2 = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_penelitian_dtps2, 'Xls');
$writer_penelitian_dtps2->save('./formatted/sapto_penelitian_dtps (F).xls');

$spreadsheet_penelitian_dtps2->disconnectWorksheets();
unset($spreadsheet_penelitian_dtps2);



/**
 * Tabel 3.b.3 PkM DTPS
 */

$sql_pkm_dtps = "SELECT id, sumber_dana, jml_penelitian, tahun, id_prodi FROM akreditasi.sapto_pkm_dtps WHERE id_prodi = '".$nama_prodi."' ORDER BY sumber_dana";
$stmt = sqlsrv_query( $conn, $sql_pkm_dtps );
if( $stmt === false) {
    die( print_r( sqlsrv_errors(), true) );
}

$data_pkm_dtps1 = [];

while( $row = sqlsrv_fetch_array( $stmt, SQLSRV_FETCH_NUMERIC) ) {
	 $data_array_pkm_dtps[] = array(
		$row[0], $row[1], $row[2], $row[3], $row[4]
	);
}

$data_pkm_dtps1 = array_merge($data_pkm_dtps1, $data_array_pkm_dtps); 

sqlsrv_free_stmt( $stmt);

$spreadsheet_pkm_dtps1 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./blank/sapto_pkm_dtps.xlsx');
$worksheet_pkm_dtps1 = $spreadsheet_pkm_dtps1->getActiveSheet();

$worksheet_pkm_dtps1->fromArray($data_pkm_dtps1, NULL, 'A2');

$writer_pkm_dtps1 = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_pkm_dtps1, 'Xlsx');
$writer_pkm_dtps1->save('./raw/sapto_pkm_dtps.xlsx');

$spreadsheet_pkm_dtps1->disconnectWorksheets();
unset($spreadsheet_pkm_dtps1);

$spreadsheet_pkm_dtps = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_pkm_dtps.xlsx');
$worksheet_pkm_dtps = $spreadsheet_pkm_dtps->getActiveSheet();

$worksheet_pkm_dtps->insertNewColumnBefore('D', 4);

$worksheet_pkm_dtps->getCell('D1')->setValue('Jumlah TS-2');
$worksheet_pkm_dtps->getCell('E1')->setValue('Jumlah TS-1');
$worksheet_pkm_dtps->getCell('F1')->setValue('Jumlah TS');
$worksheet_pkm_dtps->getCell('G1')->setValue('Jumlah');

$highestRow_pkm_dtps = $worksheet_pkm_dtps->getHighestRow();

$pkm_dtps_ts = intval(date("Y"));
$pkm_dtps_ts1 = intval(date("Y", strtotime("-1 year")));
$pkm_dtps_ts2 = intval(date("Y", strtotime("-2 year")));

$worksheet_pkm_dtps->setAutoFilter('B1:H'.$highestRow_pkm_dtps);
$autoFilter_pkm_dtps = $worksheet_pkm_dtps->getAutoFilter();
$columnFilter_pkm_dtps = $autoFilter_pkm_dtps->getColumn('G');
$columnFilter_pkm_dtps->setFilterType(
    \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_FILTER
);
$columnFilter_pkm_dtps->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        $pkm_dtps_ts2
    );
$columnFilter_pkm_dtps->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        $pkm_dtps_ts1
    );
$columnFilter_pkm_dtps->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        $pkm_dtps_ts
    );

$autoFilter_pkm_dtps->showHideRows();

for($row = 2;$row <= $highestRow_pkm_dtps; $row++) {
	$worksheet_pkm_dtps->setCellValue('D'.$row, '=IF(G'.$row.'='.$pkm_dtps_ts2.',C'.$row.',0)');
	$worksheet_pkm_dtps->getCell('D'.$row)->getStyle()->setQuotePrefix(true);
	
	$worksheet_pkm_dtps->setCellValue('E'.$row, '=IF(G'.$row.'='.$pkm_dtps_ts1.',C'.$row.',0)');
	$worksheet_pkm_dtps->getCell('E'.$row)->getStyle()->setQuotePrefix(true);
	
	$worksheet_pkm_dtps->setCellValue('F'.$row, '=IF(G'.$row.'='.$pkm_dtps_ts.',C'.$row.',0)');
	$worksheet_pkm_dtps->getCell('F'.$row)->getStyle()->setQuotePrefix(true);
}

$writer_pkm_dtps = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_pkm_dtps, 'Xls');
$writer_pkm_dtps->save('./formatted/sapto_pkm_dtps (F).xls');

$spreadsheet_pkm_dtps->disconnectWorksheets();
unset($spreadsheet_pkm_dtps);

// Load Format Baru
$spreadsheet_pkm_dtps2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_pkm_dtps (F).xls');
$worksheet_pkm_dtps2 = $spreadsheet_pkm_dtps2->getActiveSheet();

// Formasi Array SAPTO
$array_pkm_dtps = $worksheet_pkm_dtps2->toArray();
$data_pkm_dtps = [];

foreach($worksheet_pkm_dtps2->getRowIterator() as $row_id => $row) {
    if($worksheet_pkm_dtps2->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['sumber_dana'] = $array_pkm_dtps[$row_id-1][1];
            $item['judul_ts2'] = $array_pkm_dtps[$row_id-1][3];
			$item['judul_ts1'] = $array_pkm_dtps[$row_id-1][4];
			$item['judul_ts'] = $array_pkm_dtps[$row_id-1][5];
			$item['jumlah'] = $array_pkm_dtps[$row_id-1][6];
            $data_pkm_dtps[] = $item;
        }
    }
}

$worksheet_pkm_dtps3 = new \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet($spreadsheet_pkm_dtps2, 'Sheet 2');
$spreadsheet_pkm_dtps2->addSheet($worksheet_pkm_dtps3);

$worksheet_pkm_dtps3 = $spreadsheet_pkm_dtps2->getSheetByName('Sheet 2');
$worksheet_pkm_dtps3->fromArray($data_pkm_dtps, NULL, 'A1');

$highestRow_pkm_dtps3 = $worksheet_pkm_dtps3->getHighestRow();

$row_jumlah = -1;

for($group = 1;$group <= 3; $group++) {
	
	$row_jumlah += 2;
	$baris_awal = $row_jumlah;
	$judul_pkm_ts2 = $worksheet_pkm_dtps3->getCell('B'.($row_jumlah))->getValue();
	$judul_pkm_ts1 = $worksheet_pkm_dtps3->getCell('C'.($row_jumlah))->getValue();
	$judul_pkm_ts = $worksheet_pkm_dtps3->getCell('D'.($row_jumlah))->getValue();	
	
	for($row = $row_jumlah;$row <= ($highestRow_pkm_dtps3+1); $row++) {
		if($worksheet_pkm_dtps3->getCell('A'.$row)->getValue() == $worksheet_pkm_dtps3->getCell('A'.($row+1))->getValue()) {
			$judul_pkm_ts2 += $worksheet_pkm_dtps3->getCell('B'.($row+1))->getValue();
			$judul_pkm_ts1 += $worksheet_pkm_dtps3->getCell('C'.($row+1))->getValue();
			$judul_pkm_ts += $worksheet_pkm_dtps3->getCell('D'.($row+1))->getValue();
			
			$row_jumlah++;
		} else {
			break;
		}
	}
	
	$worksheet_pkm_dtps3->insertNewRowBefore(($row_jumlah+1), 1);
	$worksheet_pkm_dtps3->setCellValue('B'.($row_jumlah+1), $judul_pkm_ts2);
	$worksheet_pkm_dtps3->setCellValue('C'.($row_jumlah+1), $judul_pkm_ts1);
	$worksheet_pkm_dtps3->setCellValue('D'.($row_jumlah+1), $judul_pkm_ts);
	
	${"pkm_dtps".$group} = $worksheet_pkm_dtps3->rangeToArray('B'.($row_jumlah+1).':D'.($row_jumlah+1), NULL, TRUE, TRUE, TRUE);
}

// Menggabungkan jumlah PkM bersumber DANA LOKAL ITS dan MANDIRI
$pkm_dtps_lokal = [];
$pkm_dtps_lokal[0] = $pkm_dtps1[0] + $pkm_dtps3[0]
$pkm_dtps_lokal[1] = $pkm_dtps1[1] + $pkm_dtps3[1]
$pkm_dtps_lokal[2] = $pkm_dtps1[2] + $pkm_dtps3[2]

$writer_pkm_dtps2 = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_pkm_dtps2, 'Xls');
$writer_pkm_dtps2->save('./formatted/sapto_pkm_dtps (F).xls');

$spreadsheet_pkm_dtps2->disconnectWorksheets();
unset($spreadsheet_pkm_dtps2);



/**
 * Tabel 3.b.4 Publikasi Ilmiah DTPS
 */


// Publikasi Ilmiah DTPS Jurnal Seminar
$sql_publikasi_dtps_jurnal_seminar = "SELECT id, tingkat_publikasi, tahun, COUNT(DISTINCT judul_publikasi), id_prodi FROM akreditasi.sapto_publikasi_ilmiah_dtps_jurnal_seminar WHERE id_prodi = '".$nama_prodi."' GROUP BY tingkat_publikasi, tahun ORDER BY tingkat_publikasi";
$stmt = sqlsrv_query( $conn, $sql_publikasi_dtps_jurnal_seminar );
if( $stmt === false) {
    die( print_r( sqlsrv_errors(), true) );
}

$data_publikasi_dtps_jurnal_seminar1 = [];

while( $row = sqlsrv_fetch_array( $stmt, SQLSRV_FETCH_NUMERIC) ) {
	 $data_array_publikasi_dtps_jurnal_seminar[] = array(
		$row[0], $row[1], $row[2], $row[3], $row[4]
	);
}

$data_publikasi_dtps_jurnal_seminar1 = array_merge($data_publikasi_dtps_jurnal_seminar1, $data_array_publikasi_dtps_jurnal_seminar); 

sqlsrv_free_stmt( $stmt);

$spreadsheet_publikasi_dtps_jurnal_seminar1 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./blank/sapto_publikasi_dtps_jurnal_seminar.xlsx');
$worksheet_publikasi_dtps_jurnal_seminar1 = $spreadsheet_publikasi_dtps_jurnal_seminar1->getActiveSheet();

$worksheet_publikasi_dtps_jurnal_seminar1->fromArray($data_publikasi_dtps_jurnal_seminar1, NULL, 'A2');

$writer_publikasi_dtps_jurnal_seminar1 = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_publikasi_dtps_jurnal_seminar1, 'Xlsx');
$writer_publikasi_dtps_jurnal_seminar1->save('./raw/sapto_publikasi_dtps_jurnal_seminar.xlsx');

$spreadsheet_publikasi_dtps_jurnal_seminar1->disconnectWorksheets();
unset($spreadsheet_publikasi_dtps_jurnal_seminar1);

$spreadsheet_publikasi_dtps_jurnal_seminar = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_publikasi_dtps_jurnal_seminar.xlsx');
$worksheet_publikasi_dtps_jurnal_seminar = $spreadsheet_publikasi_dtps_jurnal_seminar->getActiveSheet();

$worksheet_publikasi_dtps_jurnal_seminar->insertNewColumnBefore('E', 4);

$worksheet_publikasi_dtps_jurnal_seminar->getCell('E1')->setValue('Jumlah TS-2');
$worksheet_publikasi_dtps_jurnal_seminar->getCell('F1')->setValue('Jumlah TS-1');
$worksheet_publikasi_dtps_jurnal_seminar->getCell('G1')->setValue('Jumlah TS');
$worksheet_publikasi_dtps_jurnal_seminar->getCell('H1')->setValue('Jumlah');

$highestRow_publikasi_dtps_jurnal_seminar = $worksheet_publikasi_dtps_jurnal_seminar->getHighestRow();

$publikasi_dtps_jurnal_seminar_ts = intval(date("Y"));
$publikasi_dtps_jurnal_seminar_ts1 = intval(date("Y", strtotime("-1 year")));
$publikasi_dtps_jurnal_seminar_ts2 = intval(date("Y", strtotime("-2 year")));

$worksheet_publikasi_dtps_jurnal_seminar->setAutoFilter('B1:I'.$highestRow_publikasi_dtps_jurnal_seminar);
$autoFilter_publikasi_dtps_jurnal_seminar = $worksheet_publikasi_dtps_jurnal_seminar->getAutoFilter();
$columnFilter_publikasi_dtps_jurnal_seminar = $autoFilter_publikasi_dtps_jurnal_seminar->getColumn('C');
$columnFilter_publikasi_dtps_jurnal_seminar->setFilterType(
    \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_FILTER
);
$columnFilter_publikasi_dtps_jurnal_seminar->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        $publikasi_dtps_jurnal_seminar_ts2
    );
$columnFilter_publikasi_dtps_jurnal_seminar->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        $publikasi_dtps_jurnal_seminar_ts1
    );
$columnFilter_publikasi_dtps_jurnal_seminar->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        $publikasi_dtps_jurnal_seminar_ts
    );

$autoFilter_publikasi_dtps_jurnal_seminar->showHideRows();

for($row = 2;$row <= $highestRow_publikasi_dtps_jurnal_seminar; $row++) {
	$worksheet_publikasi_dtps_jurnal_seminar->setCellValue('E'.$row, '=IF(C'.$row.'='.$publikasi_dtps_jurnal_seminar_ts2.',D'.$row.',0)');
	$worksheet_publikasi_dtps_jurnal_seminar->getCell('E'.$row)->getStyle()->setQuotePrefix(true);
	
	$worksheet_publikasi_dtps_jurnal_seminar->setCellValue('F'.$row, '=IF(C'.$row.'='.$publikasi_dtps_jurnal_seminar_ts1.',D'.$row.',0)');
	$worksheet_publikasi_dtps_jurnal_seminar->getCell('F'.$row)->getStyle()->setQuotePrefix(true);
	
	$worksheet_publikasi_dtps_jurnal_seminar->setCellValue('G'.$row, '=IF(C'.$row.'='.$publikasi_dtps_jurnal_seminar_ts.',D'.$row.',0)');
	$worksheet_publikasi_dtps_jurnal_seminar->getCell('G'.$row)->getStyle()->setQuotePrefix(true);
}

$writer_publikasi_dtps_jurnal_seminar = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_publikasi_dtps_jurnal_seminar, 'Xls');
$writer_publikasi_dtps_jurnal_seminar->save('./formatted/sapto_publikasi_dtps_jurnal_seminar (F).xls');

$spreadsheet_publikasi_dtps_jurnal_seminar->disconnectWorksheets();
unset($spreadsheet_publikasi_dtps_jurnal_seminar);

// Load Format Baru
$spreadsheet_publikasi_dtps_jurnal_seminar2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_publikasi_dtps_jurnal_seminar (F).xls');
$worksheet_publikasi_dtps_jurnal_seminar2 = $spreadsheet_publikasi_dtps_jurnal_seminar2->getActiveSheet();

// Formasi Array SAPTO
$array_publikasi_dtps_jurnal_seminar = $worksheet_publikasi_dtps_jurnal_seminar2->toArray();
$data_publikasi_dtps_jurnal_seminar = [];

foreach($worksheet_publikasi_dtps_jurnal_seminar2->getRowIterator() as $row_id => $row) {
    if($worksheet_publikasi_dtps_jurnal_seminar2->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['jenis_publikasi'] = $array_publikasi_dtps_jurnal_seminar[$row_id-1][1];
            $item['judul_ts2'] = $array_publikasi_dtps_jurnal_seminar[$row_id-1][4];
			$item['judul_ts1'] = $array_publikasi_dtps_jurnal_seminar[$row_id-1][5];
			$item['judul_ts'] = $array_publikasi_dtps_jurnal_seminar[$row_id-1][6];
			$item['jumlah'] = $array_publikasi_dtps_jurnal_seminar[$row_id-1][7];
            $data_publikasi_dtps_jurnal_seminar[] = $item;
        }
    }
}

$worksheet_publikasi_dtps_jurnal_seminar3 = new \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet($spreadsheet_publikasi_dtps_jurnal_seminar2, 'Sheet 2');
$spreadsheet_publikasi_dtps_jurnal_seminar2->addSheet($worksheet_publikasi_dtps_jurnal_seminar3);

$worksheet_publikasi_dtps_jurnal_seminar3 = $spreadsheet_publikasi_dtps_jurnal_seminar2->getSheetByName('Sheet 2');
$worksheet_publikasi_dtps_jurnal_seminar3->fromArray($data_publikasi_dtps_jurnal_seminar, NULL, 'A1');

$highestRow_publikasi_dtps_jurnal_seminar3 = $worksheet_publikasi_dtps_jurnal_seminar3->getHighestRow();

$row_jumlah = -1;

for($group = 1;$group <= 7; $group++) {
	
	$row_jumlah += 2;
	$baris_awal = $row_jumlah;
	$judul_publikasi_jurnal_ts2 = $worksheet_publikasi_dtps_jurnal_seminar3->getCell('B'.($row_jumlah))->getValue();
	$judul_publikasi_jurnal_ts1 = $worksheet_publikasi_dtps_jurnal_seminar3->getCell('C'.($row_jumlah))->getValue();
	$judul_publikasi_jurnal_ts = $worksheet_publikasi_dtps_jurnal_seminar3->getCell('D'.($row_jumlah))->getValue();	
	
	for($row = $row_jumlah;$row <= ($highestRow_publikasi_dtps_jurnal_seminar3+1); $row++) {
		if($worksheet_publikasi_dtps_jurnal_seminar3->getCell('A'.$row)->getValue() == $worksheet_publikasi_dtps_jurnal_seminar3->getCell('A'.($row+1))->getValue()) {
			$judul_publikasi_jurnal_ts2 += $worksheet_publikasi_dtps_jurnal_seminar3->getCell('B'.($row+1))->getValue();
			$judul_publikasi_jurnal_ts1 += $worksheet_publikasi_dtps_jurnal_seminar3->getCell('C'.($row+1))->getValue();
			$judul_publikasi_jurnal_ts += $worksheet_publikasi_dtps_jurnal_seminar3->getCell('D'.($row+1))->getValue();
			
			$row_jumlah++;
		} else {
			break;
		}
	}
	
	$worksheet_publikasi_dtps_jurnal_seminar3->insertNewRowBefore(($row_jumlah+1), 1);
	$worksheet_publikasi_dtps_jurnal_seminar3->setCellValue('B'.($row_jumlah+1), $judul_publikasi_jurnal_ts2);
	$worksheet_publikasi_dtps_jurnal_seminar3->setCellValue('C'.($row_jumlah+1), $judul_publikasi_jurnal_ts1);
	$worksheet_publikasi_dtps_jurnal_seminar3->setCellValue('D'.($row_jumlah+1), $judul_publikasi_jurnal_ts);
	
	${"publikasi_dtps_jurnal_seminar".$group} = $worksheet_publikasi_dtps_jurnal_seminar3->rangeToArray('B'.($row_jumlah+1).':D'.($row_jumlah+1), NULL, TRUE, TRUE, TRUE);
}

$writer_publikasi_dtps_jurnal_seminar2 = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_publikasi_dtps_jurnal_seminar2, 'Xls');
$writer_publikasi_dtps_jurnal_seminar2->save('./formatted/sapto_publikasi_dtps_jurnal_seminar (F).xls');

$spreadsheet_publikasi_dtps_jurnal_seminar2->disconnectWorksheets();
unset($spreadsheet_publikasi_dtps_jurnal_seminar2);


// Publikasi Ilmiah DTPS Selain Jurnal Seminar
$sql_publikasi_dtps_non_jurnal = "SELECT id, tingkat_publikasi, tahun, COUNT(DISTINCT judul_publikasi), id_prodi FROM akreditasi.sapto_publikasi_ilmiah_dtps_selain_jurnal_seminar WHERE id_prodi = '".$nama_prodi."' GROUP BY tingkat_publikasi, tahun ORDER BY tingkat_publikasi";
$stmt = sqlsrv_query( $conn, $sql_publikasi_dtps_non_jurnal );
if( $stmt === false) {
    die( print_r( sqlsrv_errors(), true) );
}

$data_publikasi_dtps_non_jurnal1 = [];

while( $row = sqlsrv_fetch_array( $stmt, SQLSRV_FETCH_NUMERIC) ) {
	 $data_array_publikasi_dtps_non_jurnal[] = array(
		$row[0], $row[1], $row[2], $row[3], $row[4]
	);
}

$data_publikasi_dtps_non_jurnal1 = array_merge($data_publikasi_dtps_non_jurnal1, $data_array_publikasi_dtps_non_jurnal); 

sqlsrv_free_stmt( $stmt);

$spreadsheet_publikasi_dtps_non_jurnal1 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./blank/sapto_publikasi_dtps_non_jurnal.xlsx');
$worksheet_publikasi_dtps_non_jurnal1 = $spreadsheet_publikasi_dtps_non_jurnal1->getActiveSheet();

$worksheet_publikasi_dtps_non_jurnal1->fromArray($data_publikasi_dtps_non_jurnal1, NULL, 'A2');

$writer_publikasi_dtps_non_jurnal1 = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_publikasi_dtps_non_jurnal1, 'Xlsx');
$writer_publikasi_dtps_non_jurnal1->save('./raw/sapto_publikasi_dtps_non_jurnal.xlsx');

$spreadsheet_publikasi_dtps_non_jurnal1->disconnectWorksheets();
unset($spreadsheet_publikasi_dtps_non_jurnal1);

$spreadsheet_publikasi_dtps_non_jurnal = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_publikasi_dtps_non_jurnal.xlsx');
$worksheet_publikasi_dtps_non_jurnal = $spreadsheet_publikasi_dtps_non_jurnal->getActiveSheet();

$worksheet_publikasi_dtps_non_jurnal->insertNewColumnBefore('E', 4);

$worksheet_publikasi_dtps_non_jurnal->getCell('E1')->setValue('Jumlah TS-2');
$worksheet_publikasi_dtps_non_jurnal->getCell('F1')->setValue('Jumlah TS-1');
$worksheet_publikasi_dtps_non_jurnal->getCell('G1')->setValue('Jumlah TS');
$worksheet_publikasi_dtps_non_jurnal->getCell('H1')->setValue('Jumlah');

$highestRow_publikasi_dtps_non_jurnal = $worksheet_publikasi_dtps_non_jurnal->getHighestRow();

$publikasi_dtps_non_jurnal_ts = intval(date("Y"));
$publikasi_dtps_non_jurnal_ts1 = intval(date("Y", strtotime("-1 year")));
$publikasi_dtps_non_jurnal_ts2 = intval(date("Y", strtotime("-2 year")));

$worksheet_publikasi_dtps_non_jurnal->setAutoFilter('B1:I'.$highestRow_publikasi_dtps_non_jurnal);
$autoFilter_publikasi_dtps_non_jurnal = $worksheet_publikasi_dtps_non_jurnal->getAutoFilter();
$columnFilter_publikasi_dtps_non_jurnal = $autoFilter_publikasi_dtps_non_jurnal->getColumn('C');
$columnFilter_publikasi_dtps_non_jurnal->setFilterType(
    \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_FILTER
);
$columnFilter_publikasi_dtps_non_jurnal->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        $publikasi_dtps_non_jurnal_ts2
    );
$columnFilter_publikasi_dtps_non_jurnal->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        $publikasi_dtps_non_jurnal_ts1
    );
$columnFilter_publikasi_dtps_non_jurnal->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        $publikasi_dtps_non_jurnal_ts
    );

$autoFilter_publikasi_dtps_non_jurnal->showHideRows();

for($row = 2;$row <= $highestRow_publikasi_dtps_non_jurnal; $row++) {
	$worksheet_publikasi_dtps_non_jurnal->setCellValue('E'.$row, '=IF(C'.$row.'='.$publikasi_dtps_non_jurnal_ts2.',D'.$row.',0)');
	$worksheet_publikasi_dtps_non_jurnal->getCell('E'.$row)->getStyle()->setQuotePrefix(true);
	
	$worksheet_publikasi_dtps_non_jurnal->setCellValue('F'.$row, '=IF(C'.$row.'='.$publikasi_dtps_non_jurnal_ts1.',D'.$row.',0)');
	$worksheet_publikasi_dtps_non_jurnal->getCell('F'.$row)->getStyle()->setQuotePrefix(true);
	
	$worksheet_publikasi_dtps_non_jurnal->setCellValue('G'.$row, '=IF(C'.$row.'='.$publikasi_dtps_non_jurnal_ts.',D'.$row.',0)');
	$worksheet_publikasi_dtps_non_jurnal->getCell('G'.$row)->getStyle()->setQuotePrefix(true);
}

$writer_publikasi_dtps_non_jurnal = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_publikasi_dtps_non_jurnal, 'Xls');
$writer_publikasi_dtps_non_jurnal->save('./formatted/sapto_publikasi_dtps_non_jurnal (F).xls');

$spreadsheet_publikasi_dtps_non_jurnal->disconnectWorksheets();
unset($spreadsheet_publikasi_dtps_non_jurnal);

// Load Format Baru
$spreadsheet_publikasi_dtps_non_jurnal2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_publikasi_dtps_non_jurnal (F).xls');
$worksheet_publikasi_dtps_non_jurnal2 = $spreadsheet_publikasi_dtps_non_jurnal2->getActiveSheet();

// Formasi Array SAPTO
$array_publikasi_dtps_non_jurnal = $worksheet_publikasi_dtps_non_jurnal2->toArray();
$data_publikasi_dtps_non_jurnal = [];

foreach($worksheet_publikasi_dtps_non_jurnal2->getRowIterator() as $row_id => $row) {
    if($worksheet_publikasi_dtps_non_jurnal2->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['jenis_publikasi'] = $array_publikasi_dtps_non_jurnal[$row_id-1][1];
            $item['judul_ts2'] = $array_publikasi_dtps_non_jurnal[$row_id-1][4];
			$item['judul_ts1'] = $array_publikasi_dtps_non_jurnal[$row_id-1][5];
			$item['judul_ts'] = $array_publikasi_dtps_non_jurnal[$row_id-1][6];
			$item['jumlah'] = $array_publikasi_dtps_non_jurnal[$row_id-1][7];
            $data_publikasi_dtps_non_jurnal[] = $item;
        }
    }
}

$worksheet_publikasi_dtps_non_jurnal3 = new \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet($spreadsheet_publikasi_dtps_non_jurnal2, 'Sheet 2');
$spreadsheet_publikasi_dtps_non_jurnal2->addSheet($worksheet_publikasi_dtps_non_jurnal3);

$worksheet_publikasi_dtps_non_jurnal3 = $spreadsheet_publikasi_dtps_non_jurnal2->getSheetByName('Sheet 2');
$worksheet_publikasi_dtps_non_jurnal3->fromArray($data_publikasi_dtps_non_jurnal, NULL, 'A1');

$highestRow_publikasi_dtps_non_jurnal3 = $worksheet_publikasi_dtps_non_jurnal3->getHighestRow();

$row_jumlah = -1;

for($group = 1;$group <= 3; $group++) {
	
	$row_jumlah += 2;
	$baris_awal = $row_jumlah;
	$judul_publikasi_non_jurnal_ts2 = $worksheet_publikasi_dtps_non_jurnal3->getCell('B'.($row_jumlah))->getValue();
	$judul_publikasi_non_jurnal_ts1 = $worksheet_publikasi_dtps_non_jurnal3->getCell('C'.($row_jumlah))->getValue();
	$judul_publikasi_non_jurnal_ts = $worksheet_publikasi_dtps_non_jurnal3->getCell('D'.($row_jumlah))->getValue();	
	
	for($row = $row_jumlah;$row <= ($highestRow_publikasi_dtps_non_jurnal3+1); $row++) {
		if($worksheet_publikasi_dtps_non_jurnal3->getCell('A'.$row)->getValue() == $worksheet_publikasi_dtps_non_jurnal3->getCell('A'.($row+1))->getValue()) {
			$judul_publikasi_non_jurnal_ts2 += $worksheet_publikasi_dtps_non_jurnal3->getCell('B'.($row+1))->getValue();
			$judul_publikasi_non_jurnal_ts1 += $worksheet_publikasi_dtps_non_jurnal3->getCell('C'.($row+1))->getValue();
			$judul_publikasi_non_jurnal_ts += $worksheet_publikasi_dtps_non_jurnal3->getCell('D'.($row+1))->getValue();
			
			$row_jumlah++;
		} else {
			break;
		}
	}
	
	$worksheet_publikasi_dtps_non_jurnal3->insertNewRowBefore(($row_jumlah+1), 1);
	$worksheet_publikasi_dtps_non_jurnal3->setCellValue('B'.($row_jumlah+1), $judul_publikasi_non_jurnal_ts2);
	$worksheet_publikasi_dtps_non_jurnal3->setCellValue('C'.($row_jumlah+1), $judul_publikasi_non_jurnal_ts1);
	$worksheet_publikasi_dtps_non_jurnal3->setCellValue('D'.($row_jumlah+1), $judul_publikasi_non_jurnal_ts);
	
	${"publikasi_dtps_non_jurnal".$group} = $worksheet_publikasi_dtps_non_jurnal3->rangeToArray('B'.($row_jumlah+1).':D'.($row_jumlah+1), NULL, TRUE, TRUE, TRUE);
}

$writer_publikasi_dtps_non_jurnal2 = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_publikasi_dtps_non_jurnal2, 'Xls');
$writer_publikasi_dtps_non_jurnal2->save('./formatted/sapto_publikasi_dtps_non_jurnal (F).xls');

$spreadsheet_publikasi_dtps_non_jurnal2->disconnectWorksheets();
unset($spreadsheet_publikasi_dtps_non_jurnal2);



/**
 * Tabel 3.b.5 Luaran Penelitian/PkM Lainnya oleh DTPS - HKI (Paten, Paten Sederhana)
 */

$sql_luaran_dtps_hki_paten = "SELECT id, judul, tahun, id_prodi FROM akreditasi.sapto_luaran_penelitian_dtps_hki WHERE id_prodi = '".$nama_prodi."' AND kategori = 'paten'";
$stmt = sqlsrv_query( $conn, $sql_luaran_dtps_hki_paten );
if( $stmt === false) {
    die( print_r( sqlsrv_errors(), true) );
}

$data_luaran_dtps_hki_paten1 = [];

while( $row = sqlsrv_fetch_array( $stmt, SQLSRV_FETCH_NUMERIC) ) {
	 $data_array_luaran_dtps_hki_paten[] = array(
		$row[0], $row[1], $row[2], $row[3]
	);
}

$data_luaran_dtps_hki_paten1 = array_merge($data_luaran_dtps_hki_paten1, $data_array_luaran_dtps_hki_paten); 

sqlsrv_free_stmt( $stmt);

$spreadsheet_luaran_dtps_hki_paten1 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./blank/sapto_luaran_penelitian_dtps_hki_paten.xlsx');
$worksheet_luaran_dtps_hki_paten1 = $spreadsheet_luaran_dtps_hki_paten1->getActiveSheet();

$worksheet_luaran_dtps_hki_paten1->fromArray($data_luaran_dtps_hki_paten1, NULL, 'A2');

$writer_luaran_dtps_hki_paten1 = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_luaran_dtps_hki_paten1, 'Xlsx');
$writer_luaran_dtps_hki_paten1->save('./raw/sapto_luaran_penelitian_dtps_hki_paten.xlsx');

$spreadsheet_luaran_dtps_hki_paten1->disconnectWorksheets();
unset($spreadsheet_luaran_dtps_hki_paten1);

$spreadsheet_luaran_dtps_hki_paten = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_luaran_penelitian_dtps_hki_paten.xlsx');

$worksheet_luaran_dtps_hki_paten = $spreadsheet_luaran_dtps_hki_paten->getActiveSheet();

$worksheet_luaran_dtps_hki_paten->insertNewColumnBefore('D', 1);
$worksheet_luaran_dtps_hki_paten->getCell('D1')->setValue('Keterangan');

$writer_luaran_dtps_hki_paten = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_luaran_dtps_hki_paten, 'Xlsx');
$writer_luaran_dtps_hki_paten->save('./formatted/sapto_luaran_penelitian_dtps_hki_paten (F).xlsx');

$spreadsheet_luaran_dtps_hki_paten->disconnectWorksheets();
unset($spreadsheet_luaran_dtps_hki_paten);

// Load Format Baru
$spreadsheet_luaran_dtps_hki_paten2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_luaran_penelitian_dtps_hki_paten (F).xlsx');
$worksheet_luaran_dtps_hki_paten2 = $spreadsheet_luaran_dtps_hki_paten2->getActiveSheet();

// Formasi Array SAPTO
$array_luaran_dtps_hki_paten = $worksheet_luaran_dtps_hki_paten2->toArray();
$data_luaran_dtps_hki_paten = [];

foreach($worksheet_luaran_dtps_hki_paten2->getRowIterator() as $row_id => $row) {
    if($worksheet_luaran_dtps_hki_paten2->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['luaran_penelitian'] = $array_luaran_dtps_hki_paten[$row_id-1][1];
            $item['tahun'] = $array_luaran_dtps_hki_paten[$row_id-1][2];
			$item['keterangan'] = $array_luaran_dtps_hki_paten[$row_id-1][3];
            $data_luaran_dtps_hki_paten[] = $item;
        }
    }
}

$spreadsheet_luaran_dtps_hki_paten2->disconnectWorksheets();
unset($spreadsheet_luaran_dtps_hki_paten2);


/**
 * Tabel 3.b.5 Luaran Penelitian/PkM Lainnya oleh DTPS - HKI (Hak Cipta, Desain Produk Industri, dll.)
 */

$sql_luaran_dtps_hki_cipta = "SELECT id, judul, tahun, id_prodi FROM akreditasi.sapto_luaran_penelitian_dtps_hki WHERE id_prodi = '".$nama_prodi."' AND kategori = 'hak cipta'";
$stmt = sqlsrv_query( $conn, $sql_luaran_dtps_hki_cipta );
if( $stmt === false) {
    die( print_r( sqlsrv_errors(), true) );
}

$data_luaran_dtps_hki_cipta1 = [];

while( $row = sqlsrv_fetch_array( $stmt, SQLSRV_FETCH_NUMERIC) ) {
	 $data_array_luaran_dtps_hki_cipta[] = array(
		$row[0], $row[1], $row[2], $row[3]
	);
}

$data_luaran_dtps_hki_cipta1 = array_merge($data_luaran_dtps_hki_cipta1, $data_array_luaran_dtps_hki_cipta); 

sqlsrv_free_stmt( $stmt);

$spreadsheet_luaran_dtps_hki_cipta1 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./blank/sapto_luaran_penelitian_dtps_hki_cipta.xlsx');
$worksheet_luaran_dtps_hki_cipta1 = $spreadsheet_luaran_dtps_hki_cipta1->getActiveSheet();

$worksheet_luaran_dtps_hki_cipta1->fromArray($data_luaran_dtps_hki_cipta1, NULL, 'A2');

$writer_luaran_dtps_hki_cipta1 = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_luaran_dtps_hki_cipta1, 'Xlsx');
$writer_luaran_dtps_hki_cipta1->save('./raw/sapto_luaran_penelitian_dtps_hki_cipta.xlsx');

$spreadsheet_luaran_dtps_hki_cipta1->disconnectWorksheets();
unset($spreadsheet_luaran_dtps_hki_cipta1);

$spreadsheet_luaran_dtps_hki_cipta = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_luaran_penelitian_dtps_hki_cipta.xlsx');

$worksheet_luaran_dtps_hki_cipta = $spreadsheet_luaran_dtps_hki_cipta->getActiveSheet();

$worksheet_luaran_dtps_hki_cipta->insertNewColumnBefore('D', 1);
$worksheet_luaran_dtps_hki_cipta->getCell('D1')->setValue('Keterangan');

$writer_luaran_dtps_hki_cipta = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_luaran_dtps_hki_cipta, 'Xlsx');
$writer_luaran_dtps_hki_cipta->save('./formatted/sapto_luaran_penelitian_dtps_hki_cipta (F).xlsx');

$spreadsheet_luaran_dtps_hki_cipta->disconnectWorksheets();
unset($spreadsheet_luaran_dtps_hki_cipta);

// Load Format Baru
$spreadsheet_luaran_dtps_hki_cipta2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_luaran_penelitian_dtps_hki_cipta (F).xlsx');
$worksheet_luaran_dtps_hki_cipta2 = $spreadsheet_luaran_dtps_hki_cipta2->getActiveSheet();

// Formasi Array SAPTO
$array_luaran_dtps_hki_cipta = $worksheet_luaran_dtps_hki_cipta2->toArray();
$data_luaran_dtps_hki_cipta = [];

foreach($worksheet_luaran_dtps_hki_cipta2->getRowIterator() as $row_id => $row) {
    if($worksheet_luaran_dtps_hki_cipta2->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['luaran_penelitian'] = $array_luaran_dtps_hki_cipta[$row_id-1][1];
            $item['tahun'] = $array_luaran_dtps_hki_cipta[$row_id-1][2];
			$item['keterangan'] = $array_luaran_dtps_hki_cipta[$row_id-1][3];
            $data_luaran_dtps_hki_cipta[] = $item;
        }
    }
}

$spreadsheet_luaran_dtps_hki_cipta2->disconnectWorksheets();
unset($spreadsheet_luaran_dtps_hki_cipta2);


/**
 * Tabel 3.b.5 Luaran Penelitian/PkM Lainnya oleh DTPS - Teknologi Tepat Guna, Produk, Karya Seni, Rekayasa Sosial
 */

$sql_luaran_dtps_produk = "SELECT id, judul, tahun, id_prodi FROM akreditasi.sapto_luaran_penelitian_dtps_produk WHERE id_prodi = '".$nama_prodi."'";
$stmt = sqlsrv_query( $conn, $sql_luaran_dtps_produk );
if( $stmt === false) {
    die( print_r( sqlsrv_errors(), true) );
}

$data_luaran_dtps_produk1 = [];

while( $row = sqlsrv_fetch_array( $stmt, SQLSRV_FETCH_NUMERIC) ) {
	 $data_array_luaran_dtps_produk[] = array(
		$row[0], $row[1], $row[2], $row[3]
	);
}

$data_luaran_dtps_produk1 = array_merge($data_luaran_dtps_produk1, $data_array_luaran_dtps_produk); 

sqlsrv_free_stmt( $stmt);

$spreadsheet_luaran_dtps_produk1 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./blank/sapto_luaran_penelitian_dtps_produk.xlsx');
$worksheet_luaran_dtps_produk1 = $spreadsheet_luaran_dtps_produk1->getActiveSheet();

$worksheet_luaran_dtps_produk1->fromArray($data_luaran_dtps_produk1, NULL, 'A2');

$writer_luaran_dtps_produk1 = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_luaran_dtps_produk1, 'Xlsx');
$writer_luaran_dtps_produk1->save('./raw/sapto_luaran_penelitian_dtps_produk.xlsx');

$spreadsheet_luaran_dtps_produk1->disconnectWorksheets();
unset($spreadsheet_luaran_dtps_produk1);

$spreadsheet_luaran_dtps_produk = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_luaran_penelitian_dtps_produk.xlsx');

$worksheet_luaran_dtps_produk = $spreadsheet_luaran_dtps_produk->getActiveSheet();

$worksheet_luaran_dtps_produk->insertNewColumnBefore('D', 1);
$worksheet_luaran_dtps_produk->getCell('D1')->setValue('Keterangan');

$writer_luaran_dtps_produk = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_luaran_dtps_produk, 'Xlsx');
$writer_luaran_dtps_produk->save('./formatted/sapto_luaran_penelitian_dtps_produk (F).xlsx');

$spreadsheet_luaran_dtps_produk->disconnectWorksheets();
unset($spreadsheet_luaran_dtps_produk);

// Load Format Baru
$spreadsheet_luaran_dtps_produk2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_luaran_penelitian_dtps_produk (F).xlsx');
$worksheet_luaran_dtps_produk2 = $spreadsheet_luaran_dtps_produk2->getActiveSheet();

// Formasi Array SAPTO
$array_luaran_dtps_produk = $worksheet_luaran_dtps_produk2->toArray();
$data_luaran_dtps_produk = [];

foreach($worksheet_luaran_dtps_produk2->getRowIterator() as $row_id => $row) {
    if($worksheet_luaran_dtps_produk2->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['luaran_penelitian'] = $array_luaran_dtps_produk[$row_id-1][1];
            $item['tahun'] = $array_luaran_dtps_produk[$row_id-1][2];
			$item['keterangan'] = $array_luaran_dtps_produk[$row_id-1][3];
            $data_luaran_dtps_produk[] = $item;
        }
    }
}

$spreadsheet_luaran_dtps_produk2->disconnectWorksheets();
unset($spreadsheet_luaran_dtps_produk2);


/**
 * Tabel 3.b.5 Luaran Penelitian/PkM Lainnya oleh DTPS - Buku Ber-ISBN, Book Chapter
 */

$sql_luaran_dtps_bukuisbn = "SELECT id, judul, tahun_terbit, keterangan, id_prodi FROM akreditasi.sapto_luaran_penelitian_dtps_bukuisbn WHERE id_prodi = '".$nama_prodi."'";
$stmt = sqlsrv_query( $conn, $sql_luaran_dtps_bukuisbn );
if( $stmt === false) {
    die( print_r( sqlsrv_errors(), true) );
}

$data_luaran_dtps_bukuisbn1 = [];

while( $row = sqlsrv_fetch_array( $stmt, SQLSRV_FETCH_NUMERIC) ) {
	 $data_array_luaran_dtps_bukuisbn[] = array(
		$row[0], $row[1], $row[2], $row[3]
	);
}

$data_luaran_dtps_bukuisbn1 = array_merge($data_luaran_dtps_bukuisbn1, $data_array_luaran_dtps_bukuisbn); 

sqlsrv_free_stmt( $stmt);

$spreadsheet_luaran_dtps_bukuisbn1 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./blank/sapto_luaran_penelitian_dtps_bukuisbn.xlsx');
$worksheet_luaran_dtps_bukuisbn1 = $spreadsheet_luaran_dtps_bukuisbn1->getActiveSheet();

$worksheet_luaran_dtps_bukuisbn1->fromArray($data_luaran_dtps_bukuisbn1, NULL, 'A2');

$writer_luaran_dtps_bukuisbn1 = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_luaran_dtps_bukuisbn1, 'Xlsx');
$writer_luaran_dtps_bukuisbn1->save('./raw/sapto_luaran_penelitian_dtps_bukuisbn.xlsx');

$spreadsheet_luaran_dtps_bukuisbn1->disconnectWorksheets();
unset($spreadsheet_luaran_dtps_bukuisbn1);

// Load Format Baru
$spreadsheet_luaran_dtps_bukuisbn2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_luaran_penelitian_dtps_bukuisbn.xlsx');
$worksheet_luaran_dtps_bukuisbn2 = $spreadsheet_luaran_dtps_bukuisbn2->getActiveSheet();

// Formasi Array SAPTO
$array_luaran_dtps_bukuisbn = $worksheet_luaran_dtps_bukuisbn2->toArray();
$data_luaran_dtps_bukuisbn = [];

foreach($worksheet_luaran_dtps_bukuisbn2->getRowIterator() as $row_id => $row) {
    if($worksheet_luaran_dtps_bukuisbn2->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['luaran_penelitian'] = $array_luaran_dtps_bukuisbn[$row_id-1][1];
            $item['tahun'] = $array_luaran_dtps_bukuisbn[$row_id-1][2];
			$item['keterangan'] = $array_luaran_dtps_bukuisbn[$row_id-1][3];
            $data_luaran_dtps_bukuisbn[] = $item;
        }
    }
}

$spreadsheet_luaran_dtps_bukuisbn2->disconnectWorksheets();
unset($spreadsheet_luaran_dtps_bukuisbn2);


/**
 * Tabel 3.b.6 Karya Ilmiah DTPS yang Disitasi
 */

$sql_karya_disitasi_dtps = "SELECT nama, title, jml_sitasi, id_prodi FROM akreditasi.sapto_karya_ilmiah_disitasi_dtps WHERE id_prodi = '".$nama_prodi."'";
$stmt = sqlsrv_query( $conn, $sql_karya_disitasi_dtps );
if( $stmt === false) {
    die( print_r( sqlsrv_errors(), true) );
}

$data_karya_disitasi_dtps1 = [];

while( $row = sqlsrv_fetch_array( $stmt, SQLSRV_FETCH_NUMERIC) ) {
	 $data_array_karya_disitasi_dtps[] = array(
		$row[0], $row[1], $row[2], $row[3]
	);
}

$data_karya_disitasi_dtps1 = array_merge($data_karya_disitasi_dtps1, $data_array_karya_disitasi_dtps); 

sqlsrv_free_stmt( $stmt);

$spreadsheet_karya_disitasi_dtps1 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./blank/sapto_karya_ilmiah_disitasi_dtps.xlsx');
$worksheet_karya_disitasi_dtps1 = $spreadsheet_karya_disitasi_dtps1->getActiveSheet();

$worksheet_karya_disitasi_dtps1->fromArray($data_karya_disitasi_dtps1, NULL, 'A2');

$writer_karya_disitasi_dtps1 = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_karya_disitasi_dtps1, 'Xlsx');
$writer_karya_disitasi_dtps1->save('./raw/sapto_karya_ilmiah_disitasi_dtps.xlsx');

$spreadsheet_karya_disitasi_dtps1->disconnectWorksheets();
unset($spreadsheet_karya_disitasi_dtps1);

// Load Format Baru
$spreadsheet_karya_disitasi_dtps2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_karya_ilmiah_disitasi_dtps.xlsx');
$worksheet_karya_disitasi_dtps2 = $spreadsheet_karya_disitasi_dtps2->getActiveSheet();

// Formasi Array SAPTO
$array_karya_disitasi_dtps = $worksheet_karya_disitasi_dtps2->toArray();
$data_karya_disitasi_dtps = [];

foreach($worksheet_karya_disitasi_dtps2->getRowIterator() as $row_id => $row) {
    if($worksheet_karya_disitasi_dtps2->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['nama_dosen'] = $array_karya_disitasi_dtps[$row_id-1][0];
            $item['judul_artikel'] = $array_karya_disitasi_dtps[$row_id-1][1];
			$item['jumlah_sitasi'] = $array_karya_disitasi_dtps[$row_id-1][2];
            $data_karya_disitasi_dtps[] = $item;
        }
    }
}

$spreadsheet_karya_disitasi_dtps2->disconnectWorksheets();
unset($spreadsheet_karya_disitasi_dtps2);


/**
 * Tabel 3.b.7 Produk/Jasa DTPS yang Diadopsi oleh Industri/Masyarakat
 */

$sql_produk_jasa_dtps = "SELECT id, nama_dosen, nama_produk_jasa, deskripsi, bukti, id_prodi FROM akreditasi.sapto_produk_jasa_masyarakat_dtps WHERE id_prodi = '".$nama_prodi."'";
$stmt = sqlsrv_query( $conn, $sql_produk_jasa_dtps );
if( $stmt === false) {
    die( print_r( sqlsrv_errors(), true) );
}

$data_produk_jasa_dtps1 = [];

while( $row = sqlsrv_fetch_array( $stmt, SQLSRV_FETCH_NUMERIC) ) {
	 $data_array_produk_jasa_dtps[] = array(
		$row[0], $row[1], $row[2], $row[3], $row[4], $row[5]
	);
}

$data_produk_jasa_dtps1 = array_merge($data_produk_jasa_dtps1, $data_array_produk_jasa_dtps); 

sqlsrv_free_stmt( $stmt);

$spreadsheet_produk_jasa_dtps1 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./blank/sapto_produk_jasa_masyarakat_dtps.xlsx');
$worksheet_produk_jasa_dtps1 = $spreadsheet_produk_jasa_dtps1->getActiveSheet();

$worksheet_produk_jasa_dtps1->fromArray($data_karya_disitasi_dtps1, NULL, 'A2');

$writer_produk_jasa_dtps1 = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_produk_jasa_dtps1, 'Xlsx');
$writer_produk_jasa_dtps1->save('./raw/sapto_produk_jasa_dtps.xlsx');

$spreadsheet_produk_jasa_dtps1->disconnectWorksheets();
unset($spreadsheet_produk_jasa_dtps1);

// Load Format Baru
$spreadsheet_produk_jasa_dtps2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_produk_jasa_masyarakat_dtps.xlsx');
$worksheet_produk_jasa_dtps2 = $spreadsheet_produk_jasa_dtps2->getActiveSheet();

// Formasi Array SAPTO
$array_produk_jasa_dtps = $worksheet_produk_jasa_dtps2->toArray();
$data_produk_jasa_dtps = [];

foreach($worksheet_produk_jasa_dtps2->getRowIterator() as $row_id => $row) {
    if($worksheet_produk_jasa_dtps2->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['nama_dosen'] = $array_produk_jasa_dtps[$row_id-1][1];
            $item['nama_produk'] = $array_produk_jasa_dtps[$row_id-1][2];
			$item['desk_produk'] = $array_produk_jasa_dtps[$row_id-1][3];
			$item['bukti'] = $array_produk_jasa_dtps[$row_id-1][4];
            $data_produk_jasa_dtps[] = $item;
        }
    }
}

$spreadsheet_produk_jasa_dtps2->disconnectWorksheets();
unset($spreadsheet_produk_jasa_dtps2);


/**
 * Set ke Template SAPTO
 */

$spreadsheet_aps = \PhpOffice\PhpSpreadsheet\IOFactory::load('./result/sapto_aps9 (F).xlsx');


// Pengakuan/Rekognisi Dosen
$worksheet_aps8 = $spreadsheet_aps->getSheetByName('3b1');
$worksheet_aps8->fromArray($data_rekognisi, NULL, 'B10');

$highestRow_aps8 = $worksheet_aps8->getHighestRow();
$worksheet_aps8->getStyle('A10:H'.$highestRow_aps8)->applyFromArray($styleBorder);
$worksheet_aps8->getStyle('B10:H'.$highestRow_aps8)->applyFromArray($styleYellow);
$worksheet_aps8->getStyle('A10:A'.$highestRow_aps8)->applyFromArray($styleCenter);
$worksheet_aps8->getStyle('E10:H'.$highestRow_aps8)->applyFromArray($styleCenter);
$worksheet_aps8->getStyle('B10:H'.$highestRow_aps8)->getAlignment()->setWrapText(true);

foreach($worksheet_aps8->getRowDimensions() as $rd8) { 
    $rd8->setRowHeight(-1); 
}

for($row = 10; $row <= $highestRow_aps8; $row++) {
	$worksheet_aps8->setCellValue('A'.$row, $row-9);
}


// Penelitian DTPS
$worksheet_aps9 = $spreadsheet_aps->getSheetByName('3b2');
$worksheet_aps9->fromArray($penelitian_dtps_lokal, NULL, 'C6');
$worksheet_aps9->fromArray($penelitian_dtps2, NULL, 'C7');


// PkM DTPS
$worksheet_aps10 = $spreadsheet_aps->getSheetByName('3b3');
$worksheet_aps10->fromArray($pkm_dtps_lokal, NULL, 'C6');
$worksheet_aps10->fromArray($pkm_dtps2, NULL, 'C7');


// Publikasi Ilmiah DTPS
$worksheet_aps11 = $spreadsheet_aps->getSheetByName('3b4-1');

// Urutan bila semua jenis publikasi lengkap ada (kondisi ideal)
$worksheet_aps11->fromArray($publikasi_dtps_jurnal_seminar4, NULL, 'C7');
$worksheet_aps11->fromArray($publikasi_dtps_jurnal_seminar3, NULL, 'C8');
$worksheet_aps11->fromArray($publikasi_dtps_jurnal_seminar1, NULL, 'C9');
$worksheet_aps11->fromArray($publikasi_dtps_jurnal_seminar2, NULL, 'C10');
$worksheet_aps11->fromArray($publikasi_dtps_jurnal_seminar7, NULL, 'C11');
$worksheet_aps11->fromArray($publikasi_dtps_jurnal_seminar6, NULL, 'C12');
$worksheet_aps11->fromArray($publikasi_dtps_jurnal_seminar5, NULL, 'C13');
$worksheet_aps11->fromArray($publikasi_dtps_non_jurnal3, NULL, 'C14');
$worksheet_aps11->fromArray($publikasi_dtps_non_jurnal2, NULL, 'C15');
$worksheet_aps11->fromArray($publikasi_dtps_non_jurnal1, NULL, 'C16');


// Pagelaran/Pameran/Presentasi/Publikasi Ilmiah DTPS
$worksheet_aps42 = $spreadsheet_aps->getSheetByName('3b4-2');

// Urutan bila semua jenis publikasi lengkap ada (kondisi ideal)
$worksheet_aps42->fromArray($publikasi_dtps_jurnal_seminar4, NULL, 'C7');
$worksheet_aps42->fromArray($publikasi_dtps_jurnal_seminar3, NULL, 'C8');
$worksheet_aps42->fromArray($publikasi_dtps_jurnal_seminar1, NULL, 'C9');
$worksheet_aps42->fromArray($publikasi_dtps_jurnal_seminar2, NULL, 'C10');
$worksheet_aps42->fromArray($publikasi_dtps_jurnal_seminar7, NULL, 'C11');
$worksheet_aps42->fromArray($publikasi_dtps_jurnal_seminar6, NULL, 'C12');
$worksheet_aps42->fromArray($publikasi_dtps_jurnal_seminar5, NULL, 'C13');
$worksheet_aps42->fromArray($publikasi_dtps_non_jurnal3, NULL, 'C14');
$worksheet_aps42->fromArray($publikasi_dtps_non_jurnal2, NULL, 'C15');
$worksheet_aps42->fromArray($publikasi_dtps_non_jurnal1, NULL, 'C16');


// Luaran Penelitian/PkM Lainnya oleh DTPS - HKI (Paten, Paten Sederhana)
$worksheet_aps12 = $spreadsheet_aps->getSheetByName('3b5-1');
$worksheet_aps12->fromArray($data_luaran_dtps_hki_paten, NULL, 'B7');

$highestRow_aps12 = $worksheet_aps12->getHighestRow();

$worksheet_aps12->getStyle('A7:D'.$highestRow_aps12)->applyFromArray($styleBorder);
$worksheet_aps12->getStyle('B7:D'.$highestRow_aps12)->applyFromArray($styleYellow);
$worksheet_aps12->getStyle('A7:A'.$highestRow_aps12)->applyFromArray($styleCenter);
$worksheet_aps12->getStyle('C7:C'.$highestRow_aps12)->applyFromArray($styleCenter);
$worksheet_aps12->getStyle('B6:D'.$highestRow_aps12)->getAlignment()->setWrapText(true);

foreach($worksheet_aps12->getRowDimensions() as $rd12) { 
    $rd12->setRowHeight(-1); 
}

for($row = 7; $row <= $highestRow_aps12; $row++) {
	$worksheet_aps12->setCellValue('A'.$row, $row-6);
}


// Luaran Penelitian/PkM Lainnya oleh DTPS - HKI (Hak Cipta, Desain Produk Industri, dll.)
$worksheet_aps13 = $spreadsheet_aps->getSheetByName('3b5-2');
$worksheet_aps13->fromArray($data_luaran_dtps_hki_cipta, NULL, 'B7');

$highestRow_aps13 = $worksheet_aps13->getHighestRow();

$worksheet_aps13->getStyle('A7:D'.$highestRow_aps13)->applyFromArray($styleBorder);
$worksheet_aps13->getStyle('B7:D'.$highestRow_aps13)->applyFromArray($styleYellow);
$worksheet_aps13->getStyle('A7:A'.$highestRow_aps13)->applyFromArray($styleCenter);
$worksheet_aps13->getStyle('C7:C'.$highestRow_aps13)->applyFromArray($styleCenter);
$worksheet_aps13->getStyle('B6:D'.$highestRow_aps13)->getAlignment()->setWrapText(true);

foreach($worksheet_aps13->getRowDimensions() as $rd13) { 
    $rd13->setRowHeight(-1); 
}

for($row = 7; $row <= $highestRow_aps13; $row++) {
	$worksheet_aps13->setCellValue('A'.$row, $row-6);
}


// Luaran Penelitian/PkM Lainnya oleh DTPS - Teknologi Tepat Guna, Produk, Karya Seni, Rekayasa Sosial
$worksheet_aps14 = $spreadsheet_aps->getSheetByName('3b5-3');
$worksheet_aps14->fromArray($data_luaran_dtps_produk, NULL, 'B7');

$highestRow_aps14 = $worksheet_aps14->getHighestRow();

$worksheet_aps14->getStyle('A7:D'.$highestRow_aps14)->applyFromArray($styleBorder);
$worksheet_aps14->getStyle('B7:D'.$highestRow_aps14)->applyFromArray($styleYellow);
$worksheet_aps14->getStyle('A7:A'.$highestRow_aps14)->applyFromArray($styleCenter);
$worksheet_aps14->getStyle('C7:C'.$highestRow_aps14)->applyFromArray($styleCenter);
$worksheet_aps14->getStyle('B6:D'.$highestRow_aps14)->getAlignment()->setWrapText(true);

foreach($worksheet_aps14->getRowDimensions() as $rd14) { 
    $rd14->setRowHeight(-1); 
}

for($row = 7; $row <= $highestRow_aps14; $row++) {
	$worksheet_aps14->setCellValue('A'.$row, $row-6);
}


// Luaran Penelitian/PkM Lainnya oleh DTPS - Buku Ber-ISBN, Book Chapter
$worksheet_aps15 = $spreadsheet_aps->getSheetByName('3b5-4');
$worksheet_aps15->fromArray($data_luaran_dtps_bukuisbn, NULL, 'B7');

$highestRow_aps15 = $worksheet_aps15->getHighestRow();

$worksheet_aps15->getStyle('A7:D'.$highestRow_aps15)->applyFromArray($styleBorder);
$worksheet_aps15->getStyle('B7:D'.$highestRow_aps15)->applyFromArray($styleYellow);
$worksheet_aps15->getStyle('A7:A'.$highestRow_aps15)->applyFromArray($styleCenter);
$worksheet_aps15->getStyle('C7:C'.$highestRow_aps15)->applyFromArray($styleCenter);
$worksheet_aps15->getStyle('B6:D'.$highestRow_aps15)->getAlignment()->setWrapText(true);

foreach($worksheet_aps15->getRowDimensions() as $rd15) { 
    $rd15->setRowHeight(-1); 
}

for($row = 7; $row <= $highestRow_aps15; $row++) {
	$worksheet_aps15->setCellValue('A'.$row, $row-6);
}


// Karya Ilmiah DTPS yang Disitasi
$worksheet_aps16 = $spreadsheet_aps->getSheetByName('3b6');
$worksheet_aps16->fromArray($data_karya_disitasi_dtps, NULL, 'B6');

$highestRow_aps16 = $worksheet_aps16->getHighestRow();

$worksheet_aps16->getStyle('A6:D'.$highestRow_aps16)->applyFromArray($styleBorder);
$worksheet_aps16->getStyle('B6:D'.$highestRow_aps16)->applyFromArray($styleYellow);
$worksheet_aps16->getStyle('A6:A'.$highestRow_aps16)->applyFromArray($styleCenter);
$worksheet_aps16->getStyle('D6:D'.$highestRow_aps16)->applyFromArray($styleCenter);
$worksheet_aps16->getStyle('B6:D'.$highestRow_aps16)->getAlignment()->setWrapText(true);

foreach($worksheet_aps16->getRowDimensions() as $rd16) { 
    $rd16->setRowHeight(-1); 
}

for($row = 6; $row <= $highestRow_aps16; $row++) {
	$worksheet_aps16->setCellValue('A'.$row, $row-5);
}


// Produk/Jasa DTPS yang Diadopsi oleh Industri/Masyarakat
$worksheet_aps43 = $spreadsheet_aps->getSheetByName('3b7');
$worksheet_aps43->fromArray($data_produk_jasa_dtps, NULL, 'B6');

$highestRow_aps43 = $worksheet_aps43->getHighestRow();

$worksheet_aps43->getStyle('A6:E'.$highestRow_aps43)->applyFromArray($styleBorder);
$worksheet_aps43->getStyle('B6:E'.$highestRow_aps43)->applyFromArray($styleYellow);
$worksheet_aps43->getStyle('A6:A'.$highestRow_aps43)->applyFromArray($styleCenter);
$worksheet_aps43->getStyle('D6:E'.$highestRow_aps43)->applyFromArray($styleCenter);
$worksheet_aps43->getStyle('B6:E'.$highestRow_aps43)->getAlignment()->setWrapText(true);

foreach($worksheet_aps43->getRowDimensions() as $rd43) { 
    $rd43->setRowHeight(-1); 
}

for($row = 6; $row <= $highestRow_aps43; $row++) {
	$worksheet_aps43->setCellValue('A'.$row, $row-5);
}


$writer_aps = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_aps, 'Xlsx');
$writer_aps->save('./result/sapto_aps9 (F).xlsx');

$spreadsheet_aps->disconnectWorksheets();
unset($spreadsheet_aps);

?>