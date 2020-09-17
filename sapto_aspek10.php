<?php

/*

ASPEK 10: LUARAN DHARMA PENELITIAN & PkM

*/

require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

/* if ($argc < 2 )
{
    exit( "Usage: sapto_aspek10.php <id prodi sesuai di database>\n" );
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
 * Tabel 8.f.1 Publikasi Ilmiah Mahasiswa
 */


// Publikasi Ilmiah Mahasiswa Jurnal Seminar
$sql_publikasi_mhs_jurnal_seminar = "SELECT id, tingkat_publikasi, tahun, COUNT(DISTINCT judul_publikasi), id_prodi FROM akreditasi.sapto_publikasi_ilmiah_mhs_jurnal_seminar WHERE id_prodi = '".$nama_prodi."' GROUP BY tingkat_publikasi, tahun ORDER BY tingkat_publikasi";
$stmt = sqlsrv_query( $conn, $sql_publikasi_mhs_jurnal_seminar );
if( $stmt === false) {
    die( print_r( sqlsrv_errors(), true) );
}

$data_publikasi_mhs_jurnal_seminar1 = [];

while( $row = sqlsrv_fetch_array( $stmt, SQLSRV_FETCH_NUMERIC) ) {
	 $data_array_publikasi_mhs_jurnal_seminar[] = array(
		$row[0], $row[1], $row[2], $row[3], $row[4]
	);
}

$data_publikasi_mhs_jurnal_seminar1 = array_merge($data_publikasi_mhs_jurnal_seminar1, $data_array_publikasi_mhs_jurnal_seminar); 

sqlsrv_free_stmt( $stmt);

$spreadsheet_publikasi_mhs_jurnal_seminar1 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./blank/sapto_publikasi_mhs_jurnal_seminar.xlsx');
$worksheet_publikasi_mhs_jurnal_seminar1 = $spreadsheet_publikasi_mhs_jurnal_seminar1->getActiveSheet();

$worksheet_publikasi_mhs_jurnal_seminar1->fromArray($data_publikasi_mhs_jurnal_seminar1, NULL, 'A2');

$writer_publikasi_mhs_jurnal_seminar1 = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_publikasi_mhs_jurnal_seminar1, 'Xlsx');
$writer_publikasi_mhs_jurnal_seminar1->save('./raw/sapto_publikasi_mhs_jurnal_seminar.xlsx');

$spreadsheet_publikasi_mhs_jurnal_seminar1->disconnectWorksheets();
unset($spreadsheet_publikasi_mhs_jurnal_seminar1);

$spreadsheet_publikasi_mhs_jurnal_seminar = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_publikasi_mhs_jurnal_seminar.xlsx');
$worksheet_publikasi_mhs_jurnal_seminar = $spreadsheet_publikasi_mhs_jurnal_seminar->getActiveSheet();

$worksheet_publikasi_mhs_jurnal_seminar->insertNewColumnBefore('E', 4);

$worksheet_publikasi_mhs_jurnal_seminar->getCell('E1')->setValue('Jumlah TS-2');
$worksheet_publikasi_mhs_jurnal_seminar->getCell('F1')->setValue('Jumlah TS-1');
$worksheet_publikasi_mhs_jurnal_seminar->getCell('G1')->setValue('Jumlah TS');
$worksheet_publikasi_mhs_jurnal_seminar->getCell('H1')->setValue('Jumlah');

$highestRow_publikasi_mhs_jurnal_seminar = $worksheet_publikasi_mhs_jurnal_seminar->getHighestRow();

$publikasi_mhs_jurnal_seminar_ts = intval(date("Y"));
$publikasi_mhs_jurnal_seminar_ts1 = intval(date("Y", strtotime("-1 year")));
$publikasi_mhs_jurnal_seminar_ts2 = intval(date("Y", strtotime("-2 year")));

$worksheet_publikasi_mhs_jurnal_seminar->setAutoFilter('B1:I'.$highestRow_publikasi_mhs_jurnal_seminar);
$autoFilter_publikasi_mhs_jurnal_seminar = $worksheet_publikasi_mhs_jurnal_seminar->getAutoFilter();
$columnFilter_publikasi_mhs_jurnal_seminar = $autoFilter_publikasi_mhs_jurnal_seminar->getColumn('C');
$columnFilter_publikasi_mhs_jurnal_seminar->setFilterType(
    \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_FILTER
);
$columnFilter_publikasi_mhs_jurnal_seminar->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        $publikasi_mhs_jurnal_seminar_ts2
    );
$columnFilter_publikasi_mhs_jurnal_seminar->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        $publikasi_mhs_jurnal_seminar_ts1
    );
$columnFilter_publikasi_mhs_jurnal_seminar->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        $publikasi_mhs_jurnal_seminar_ts
    );

$autoFilter_publikasi_mhs_jurnal_seminar->showHideRows();

for($row = 2;$row <= $highestRow_publikasi_mhs_jurnal_seminar; $row++) {
	$worksheet_publikasi_mhs_jurnal_seminar->setCellValue('E'.$row, '=IF(C'.$row.'='.$ts2.',D'.$row.',0)');
	$worksheet_publikasi_mhs_jurnal_seminar->getCell('E'.$row)->getStyle()->setQuotePrefix(true);
	
	$worksheet_publikasi_mhs_jurnal_seminar->setCellValue('F'.$row, '=IF(C'.$row.'='.$ts1.',D'.$row.',0)');
	$worksheet_publikasi_mhs_jurnal_seminar->getCell('F'.$row)->getStyle()->setQuotePrefix(true);
	
	$worksheet_publikasi_mhs_jurnal_seminar->setCellValue('G'.$row, '=IF(C'.$row.'='.$ts.',D'.$row.',0)');
	$worksheet_publikasi_mhs_jurnal_seminar->getCell('G'.$row)->getStyle()->setQuotePrefix(true);
}

$writer_publikasi_mhs_jurnal_seminar = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_publikasi_mhs_jurnal_seminar, 'Xls');
$writer_publikasi_mhs_jurnal_seminar->save('./formatted/sapto_publikasi_mhs_jurnal_seminar (F).xls');

$spreadsheet_publikasi_mhs_jurnal_seminar->disconnectWorksheets();
unset($spreadsheet_publikasi_mhs_jurnal_seminar);

// Load Format Baru
$spreadsheet_publikasi_mhs_jurnal_seminar2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_publikasi_mhs_jurnal_seminar (F).xls');
$worksheet_publikasi_mhs_jurnal_seminar2 = $spreadsheet_publikasi_mhs_jurnal_seminar2->getActiveSheet();

// Formasi Array SAPTO
$array_publikasi_mhs_jurnal_seminar = $worksheet_publikasi_mhs_jurnal_seminar2->toArray();
$data_publikasi_mhs_jurnal_seminar = [];

foreach($worksheet_publikasi_mhs_jurnal_seminar2->getRowIterator() as $row_id => $row) {
    if($worksheet_publikasi_mhs_jurnal_seminar2->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['jenis_publikasi'] = $array_publikasi_mhs_jurnal_seminar[$row_id-1][1];
            $item['judul_ts2'] = $array_publikasi_mhs_jurnal_seminar[$row_id-1][4];
			$item['judul_ts1'] = $array_publikasi_mhs_jurnal_seminar[$row_id-1][5];
			$item['judul_ts'] = $array_publikasi_mhs_jurnal_seminar[$row_id-1][6];
			$item['jumlah'] = $array_publikasi_mhs_jurnal_seminar[$row_id-1][7];
            $data_publikasi_mhs_jurnal_seminar[] = $item;
        }
    }
}

$worksheet_publikasi_mhs_jurnal_seminar3 = new \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet($spreadsheet_publikasi_mhs_jurnal_seminar2, 'Sheet 2');
$spreadsheet_publikasi_mhs_jurnal_seminar2->addSheet($worksheet_publikasi_mhs_jurnal_seminar3);

$worksheet_publikasi_mhs_jurnal_seminar3 = $spreadsheet_publikasi_mhs_jurnal_seminar2->getSheetByName('Sheet 2');
$worksheet_publikasi_mhs_jurnal_seminar3->fromArray($data_publikasi_mhs_jurnal_seminar, NULL, 'A1');

$highestRow_publikasi_mhs_jurnal_seminar3 = $worksheet_publikasi_mhs_jurnal_seminar3->getHighestRow();

$row_jumlah = -1;

for($group = 1;$group <= 7; $group++) {
	
	$row_jumlah += 2;
	$baris_awal = $row_jumlah;
	$judul_publikasi_jurnal_ts2 = $worksheet_publikasi_mhs_jurnal_seminar3->getCell('B'.($row_jumlah))->getValue();
	$judul_publikasi_jurnal_ts1 = $worksheet_publikasi_mhs_jurnal_seminar3->getCell('C'.($row_jumlah))->getValue();
	$judul_publikasi_jurnal_ts = $worksheet_publikasi_mhs_jurnal_seminar3->getCell('D'.($row_jumlah))->getValue();	
	
	for($row = $row_jumlah;$row <= ($highestRow_publikasi_mhs_jurnal_seminar3+1); $row++) {
		if($worksheet_publikasi_mhs_jurnal_seminar3->getCell('A'.$row)->getValue() == $worksheet_publikasi_mhs_jurnal_seminar3->getCell('A'.($row+1))->getValue()) {
			$judul_publikasi_jurnal_ts2 += $worksheet_publikasi_mhs_jurnal_seminar3->getCell('B'.($row+1))->getValue();
			$judul_publikasi_jurnal_ts1 += $worksheet_publikasi_mhs_jurnal_seminar3->getCell('C'.($row+1))->getValue();
			$judul_publikasi_jurnal_ts += $worksheet_publikasi_mhs_jurnal_seminar3->getCell('D'.($row+1))->getValue();
			
			$row_jumlah++;
		} else {
			break;
		}
	}
	
	$worksheet_publikasi_mhs_jurnal_seminar3->insertNewRowBefore(($row_jumlah+1), 1);
	$worksheet_publikasi_mhs_jurnal_seminar3->setCellValue('B'.($row_jumlah+1), $judul_publikasi_jurnal_ts2);
	$worksheet_publikasi_mhs_jurnal_seminar3->setCellValue('C'.($row_jumlah+1), $judul_publikasi_jurnal_ts1);
	$worksheet_publikasi_mhs_jurnal_seminar3->setCellValue('D'.($row_jumlah+1), $judul_publikasi_jurnal_ts);
	
	${"publikasi_mhs_jurnal_seminar".$group} = $worksheet_publikasi_mhs_jurnal_seminar3->rangeToArray('B'.($row_jumlah+1).':D'.($row_jumlah+1), NULL, TRUE, TRUE, TRUE);
}

$writer_publikasi_mhs_jurnal_seminar2 = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_publikasi_mhs_jurnal_seminar2, 'Xls');
$writer_publikasi_mhs_jurnal_seminar2->save('./formatted/sapto_publikasi_mhs_jurnal_seminar (F).xls');

$spreadsheet_publikasi_mhs_jurnal_seminar2->disconnectWorksheets();
unset($spreadsheet_publikasi_mhs_jurnal_seminar2);


// Publikasi Ilmiah Mahasiswa Selain Jurnal Seminar
$sql_publikasi_mhs_non_jurnal = "SELECT id, tingkat_publikasi, tahun, COUNT(DISTINCT judul_publikasi), id_prodi FROM akreditasi.sapto_publikasi_ilmiah_mhs_selain_jurnal_seminar WHERE id_prodi = '".$nama_prodi."' GROUP BY tingkat_publikasi, tahun ORDER BY tingkat_publikasi";
$stmt = sqlsrv_query( $conn, $sql_publikasi_mhs_non_jurnal );
if( $stmt === false) {
    die( print_r( sqlsrv_errors(), true) );
}

$data_publikasi_mhs_non_jurnal1 = [];

while( $row = sqlsrv_fetch_array( $stmt, SQLSRV_FETCH_NUMERIC) ) {
	 $data_array_publikasi_mhs_non_jurnal[] = array(
		$row[0], $row[1], $row[2], $row[3], $row[4]
	);
}

$data_publikasi_mhs_non_jurnal1 = array_merge($data_publikasi_mhs_non_jurnal1, $data_array_publikasi_mhs_non_jurnal); 

sqlsrv_free_stmt( $stmt);

$spreadsheet_publikasi_mhs_non_jurnal1 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./blank/sapto_publikasi_mhs_non_jurnal.xlsx');
$worksheet_publikasi_mhs_non_jurnal1 = $spreadsheet_publikasi_mhs_non_jurnal1->getActiveSheet();

$worksheet_publikasi_mhs_non_jurnal1->fromArray($data_publikasi_mhs_non_jurnal1, NULL, 'A2');

$writer_publikasi_mhs_non_jurnal1 = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_publikasi_mhs_non_jurnal1, 'Xlsx');
$writer_publikasi_mhs_non_jurnal1->save('./raw/sapto_publikasi_mhs_non_jurnal.xlsx');

$spreadsheet_publikasi_mhs_non_jurnal1->disconnectWorksheets();
unset($spreadsheet_publikasi_mhs_non_jurnal1);

$spreadsheet_publikasi_mhs_non_jurnal = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_publikasi_mhs_non_jurnal.xlsx');
$worksheet_publikasi_mhs_non_jurnal = $spreadsheet_publikasi_mhs_non_jurnal->getActiveSheet();

$worksheet_publikasi_mhs_non_jurnal->insertNewColumnBefore('E', 4);

$worksheet_publikasi_mhs_non_jurnal->getCell('E1')->setValue('Jumlah TS-2');
$worksheet_publikasi_mhs_non_jurnal->getCell('F1')->setValue('Jumlah TS-1');
$worksheet_publikasi_mhs_non_jurnal->getCell('G1')->setValue('Jumlah TS');
$worksheet_publikasi_mhs_non_jurnal->getCell('H1')->setValue('Jumlah');

$highestRow_publikasi_mhs_non_jurnal = $worksheet_publikasi_mhs_non_jurnal->getHighestRow();

$publikasi_mhs_non_jurnal_ts = intval(date("Y"));
$publikasi_mhs_non_jurnal_ts1 = intval(date("Y", strtotime("-1 year")));
$publikasi_mhs_non_jurnal_ts2 = intval(date("Y", strtotime("-2 year")));

$worksheet_publikasi_mhs_non_jurnal->setAutoFilter('B1:I'.$highestRow_publikasi_mhs_non_jurnal);
$autoFilter_publikasi_mhs_non_jurnal = $worksheet_publikasi_mhs_non_jurnal->getAutoFilter();
$columnFilter_publikasi_mhs_non_jurnal = $autoFilter_publikasi_mhs_non_jurnal->getColumn('C');
$columnFilter_publikasi_mhs_non_jurnal->setFilterType(
    \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_FILTER
);
$columnFilter_publikasi_mhs_non_jurnal->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        $publikasi_mhs_non_jurnal_ts2
    );
$columnFilter_publikasi_mhs_non_jurnal->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        $publikasi_mhs_non_jurnal_ts1
    );
$columnFilter_publikasi_mhs_non_jurnal->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        $publikasi_mhs_non_jurnal_ts
    );

$autoFilter_publikasi_mhs_non_jurnal->showHideRows();

for($row = 2;$row <= $highestRow_publikasi_mhs_non_jurnal; $row++) {
	$worksheet_publikasi_mhs_non_jurnal->setCellValue('E'.$row, '=IF(C'.$row.'='.$ts2.',D'.$row.',0)');
	$worksheet_publikasi_mhs_non_jurnal->getCell('E'.$row)->getStyle()->setQuotePrefix(true);
	
	$worksheet_publikasi_mhs_non_jurnal->setCellValue('F'.$row, '=IF(C'.$row.'='.$ts1.',D'.$row.',0)');
	$worksheet_publikasi_mhs_non_jurnal->getCell('F'.$row)->getStyle()->setQuotePrefix(true);
	
	$worksheet_publikasi_mhs_non_jurnal->setCellValue('G'.$row, '=IF(C'.$row.'='.$ts.',D'.$row.',0)');
	$worksheet_publikasi_mhs_non_jurnal->getCell('G'.$row)->getStyle()->setQuotePrefix(true);
}

$writer_publikasi_mhs_non_jurnal = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_publikasi_mhs_non_jurnal, 'Xls');
$writer_publikasi_mhs_non_jurnal->save('./formatted/sapto_publikasi_mhs_non_jurnal (F).xls');

$spreadsheet_publikasi_mhs_non_jurnal->disconnectWorksheets();
unset($spreadsheet_publikasi_mhs_non_jurnal);

// Load Format Baru
$spreadsheet_publikasi_mhs_non_jurnal2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_publikasi_mhs_non_jurnal (F).xls');
$worksheet_publikasi_mhs_non_jurnal2 = $spreadsheet_publikasi_mhs_non_jurnal2->getActiveSheet();

// Formasi Array SAPTO
$array_publikasi_mhs_non_jurnal = $worksheet_publikasi_mhs_non_jurnal2->toArray();
$data_publikasi_mhs_non_jurnal = [];

foreach($worksheet_publikasi_mhs_non_jurnal2->getRowIterator() as $row_id => $row) {
    if($worksheet_publikasi_mhs_non_jurnal2->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['jenis_publikasi'] = $array_publikasi_mhs_non_jurnal[$row_id-1][1];
            $item['judul_ts2'] = $array_publikasi_mhs_non_jurnal[$row_id-1][4];
			$item['judul_ts1'] = $array_publikasi_mhs_non_jurnal[$row_id-1][5];
			$item['judul_ts'] = $array_publikasi_mhs_non_jurnal[$row_id-1][6];
			$item['jumlah'] = $array_publikasi_mhs_non_jurnal[$row_id-1][7];
            $data_publikasi_mhs_non_jurnal[] = $item;
        }
    }
}

$worksheet_publikasi_mhs_non_jurnal3 = new \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet($spreadsheet_publikasi_mhs_non_jurnal2, 'Sheet 2');
$spreadsheet_publikasi_mhs_non_jurnal2->addSheet($worksheet_publikasi_mhs_non_jurnal3);

$worksheet_publikasi_mhs_non_jurnal3 = $spreadsheet_publikasi_mhs_non_jurnal2->getSheetByName('Sheet 2');
$worksheet_publikasi_mhs_non_jurnal3->fromArray($data_publikasi_mhs_non_jurnal, NULL, 'A1');

$highestRow_publikasi_mhs_non_jurnal3 = $worksheet_publikasi_mhs_non_jurnal3->getHighestRow();

$row_jumlah = -1;

for($group = 1;$group <= 3; $group++) {
	
	$row_jumlah += 2;
	$baris_awal = $row_jumlah;
	$judul_publikasi_non_jurnal_ts2 = $worksheet_publikasi_mhs_non_jurnal3->getCell('B'.($row_jumlah))->getValue();
	$judul_publikasi_non_jurnal_ts1 = $worksheet_publikasi_mhs_non_jurnal3->getCell('C'.($row_jumlah))->getValue();
	$judul_publikasi_non_jurnal_ts = $worksheet_publikasi_mhs_non_jurnal3->getCell('D'.($row_jumlah))->getValue();	
	
	for($row = $row_jumlah;$row <= ($highestRow_publikasi_mhs_non_jurnal3+1); $row++) {
		if($worksheet_publikasi_mhs_non_jurnal3->getCell('A'.$row)->getValue() == $worksheet_publikasi_mhs_non_jurnal3->getCell('A'.($row+1))->getValue()) {
			$judul_publikasi_non_jurnal_ts2 += $worksheet_publikasi_mhs_non_jurnal3->getCell('B'.($row+1))->getValue();
			$judul_publikasi_non_jurnal_ts1 += $worksheet_publikasi_mhs_non_jurnal3->getCell('C'.($row+1))->getValue();
			$judul_publikasi_non_jurnal_ts += $worksheet_publikasi_mhs_non_jurnal3->getCell('D'.($row+1))->getValue();
			
			$row_jumlah++;
		} else {
			break;
		}
	}
	
	$worksheet_publikasi_mhs_non_jurnal3->insertNewRowBefore(($row_jumlah+1), 1);
	$worksheet_publikasi_mhs_non_jurnal3->setCellValue('B'.($row_jumlah+1), $judul_publikasi_non_jurnal_ts2);
	$worksheet_publikasi_mhs_non_jurnal3->setCellValue('C'.($row_jumlah+1), $judul_publikasi_non_jurnal_ts1);
	$worksheet_publikasi_mhs_non_jurnal3->setCellValue('D'.($row_jumlah+1), $judul_publikasi_non_jurnal_ts);
	
	${"publikasi_mhs_non_jurnal".$group} = $worksheet_publikasi_mhs_non_jurnal3->rangeToArray('B'.($row_jumlah+1).':D'.($row_jumlah+1), NULL, TRUE, TRUE, TRUE);
}

$writer_publikasi_mhs_non_jurnal2 = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_publikasi_mhs_non_jurnal2, 'Xls');
$writer_publikasi_mhs_non_jurnal2->save('./formatted/sapto_publikasi_mhs_non_jurnal (F).xls');

$spreadsheet_publikasi_mhs_non_jurnal2->disconnectWorksheets();
unset($spreadsheet_publikasi_mhs_non_jurnal2);



/**
 * Tabel 8.f.2 Karya Ilmiah Mahasiswa yang Disitasi
 */

$sql_karya_disitasi_mhs = "SELECT nama, title, jml_sitasi, id_prodi FROM akreditasi.sapto_karya_ilmiah_disitasi_mhs WHERE id_prodi = '".$nama_prodi."'";
$stmt = sqlsrv_query( $conn, $sql_karya_disitasi_mhs );
if( $stmt === false) {
    die( print_r( sqlsrv_errors(), true) );
}

$data_karya_disitasi_mhs1 = [];

while( $row = sqlsrv_fetch_array( $stmt, SQLSRV_FETCH_NUMERIC) ) {
	 $data_array_karya_disitasi_mhs[] = array(
		$row[0], $row[1], $row[2], $row[3]
	);
}

$data_karya_disitasi_mhs1 = array_merge($data_karya_disitasi_mhs1, $data_array_karya_disitasi_mhs); 

sqlsrv_free_stmt( $stmt);

$spreadsheet_karya_disitasi_mhs1 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./blank/sapto_karya_ilmiah_disitasi_mhs.xlsx');
$worksheet_karya_disitasi_mhs1 = $spreadsheet_karya_disitasi_mhs1->getActiveSheet();

$worksheet_karya_disitasi_mhs1->fromArray($data_karya_disitasi_mhs1, NULL, 'A2');

$writer_karya_disitasi_mhs1 = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_karya_disitasi_mhs1, 'Xlsx');
$writer_karya_disitasi_mhs1->save('./raw/sapto_karya_ilmiah_disitasi_mhs.xlsx');

$spreadsheet_karya_disitasi_mhs1->disconnectWorksheets();
unset($spreadsheet_karya_disitasi_mhs1);

// Load Format Baru
$spreadsheet_karya_disitasi_mhs2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_karya_ilmiah_disitasi_mhs.xlsx');
$worksheet_karya_disitasi_mhs2 = $spreadsheet_karya_disitasi_mhs2->getActiveSheet();

// Formasi Array SAPTO
$array_karya_disitasi_mhs = $worksheet_karya_disitasi_mhs2->toArray();
$data_karya_disitasi_mhs = [];

foreach($worksheet_karya_disitasi_mhs2->getRowIterator() as $row_id => $row) {
    if($worksheet_karya_disitasi_mhs2->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['nama_mahasiswa'] = $array_karya_disitasi_mhs[$row_id-1][0];
            $item['judul_artikel'] = $array_karya_disitasi_mhs[$row_id-1][1];
			$item['jumlah_sitasi'] = $array_karya_disitasi_mhs[$row_id-1][2];
            $data_karya_disitasi_mhs[] = $item;
        }
    }
}

$spreadsheet_karya_disitasi_mhs2->disconnectWorksheets();
unset($spreadsheet_karya_disitasi_mhs2);



/**
 * Tabel 8.f.3 Produk/Jasa Mahasiswa yang Diadopsi oleh Industri/Masyarakat
 */

$sql_produk_jasa_mhs = "SELECT id, nama, nama_produk_jasa, deskripsi, bukti, id_prodi FROM akreditasi.sapto_produk_jasa_masyarakat_mhs WHERE prodi_id = '".$nama_prodi."'";
$stmt = sqlsrv_query( $conn, $sql_produk_jasa_mhs );
if( $stmt === false) {
    die( print_r( sqlsrv_errors(), true) );
}

$data_produk_jasa_mhs1 = [];

while( $row = sqlsrv_fetch_array( $stmt, SQLSRV_FETCH_NUMERIC) ) {
	 $data_array_produk_jasa_mhs[] = array(
		$row[0], $row[1], $row[2], $row[3], $row[4], $row[5]
	);
}

$data_produk_jasa_mhs1 = array_merge($data_produk_jasa_mhs1, $data_array_produk_jasa_mhs); 

sqlsrv_free_stmt( $stmt);

$spreadsheet_produk_jasa_mhs1 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./blank/sapto_produk_jasa_masyarakat_mhs.xlsx');
$worksheet_produk_jasa_mhs1 = $spreadsheet_produk_jasa_mhs1->getActiveSheet();

$worksheet_produk_jasa_mhs1->fromArray($data_karya_disitasi_mhs1, NULL, 'A2');

$writer_produk_jasa_mhs1 = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_produk_jasa_mhs1, 'Xlsx');
$writer_produk_jasa_mhs1->save('./raw/sapto_produk_jasa_mhs.xlsx');

$spreadsheet_produk_jasa_mhs1->disconnectWorksheets();
unset($spreadsheet_produk_jasa_mhs1);

// Load Format Baru
$spreadsheet_produk_jasa_mhs2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_produk_jasa_masyarakat_mhs.xlsx');
$worksheet_produk_jasa_mhs2 = $spreadsheet_produk_jasa_mhs2->getActiveSheet();

// Formasi Array SAPTO
$array_produk_jasa_mhs = $worksheet_produk_jasa_mhs2->toArray();
$data_produk_jasa_mhs = [];

foreach($worksheet_produk_jasa_mhs2->getRowIterator() as $row_id => $row) {
    if($worksheet_produk_jasa_mhs2->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['nama_mahasiswa'] = $array_produk_jasa_mhs[$row_id-1][1];
            $item['nama_produk'] = $array_produk_jasa_mhs[$row_id-1][2];
			$item['desk_produk'] = $array_produk_jasa_mhs[$row_id-1][3];
			$item['bukti'] = $array_produk_jasa_mhs[$row_id-1][4];
            $data_produk_jasa_mhs[] = $item;
        }
    }
}

$spreadsheet_produk_jasa_mhs2->disconnectWorksheets();
unset($spreadsheet_produk_jasa_mhs2);


/**
 * Tabel 8.f.4 Luaran Penelitian yang Dihasilkan Mahasiswa - HKI (Paten, Paten Sederhana)
 */

$sql_luaran_mhs_hki_paten = "SELECT id, judul, tahun, id_prodi FROM akreditasi.sapto_luaran_penelitian_mhs_hki WHERE id_prodi = '".$nama_prodi."' AND kategori = 'paten'";
$stmt = sqlsrv_query( $conn, $sql_luaran_mhs_hki_paten );
if( $stmt === false) {
    die( print_r( sqlsrv_errors(), true) );
}

$data_luaran_mhs_hki_paten1 = [];

while( $row = sqlsrv_fetch_array( $stmt, SQLSRV_FETCH_NUMERIC) ) {
	 $data_array_luaran_mhs_hki_paten[] = array(
		$row[0], $row[1], $row[2], $row[3]
	);
}

$data_luaran_mhs_hki_paten1 = array_merge($data_luaran_mhs_hki_paten1, $data_array_luaran_mhs_hki_paten); 

sqlsrv_free_stmt( $stmt);

$spreadsheet_luaran_mhs_hki_paten1 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./blank/sapto_luaran_penelitian_mhs_hki_paten.xlsx');
$worksheet_luaran_mhs_hki_paten1 = $spreadsheet_luaran_mhs_hki_paten1->getActiveSheet();

$worksheet_luaran_mhs_hki_paten1->fromArray($data_luaran_mhs_hki_paten1, NULL, 'A2');

$writer_luaran_mhs_hki_paten1 = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_luaran_mhs_hki_paten1, 'Xlsx');
$writer_luaran_mhs_hki_paten1->save('./raw/sapto_luaran_penelitian_mhs_hki_paten.xlsx');

$spreadsheet_luaran_mhs_hki_paten1->disconnectWorksheets();
unset($spreadsheet_luaran_mhs_hki_paten1);

$spreadsheet_luaran_mhs_hki_paten = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_luaran_penelitian_mhs_hki_paten.xlsx');

$worksheet_luaran_mhs_hki_paten = $spreadsheet_luaran_mhs_hki_paten->getActiveSheet();

$worksheet_luaran_mhs_hki_paten->insertNewColumnBefore('D', 1);
$worksheet_luaran_mhs_hki_paten->getCell('D1')->setValue('Keterangan');

$writer_luaran_mhs_hki_paten = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_luaran_mhs_hki_paten, 'Xlsx');
$writer_luaran_mhs_hki_paten->save('./formatted/sapto_luaran_penelitian_mhs_hki_paten (F).xlsx');

$spreadsheet_luaran_mhs_hki_paten->disconnectWorksheets();
unset($spreadsheet_luaran_mhs_hki_paten);

// Load Format Baru
$spreadsheet_luaran_mhs_hki_paten2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_luaran_penelitian_mhs_hki_paten (F).xlsx');
$worksheet_luaran_mhs_hki_paten2 = $spreadsheet_luaran_mhs_hki_paten2->getActiveSheet();

// Formasi Array SAPTO
$array_luaran_mhs_hki_paten = $worksheet_luaran_mhs_hki_paten2->toArray();
$data_luaran_mhs_hki_paten = [];

foreach($worksheet_luaran_mhs_hki_paten2->getRowIterator() as $row_id => $row) {
    if($worksheet_luaran_mhs_hki_paten2->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['luaran_penelitian'] = $array_luaran_mhs_hki_paten[$row_id-1][1];
            $item['tahun'] = $array_luaran_mhs_hki_paten[$row_id-1][2];
			$item['keterangan'] = $array_luaran_mhs_hki_paten[$row_id-1][3];
            $data_luaran_mhs_hki_paten[] = $item;
        }
    }
}

$spreadsheet_luaran_mhs_hki_paten2->disconnectWorksheets();
unset($spreadsheet_luaran_mhs_hki_paten2);


/**
 * Tabel 8.f.4 Luaran Penelitian yang Dihasilkan Mahasiswa - HKI (Hak Cipta, Desain Produk Industri, dll.)
 */

$sql_luaran_mhs_hki_cipta = "SELECT id, judul, tahun, id_prodi FROM akreditasi.sapto_luaran_penelitian_mhs_hki WHERE id_prodi = '".$nama_prodi."' AND kategori = 'hak cipta'";
$stmt = sqlsrv_query( $conn, $sql_luaran_mhs_hki_cipta );
if( $stmt === false) {
    die( print_r( sqlsrv_errors(), true) );
}

$data_luaran_mhs_hki_cipta1 = [];

while( $row = sqlsrv_fetch_array( $stmt, SQLSRV_FETCH_NUMERIC) ) {
	 $data_array_luaran_mhs_hki_cipta[] = array(
		$row[0], $row[1], $row[2], $row[3]
	);
}

$data_luaran_mhs_hki_cipta1 = array_merge($data_luaran_mhs_hki_cipta1, $data_array_luaran_mhs_hki_cipta); 

sqlsrv_free_stmt( $stmt);

$spreadsheet_luaran_mhs_hki_cipta1 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./blank/sapto_luaran_penelitian_mhs_hki_cipta.xlsx');
$worksheet_luaran_mhs_hki_cipta1 = $spreadsheet_luaran_mhs_hki_cipta1->getActiveSheet();

$worksheet_luaran_mhs_hki_cipta1->fromArray($data_luaran_mhs_hki_cipta1, NULL, 'A2');

$writer_luaran_mhs_hki_cipta1 = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_luaran_mhs_hki_cipta1, 'Xlsx');
$writer_luaran_mhs_hki_cipta1->save('./raw/sapto_luaran_penelitian_mhs_hki_cipta.xlsx');

$spreadsheet_luaran_mhs_hki_cipta1->disconnectWorksheets();
unset($spreadsheet_luaran_mhs_hki_cipta1);

$spreadsheet_luaran_mhs_hki_cipta = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_luaran_penelitian_mhs_hki_cipta.xlsx');

$worksheet_luaran_mhs_hki_cipta = $spreadsheet_luaran_mhs_hki_cipta->getActiveSheet();

$worksheet_luaran_mhs_hki_cipta->insertNewColumnBefore('D', 1);
$worksheet_luaran_mhs_hki_cipta->getCell('D1')->setValue('Keterangan');

$writer_luaran_mhs_hki_cipta = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_luaran_mhs_hki_cipta, 'Xlsx');
$writer_luaran_mhs_hki_cipta->save('./formatted/sapto_luaran_penelitian_mhs_hki_cipta (F).xlsx');

$spreadsheet_luaran_mhs_hki_cipta->disconnectWorksheets();
unset($spreadsheet_luaran_mhs_hki_cipta);

// Load Format Baru
$spreadsheet_luaran_mhs_hki_cipta2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_luaran_penelitian_mhs_hki_cipta (F).xlsx');
$worksheet_luaran_mhs_hki_cipta2 = $spreadsheet_luaran_mhs_hki_cipta2->getActiveSheet();

// Formasi Array SAPTO
$array_luaran_mhs_hki_cipta = $worksheet_luaran_mhs_hki_cipta2->toArray();
$data_luaran_mhs_hki_cipta = [];

foreach($worksheet_luaran_mhs_hki_cipta2->getRowIterator() as $row_id => $row) {
    if($worksheet_luaran_mhs_hki_cipta2->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['luaran_penelitian'] = $array_luaran_mhs_hki_cipta[$row_id-1][1];
            $item['tahun'] = $array_luaran_mhs_hki_cipta[$row_id-1][2];
			$item['keterangan'] = $array_luaran_mhs_hki_cipta[$row_id-1][3];
            $data_luaran_mhs_hki_cipta[] = $item;
        }
    }
}

$spreadsheet_luaran_mhs_hki_cipta2->disconnectWorksheets();
unset($spreadsheet_luaran_mhs_hki_cipta2);


/**
 * Tabel 8.f.4 Luaran Penelitian yang Dihasilkan Mahasiswa - Teknologi Tepat Guna, Produk, Karya Seni, Rekayasa Sosial
 */

$sql_luaran_mhs_produk = "SELECT id, judul, tahun, id_prodi FROM akreditasi.sapto_luaran_penelitian_mhs_produk WHERE id_prodi = '".$nama_prodi."'";
$stmt = sqlsrv_query( $conn, $sql_luaran_mhs_produk );
if( $stmt === false) {
    die( print_r( sqlsrv_errors(), true) );
}

$data_luaran_mhs_produk1 = [];

while( $row = sqlsrv_fetch_array( $stmt, SQLSRV_FETCH_NUMERIC) ) {
	 $data_array_luaran_mhs_produk[] = array(
		$row[0], $row[1], $row[2], $row[3]
	);
}

$data_luaran_mhs_produk1 = array_merge($data_luaran_mhs_produk1, $data_array_luaran_mhs_produk); 

sqlsrv_free_stmt( $stmt);

$spreadsheet_luaran_mhs_produk1 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./blank/sapto_luaran_penelitian_mhs_produk.xlsx');
$worksheet_luaran_mhs_produk1 = $spreadsheet_luaran_mhs_produk1->getActiveSheet();

$worksheet_luaran_mhs_produk1->fromArray($data_luaran_mhs_produk1, NULL, 'A2');

$writer_luaran_mhs_produk1 = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_luaran_mhs_produk1, 'Xlsx');
$writer_luaran_mhs_produk1->save('./raw/sapto_luaran_penelitian_mhs_produk.xlsx');

$spreadsheet_luaran_mhs_produk1->disconnectWorksheets();
unset($spreadsheet_luaran_mhs_produk1);

$spreadsheet_luaran_mhs_produk = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_luaran_penelitian_mhs_produk.xlsx');

$worksheet_luaran_mhs_produk = $spreadsheet_luaran_mhs_produk->getActiveSheet();

$worksheet_luaran_mhs_produk->insertNewColumnBefore('D', 1);
$worksheet_luaran_mhs_produk->getCell('D1')->setValue('Keterangan');

$writer_luaran_mhs_produk = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_luaran_mhs_produk, 'Xlsx');
$writer_luaran_mhs_produk->save('./formatted/sapto_luaran_penelitian_mhs_produk (F).xlsx');

$spreadsheet_luaran_mhs_produk->disconnectWorksheets();
unset($spreadsheet_luaran_mhs_produk);

// Load Format Baru
$spreadsheet_luaran_mhs_produk2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_luaran_penelitian_mhs_produk (F).xlsx');
$worksheet_luaran_mhs_produk2 = $spreadsheet_luaran_mhs_produk2->getActiveSheet();

// Formasi Array SAPTO
$array_luaran_mhs_produk = $worksheet_luaran_mhs_produk2->toArray();
$data_luaran_mhs_produk = [];

foreach($worksheet_luaran_mhs_produk2->getRowIterator() as $row_id => $row) {
    if($worksheet_luaran_mhs_produk2->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['luaran_penelitian'] = $array_luaran_mhs_produk[$row_id-1][1];
            $item['tahun'] = $array_luaran_mhs_produk[$row_id-1][2];
			$item['keterangan'] = $array_luaran_mhs_produk[$row_id-1][3];
            $data_luaran_mhs_produk[] = $item;
        }
    }
}

$spreadsheet_luaran_mhs_produk2->disconnectWorksheets();
unset($spreadsheet_luaran_mhs_produk2);


/**
 * Tabel 8.f.4 Luaran Penelitian yang Dihasilkan Mahasiswa - Buku ber-ISBN, Book Chapter
 */

$sql_luaran_mhs_bukuisbn = "SELECT id, judul, tahun_terbit, keterangan, id_prodi FROM akreditasi.sapto_luaran_penelitian_mhs_bukuisbn WHERE id_prodi = '".$nama_prodi."'";
$stmt = sqlsrv_query( $conn, $sql_luaran_mhs_bukuisbn );
if( $stmt === false) {
    die( print_r( sqlsrv_errors(), true) );
}

$data_luaran_mhs_bukuisbn1 = [];

while( $row = sqlsrv_fetch_array( $stmt, SQLSRV_FETCH_NUMERIC) ) {
	 $data_array_luaran_mhs_bukuisbn[] = array(
		$row[0], $row[1], $row[2], $row[3]
	);
}

$data_luaran_mhs_bukuisbn1 = array_merge($data_luaran_mhs_bukuisbn1, $data_array_luaran_mhs_bukuisbn); 

sqlsrv_free_stmt( $stmt);

$spreadsheet_luaran_mhs_bukuisbn1 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./blank/sapto_luaran_penelitian_mhs_bukuisbn.xlsx');
$worksheet_luaran_mhs_bukuisbn1 = $spreadsheet_luaran_mhs_bukuisbn1->getActiveSheet();

$worksheet_luaran_mhs_bukuisbn1->fromArray($data_luaran_mhs_bukuisbn1, NULL, 'A2');

$writer_luaran_mhs_bukuisbn1 = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_luaran_mhs_bukuisbn1, 'Xlsx');
$writer_luaran_mhs_bukuisbn1->save('./raw/sapto_luaran_penelitian_mhs_bukuisbn.xlsx');

$spreadsheet_luaran_mhs_bukuisbn1->disconnectWorksheets();
unset($spreadsheet_luaran_mhs_bukuisbn1);

// Load Format Baru
$spreadsheet_luaran_mhs_bukuisbn2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_luaran_penelitian_mhs_bukuisbn.xlsx');
$worksheet_luaran_mhs_bukuisbn2 = $spreadsheet_luaran_mhs_bukuisbn2->getActiveSheet();

// Formasi Array SAPTO
$array_luaran_mhs_bukuisbn = $worksheet_luaran_mhs_bukuisbn2->toArray();
$data_luaran_mhs_bukuisbn = [];

foreach($worksheet_luaran_mhs_bukuisbn2->getRowIterator() as $row_id => $row) {
    if($worksheet_luaran_mhs_bukuisbn2->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['luaran_penelitian'] = $array_luaran_mhs_bukuisbn[$row_id-1][1];
            $item['tahun'] = $array_luaran_mhs_bukuisbn[$row_id-1][2];
			$item['keterangan'] = $array_luaran_mhs_bukuisbn[$row_id-1][3];
            $data_luaran_mhs_bukuisbn[] = $item;
        }
    }
}

$spreadsheet_luaran_mhs_bukuisbn2->disconnectWorksheets();
unset($spreadsheet_luaran_mhs_bukuisbn2);



/**
 * Set ke Template SAPTO
 */

$spreadsheet_aps = \PhpOffice\PhpSpreadsheet\IOFactory::load('./result/sapto_aps9 (F).xlsx');


// Publikasi Ilmiah Mahasiswa
$worksheet_aps30 = $spreadsheet_aps->getSheetByName('8f1-1');

// Urutan bila semua jenis publikasi lengkap ada (kondisi ideal)
$worksheet_aps30->fromArray($publikasi_mhs_jurnal_seminar4, NULL, 'C7');
$worksheet_aps30->fromArray($publikasi_mhs_jurnal_seminar3, NULL, 'C8');
$worksheet_aps30->fromArray($publikasi_mhs_jurnal_seminar1, NULL, 'C9');
$worksheet_aps30->fromArray($publikasi_mhs_jurnal_seminar2, NULL, 'C10');
$worksheet_aps30->fromArray($publikasi_mhs_jurnal_seminar7, NULL, 'C11');
$worksheet_aps30->fromArray($publikasi_mhs_jurnal_seminar6, NULL, 'C12');
$worksheet_aps30->fromArray($publikasi_mhs_jurnal_seminar5, NULL, 'C13');
$worksheet_aps30->fromArray($publikasi_mhs_non_jurnal3, NULL, 'C14');
$worksheet_aps30->fromArray($publikasi_mhs_non_jurnal2, NULL, 'C15');
$worksheet_aps30->fromArray($publikasi_mhs_non_jurnal1, NULL, 'C16');


// Pagelaran/Pameran/Presentasi/Publikasi Ilmiah Mahasiswa
$worksheet_aps31 = $spreadsheet_aps->getSheetByName('8f1-2');

// Urutan bila semua jenis publikasi lengkap ada (kondisi ideal)
$worksheet_aps31->fromArray($publikasi_mhs_jurnal_seminar4, NULL, 'C7');
$worksheet_aps31->fromArray($publikasi_mhs_jurnal_seminar3, NULL, 'C8');
$worksheet_aps31->fromArray($publikasi_mhs_jurnal_seminar1, NULL, 'C9');
$worksheet_aps31->fromArray($publikasi_mhs_jurnal_seminar2, NULL, 'C10');
$worksheet_aps31->fromArray($publikasi_mhs_jurnal_seminar7, NULL, 'C11');
$worksheet_aps31->fromArray($publikasi_mhs_jurnal_seminar6, NULL, 'C12');
$worksheet_aps31->fromArray($publikasi_mhs_jurnal_seminar5, NULL, 'C13');
$worksheet_aps31->fromArray($publikasi_mhs_non_jurnal3, NULL, 'C14');
$worksheet_aps31->fromArray($publikasi_mhs_non_jurnal2, NULL, 'C15');
$worksheet_aps31->fromArray($publikasi_mhs_non_jurnal1, NULL, 'C16');


// Karya Ilmiah Mahasiswa yang Disitasi
$worksheet_aps32 = $spreadsheet_aps->getSheetByName('8f2');;
$worksheet_aps32->fromArray($data_karya_disitasi_mhs, NULL, 'B6');

$highestRow_aps32 = $worksheet_aps32->getHighestRow();

$worksheet_aps32->getStyle('A6:D'.$highestRow_aps32)->applyFromArray($styleBorder);
$worksheet_aps32->getStyle('B6:D'.$highestRow_aps32)->applyFromArray($styleYellow);
$worksheet_aps32->getStyle('A6:A'.$highestRow_aps32)->applyFromArray($styleCenter);
$worksheet_aps32->getStyle('D6:D'.$highestRow_aps32)->applyFromArray($styleCenter);
$worksheet_aps32->getStyle('B6:D'.$highestRow_aps32)->getAlignment()->setWrapText(true);

foreach($worksheet_aps32->getRowDimensions() as $rd32) { 
    $rd32->setRowHeight(-1); 
}

for($row = 6; $row <= $highestRow_aps32; $row++) {
	$worksheet_aps32->setCellValue('A'.$row, $row-5);
}


// Produk/Jasa Mahasiswa yang Diadopsi oleh Industri/Masyarakat
$worksheet_aps44 = $spreadsheet_aps->getSheetByName('8f3');
$worksheet_aps44->fromArray($data_produk_jasa_mhs, NULL, 'B6');

$highestRow_aps44 = $worksheet_aps44->getHighestRow();

$worksheet_aps44->getStyle('A6:E'.$highestRow_aps44)->applyFromArray($styleBorder);
$worksheet_aps44->getStyle('B6:E'.$highestRow_aps44)->applyFromArray($styleYellow);
$worksheet_aps44->getStyle('A6:A'.$highestRow_aps44)->applyFromArray($styleCenter);
$worksheet_aps44->getStyle('D6:E'.$highestRow_aps44)->applyFromArray($styleCenter);
$worksheet_aps44->getStyle('B6:E'.$highestRow_aps44)->getAlignment()->setWrapText(true);

foreach($worksheet_aps44->getRowDimensions() as $rd44) { 
    $rd44->setRowHeight(-1); 
}

for($row = 6; $row <= $highestRow_aps44; $row++) {
	$worksheet_aps44->setCellValue('A'.$row, $row-5);
}


// Luaran Penelitian/PkM Lainnya oleh Mahasiswa - HKI (Paten, Paten Sederhana)
$worksheet_aps33 = $spreadsheet_aps->getSheetByName('8f4-1');
$worksheet_aps33->fromArray($data_luaran_mhs_hki_paten, NULL, 'B8');

$highestRow_aps33 = $worksheet_aps33->getHighestRow();

$worksheet_aps33->getStyle('A8:D'.$highestRow_aps33)->applyFromArray($styleBorder);
$worksheet_aps33->getStyle('B8:D'.$highestRow_aps33)->applyFromArray($styleYellow);
$worksheet_aps33->getStyle('A8:A'.$highestRow_aps33)->applyFromArray($styleCenter);
$worksheet_aps33->getStyle('C8:C'.$highestRow_aps33)->applyFromArray($styleCenter);
$worksheet_aps33->getStyle('B7:D'.$highestRow_aps33)->getAlignment()->setWrapText(true);

foreach($worksheet_aps33->getRowDimensions() as $rd33) { 
    $rd33->setRowHeight(-1); 
}

for($row = 8; $row <= $highestRow_aps33; $row++) {
	$worksheet_aps33->setCellValue('A'.$row, $row-7);
}


// Luaran Penelitian/PkM Lainnya oleh Mahasiswa - HKI (Hak Cipta, Desain Produk Industri, dll.)
$worksheet_aps34 = $spreadsheet_aps->getSheetByName('8f4-2');
$worksheet_aps34->fromArray($data_luaran_mhs_hki_cipta, NULL, 'B8');

$highestRow_aps34 = $worksheet_aps34->getHighestRow();

$worksheet_aps34->getStyle('A8:D'.$highestRow_aps34)->applyFromArray($styleBorder);
$worksheet_aps34->getStyle('B8:D'.$highestRow_aps34)->applyFromArray($styleYellow);
$worksheet_aps34->getStyle('A8:A'.$highestRow_aps34)->applyFromArray($styleCenter);
$worksheet_aps34->getStyle('C8:C'.$highestRow_aps34)->applyFromArray($styleCenter);
$worksheet_aps34->getStyle('B7:D'.$highestRow_aps34)->getAlignment()->setWrapText(true);

foreach($worksheet_aps34->getRowDimensions() as $rd34) { 
    $rd34->setRowHeight(-1); 
}

for($row = 8; $row <= $highestRow_aps34; $row++) {
	$worksheet_aps34->setCellValue('A'.$row, $row-7);
}


// Luaran Penelitian/PkM Lainnya oleh Mahasiswa - Teknologi Tepat Guna, Produk, Karya Seni, Rekayasa Sosial
$worksheet_aps35 = $spreadsheet_aps->getSheetByName('8f4-3');
$worksheet_aps35->fromArray($data_luaran_mhs_produk, NULL, 'B8');

$highestRow_aps35 = $worksheet_aps35->getHighestRow();

$worksheet_aps35->getStyle('A8:D'.$highestRow_aps35)->applyFromArray($styleBorder);
$worksheet_aps35->getStyle('B8:D'.$highestRow_aps35)->applyFromArray($styleYellow);
$worksheet_aps35->getStyle('A8:A'.$highestRow_aps35)->applyFromArray($styleCenter);
$worksheet_aps35->getStyle('C8:C'.$highestRow_aps35)->applyFromArray($styleCenter);
$worksheet_aps35->getStyle('B7:D'.$highestRow_aps35)->getAlignment()->setWrapText(true);

foreach($worksheet_aps35->getRowDimensions() as $rd35) { 
    $rd35->setRowHeight(-1); 
}

for($row = 8; $row <= $highestRow_aps35; $row++) {
	$worksheet_aps35->setCellValue('A'.$row, $row-7);
}


// Luaran Penelitian/PkM Lainnya oleh Mahasiswa - Buku ber-ISBN, Book Chapter
$worksheet_aps36 = $spreadsheet_aps->getSheetByName('8f4-4');
$worksheet_aps36->fromArray($data_luaran_mhs_bukuisbn, NULL, 'B8');

$highestRow_aps36 = $worksheet_aps36->getHighestRow();

$worksheet_aps36->getStyle('A8:D'.$highestRow_aps36)->applyFromArray($styleBorder);
$worksheet_aps36->getStyle('B8:D'.$highestRow_aps36)->applyFromArray($styleYellow);
$worksheet_aps36->getStyle('A8:A'.$highestRow_aps36)->applyFromArray($styleCenter);
$worksheet_aps36->getStyle('C8:C'.$highestRow_aps36)->applyFromArray($styleCenter);
$worksheet_aps36->getStyle('B7:D'.$highestRow_aps36)->getAlignment()->setWrapText(true);

foreach($worksheet_aps36->getRowDimensions() as $rd36) { 
    $rd36->setRowHeight(-1); 
}

for($row = 8; $row <= $highestRow_aps36; $row++) {
	$worksheet_aps36->setCellValue('A'.$row, $row-7);
}



$writer_aps = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_aps, 'Xlsx');
$writer_aps->save('./result/sapto_aps9 (F).xlsx');

$spreadsheet_aps->disconnectWorksheets();
unset($spreadsheet_aps);

?>