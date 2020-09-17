<?php

/*

ASPEK 1: KERJASAMA

*/


require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

/* if ($argc < 2 )
{
    exit( "Usage: sapto_aspek1.php <id prodi sesuai di database>\n" );
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
 * Kerjasama Tridharma
 */

$sql_kerjasama_tridharma = "SELECT id, lembaga_mitra, tingkat, judul_kerjasama, manfaat, durasi, tahun_berakhir, jenis_tridharma, prodi_id FROM akreditasi.sapto_kerjasama_tridharma WHERE prodi_id = '".$nama_prodi."'";
$stmt = sqlsrv_query( $conn, $sql_kerjasama_tridharma );
if( $stmt === false) {
    die( print_r( sqlsrv_errors(), true) );
}

$data_kerjasama_tridharma1 = [];

while( $row = sqlsrv_fetch_array( $stmt, SQLSRV_FETCH_NUMERIC) ) {
	 $data_array_kerjasama_tridharma[] = array(
		$row[0], $row[1], $row[2], $row[3], $row[4], $row[5], $row[6], $row[7], $row[8]
	);
}

$data_kerjasama_tridharma1 = array_merge($data_kerjasama_tridharma1, $data_array_kerjasama_tridharma); 

sqlsrv_free_stmt( $stmt);

$spreadsheet_kerjasama_tridharma1 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./blank/sapto_kerjasama_tridharma.xls');
$worksheet_kerjasama_tridharma1 = $spreadsheet_kerjasama_tridharma1->getActiveSheet();

$worksheet_kerjasama_tridharma1->fromArray($data_kerjasama_tridharma1, NULL, 'A2');

$worksheet_kerjasama_tridharma1->insertNewColumnBefore('D', 3);

$highestRow_kerjasama_tridharma1 = $worksheet_kerjasama_tridharma1->getHighestRow();

for($row = 2;$row <= $highestRow_kerjasama_tridharma1; $row++) {
	$worksheet_kerjasama_tridharma1->setCellValue('C'.$row, '=IF(C'.$row.'="Internasional";"V";"")');
	$worksheet_kerjasama_tridharma1->getCell('C'.$row)->getStyle()->setQuotePrefix(true);
	
	$worksheet_kerjasama_tridharma1->setCellValue('D'.$row, '=IF(C'.$row.'="Nasional";"V";"")');
	$worksheet_kerjasama_tridharma1->getCell('D'.$row)->getStyle()->setQuotePrefix(true);
	
	$worksheet_kerjasama_tridharma1->setCellValue('E'.$row, '=IF(C'.$row.'="Lokal";"V";"")');
	$worksheet_kerjasama_tridharma1->getCell('E'.$row)->getStyle()->setQuotePrefix(true);
}

$data_kerjasama = $worksheet_kerjasama_tridharma1->rangeToArray('A1:M'.$highestRow_kerjasama_tridharma1, NULL, TRUE, TRUE, TRUE);

$writer_kerjasama_tridharma1 = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_kerjasama_tridharma1, 'Xls');
$writer_kerjasama_tridharma1->save('./raw/sapto_kerjasama_tridharma.xls');

$spreadsheet_kerjasama_tridharma1->disconnectWorksheets();
unset($spreadsheet_kerjasama_tridharma1);

$spreadsheet_kerjasama_tridharma = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_kerjasama_tridharma.xls');
$worksheet_kerjasama_tridharma = $spreadsheet_kerjasama_tridharma->getActiveSheet();


/**
 * Tabel 1.1 Kerjasama Tridharma Pendidikan
 */

$worksheet_kerjasama_pendidikan = new \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet($spreadsheet_kerjasama_tridharma, 'Pendidikan');
$spreadsheet_kerjasama_tridharma->addSheet($worksheet_kerjasama_pendidikan);

$worksheet_kerjasama_pendidikan = $spreadsheet_kerjasama_tridharma->getSheetByName('Pendidikan');
$worksheet_kerjasama_pendidikan->fromArray($data_kerjasama, NULL, 'A1');

$highestRow_kerjasama_pendidikan = $worksheet_kerjasama_pendidikan->getHighestRow();

$worksheet_kerjasama_pendidikan->setAutoFilter('B1:L'.$highestRow_kerjasama_pendidikan);
$autoFilter_kerjasama_pendidikan = $worksheet_kerjasama_pendidikan->getAutoFilter();
$columnFilter_kerjasama_pendidikan = $autoFilter_kerjasama_pendidikan->getColumn('K');
$columnFilter_kerjasama_pendidikan->setFilterType(
    \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_FILTER
);
$columnFilter_kerjasama_pendidikan->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        'Pendidikan'
    );

$autoFilter_kerjasama_pendidikan->showHideRows();


/**
 * Tabel 1.2 Kerjasama Tridharma Penelitian
 */

$worksheet_kerjasama_penelitian = new \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet($spreadsheet_kerjasama_tridharma, 'Penelitian');
$spreadsheet_kerjasama_tridharma->addSheet($worksheet_kerjasama_penelitian);

$worksheet_kerjasama_penelitian = $spreadsheet_kerjasama_tridharma->getSheetByName('Penelitian');
$worksheet_kerjasama_penelitian->fromArray($data_kerjasama, NULL, 'A1');

$highestRow_kerjasama_penelitian = $worksheet_kerjasama_penelitian->getHighestRow();

$worksheet_kerjasama_penelitian->setAutoFilter('B1:L'.$highestRow_kerjasama_penelitian);
$autoFilter_kerjasama_penelitian = $worksheet_kerjasama_penelitian->getAutoFilter();
$columnFilter_kerjasama_penelitian = $autoFilter_kerjasama_penelitian->getColumn('K');
$columnFilter_kerjasama_penelitian->setFilterType(
    \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_FILTER
);
$columnFilter_kerjasama_penelitian->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        'Penelitian'
    );

$autoFilter_kerjasama_penelitian->showHideRows();


/**
 * Tabel 1.3 Kerjasama Tridharma PkM
 */

$worksheet_kerjasama_pkm = new \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet($spreadsheet_kerjasama_tridharma, 'PkM');
$spreadsheet_kerjasama_tridharma->addSheet($worksheet_kerjasama_pkm);

$worksheet_kerjasama_pkm = $spreadsheet_kerjasama_tridharma->getSheetByName('PkM');
$worksheet_kerjasama_pkm->fromArray($data_kerjasama, NULL, 'A1');

$highestRow_kerjasama_pkm = $worksheet_kerjasama_pkm->getHighestRow();

$worksheet_kerjasama_pkm->setAutoFilter('B1:L'.$highestRow_kerjasama_pkm);
$autoFilter_kerjasama_pkm = $worksheet_kerjasama_pkm->getAutoFilter();
$columnFilter_kerjasama_pkm = $autoFilter_kerjasama_pkm->getColumn('K');
$columnFilter_kerjasama_pkm->setFilterType(
    \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_FILTER
);
$columnFilter_kerjasama_pkm->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        'Pengmas'
    );

$autoFilter_kerjasama_pkm->showHideRows();


$writer_kerjasama_tridharma = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_kerjasama_tridharma, 'Xls');
$writer_kerjasama_tridharma->save('./formatted/sapto_kerjasama_tridharma (F).xls');

$spreadsheet_kerjasama_tridharma->disconnectWorksheets();
unset($spreadsheet_kerjasama_tridharma);


// Load Format Baru
$spreadsheet_kerjasama_tridharma2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_kerjasama_tridharma(F).xls');
$worksheet_kerjasama_pendidikan2 = $spreadsheet_kerjasama_tridharma->getSheetByName('Pendidikan');
$worksheet_kerjasama_penelitian2 = $spreadsheet_kerjasama_tridharma->getSheetByName('Penelitian');
$worksheet_kerjasama_pkm2 = $spreadsheet_kerjasama_tridharma->getSheetByName('PkM');

// Formasi Array SAPTO
$array_kerjasama_pendidikan = $worksheet_kerjasama_pendidikan2->toArray();
$data_kerjasama_pendidikan = [];

foreach($worksheet_kerjasama_pendidikan2->getRowIterator() as $row_id => $row) {
    if($worksheet_kerjasama_pendidikan2->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['lembaga_mitra'] = $array_kerjasama_pendidikan[$row_id-1][1];
            $item['lokal'] = $array_kerjasama_pendidikan[$row_id-1][3];
			$item['nasional'] = $array_kerjasama_pendidikan[$row_id-1][4];
			$item['internasional'] = $array_kerjasama_pendidikan[$row_id-1][5];
			$item['judul_kegiatan'] = $array_kerjasama_pendidikan[$row_id-1][6];
			$item['manfaat'] = $array_kerjasama_pendidikan[$row_id-1][7];
			$item['durasi'] = $array_kerjasama_pendidikan[$row_id-1][8];
			$item['bukti'] = $array_kerjasama_pendidikan[$row_id-1][12];
			$item['tahun_berakhir'] = $array_kerjasama_pendidikan[$row_id-1][9];
            $data_kerjasama_pendidikan[] = $item;
        }
    }
}

$spreadsheet_kerjasama_pendidikan2->disconnectWorksheets();
unset($spreadsheet_kerjasama_pendidikan2);


$array_kerjasama_penelitian = $worksheet_kerjasama_penelitian2->toArray();
$data_kerjasama_penelitian = [];

foreach($worksheet_kerjasama_penelitian2->getRowIterator() as $row_id => $row) {
    if($worksheet_kerjasama_penelitian2->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['lembaga_mitra'] = $array_kerjasama_penelitian[$row_id-1][1];
            $item['lokal'] = $array_kerjasama_penelitian[$row_id-1][3];
			$item['nasional'] = $array_kerjasama_penelitian[$row_id-1][4];
			$item['internasional'] = $array_kerjasama_penelitian[$row_id-1][5];
			$item['judul_kegiatan'] = $array_kerjasama_penelitian[$row_id-1][6];
			$item['manfaat'] = $array_kerjasama_penelitian[$row_id-1][7];
			$item['durasi'] = $array_kerjasama_penelitian[$row_id-1][8];
			$item['bukti'] = $array_kerjasama_penelitian[$row_id-1][12];
			$item['tahun_berakhir'] = $array_kerjasama_penelitian[$row_id-1][9];
            $data_kerjasama_penelitian[] = $item;
        }
    }
}

$spreadsheet_kerjasama_penelitian2->disconnectWorksheets();
unset($spreadsheet_kerjasama_penelitian2);


$array_kerjasama_pkm = $worksheet_kerjasama_pkm2->toArray();
$data_kerjasama_pkm = [];

foreach($worksheet_kerjasama_pkm2->getRowIterator() as $row_id => $row) {
    if($worksheet_kerjasama_pkm2->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['lembaga_mitra'] = $array_kerjasama_pkm[$row_id-1][1];
            $item['lokal'] = $array_kerjasama_pkm[$row_id-1][3];
			$item['nasional'] = $array_kerjasama_pkm[$row_id-1][4];
			$item['internasional'] = $array_kerjasama_pkm[$row_id-1][5];
			$item['judul_kegiatan'] = $array_kerjasama_pkm[$row_id-1][6];
			$item['manfaat'] = $array_kerjasama_pkm[$row_id-1][7];
			$item['durasi'] = $array_kerjasama_pkm[$row_id-1][8];
			$item['bukti'] = $array_kerjasama_pkm[$row_id-1][12];
			$item['tahun_berakhir'] = $array_kerjasama_pkm[$row_id-1][9];
            $data_kerjasama_pkm[] = $item;
        }
    }
}

$spreadsheet_kerjasama_pkm2->disconnectWorksheets();
unset($spreadsheet_kerjasama_pkm2);


/**
 * Set ke Template SAPTO
 */

$spreadsheet_aps = \PhpOffice\PhpSpreadsheet\IOFactory::load('./blank/sapto_aps9.xlsx');


// Kerjasama Tridharma Pendidikan
$worksheet_aps = $spreadsheet_aps->getSheetByName('1-1');
$worksheet_aps->fromArray($data_kerjasama_pendidikan, NULL, 'B12');

$highestRow_aps = $worksheet_aps->getHighestRow();
$worksheet_aps->getStyle('A12:J'.$highestRow_aps)->applyFromArray($styleBorder);
$worksheet_aps->getStyle('B12:I'.$highestRow_aps)->applyFromArray($styleYellow);
$worksheet_aps->getStyle('J12:J'.$highestRow_aps)->applyFromArray($styleGreen);
$worksheet_aps->getStyle('A12:A'.$highestRow_aps)->applyFromArray($styleCenter);
$worksheet_aps->getStyle('C12:E'.$highestRow_aps)->applyFromArray($styleCenter);
$worksheet_aps->getStyle('B12:J'.$highestRow_aps)->getAlignment()->setWrapText(true);

foreach($worksheet_aps->getRowDimensions() as $rd) { 
    $rd->setRowHeight(-1); 
}

for($row = 12; $row <= $highestRow_aps; $row++) {
	$worksheet_aps->setCellValue('A'.$row, $row-11);
}
 
// Kerjasama Tridharma Penelitian
$worksheet_aps2 = $spreadsheet_aps->getSheetByName('1-2');
$worksheet_aps2->fromArray($data_kerjasama_penelitian, NULL, 'B12');

$highestRow_aps2 = $worksheet_aps2->getHighestRow();
$worksheet_aps2->getStyle('A12:J'.$highestRow_aps2)->applyFromArray($styleBorder);
$worksheet_aps2->getStyle('B12:I'.$highestRow_aps2)->applyFromArray($styleYellow);
$worksheet_aps2->getStyle('J12:J'.$highestRow_aps2)->applyFromArray($styleGreen);
$worksheet_aps2->getStyle('A12:A'.$highestRow_aps2)->applyFromArray($styleCenter);
$worksheet_aps2->getStyle('C12:E'.$highestRow_aps2)->applyFromArray($styleCenter);
$worksheet_aps2->getStyle('B12:J'.$highestRow_aps2)->getAlignment()->setWrapText(true);

foreach($worksheet_aps2->getRowDimensions() as $rd2) { 
    $rd2->setRowHeight(-1); 
}

for($row = 12; $row <= $highestRow_aps2; $row++) {
	$worksheet_aps2->setCellValue('A'.$row, $row-11);
}

// Kerjasama Tridharma PkM
$worksheet_aps3 = $spreadsheet_aps->getSheetByName('1-3');
$worksheet_aps3->fromArray($data_kerjasama_pkm, NULL, 'B12');

$highestRow_aps3 = $worksheet_aps3->getHighestRow();
$worksheet_aps3->getStyle('A12:J'.$highestRow_aps3)->applyFromArray($styleBorder);
$worksheet_aps3->getStyle('B12:I'.$highestRow_aps3)->applyFromArray($styleYellow);
$worksheet_aps3->getStyle('J12:J'.$highestRow_aps3)->applyFromArray($styleGreen);
$worksheet_aps3->getStyle('A12:A'.$highestRow_aps3)->applyFromArray($styleCenter);
$worksheet_aps3->getStyle('C12:E'.$highestRow_aps3)->applyFromArray($styleCenter);
$worksheet_aps3->getStyle('B12:J'.$highestRow_aps3)->getAlignment()->setWrapText(true);

foreach($worksheet_aps3->getRowDimensions() as $rd3) { 
    $rd3->setRowHeight(-1); 
}

for($row = 12; $row <= $highestRow_aps3; $row++) {
	$worksheet_aps3->setCellValue('A'.$row, $row-11);
}


$writer_aps = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_aps, 'Xlsx');
$writer_aps->save('./result/sapto_aps9 (F).xlsx');

$spreadsheet_aps->disconnectWorksheets();
unset($spreadsheet_aps);

?>