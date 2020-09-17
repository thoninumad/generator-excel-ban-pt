<?php

/*

ASPEK 7: PENELITIAN

*/

require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

/* if ($argc < 2 )
{
    exit( "Usage: sapto_aspek7.php <id prodi sesuai di database>\n" );
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
 * Tabel 6.a Penelitian DTPS yang Melibatkan Mahasiswa
 */

$sql_penelitian_dtps_mhs = "SELECT id, nama_peneliti, bidang_penelitian, anggota_mhs_nama, judul_penelitian, tahun, id_prodi FROM akreditasi.sapto_penelitian_dtps_mhs WHERE id_prodi = '".$nama_prodi."'";
$stmt = sqlsrv_query( $conn, $sql_penelitian_dtps_mhs );
if( $stmt === false) {
    die( print_r( sqlsrv_errors(), true) );
}

$data_penelitian_dtps_mhs1 = [];

while( $row = sqlsrv_fetch_array( $stmt, SQLSRV_FETCH_NUMERIC) ) {
	 $data_array_penelitian_dtps_mhs[] = array(
		$row[0], $row[1], $row[2], $row[3], $row[4], $row[5], $row[6]
	);
}

$data_penelitian_dtps_mhs1 = array_merge($data_penelitian_dtps_mhs1, $data_array_penelitian_dtps_mhs); 

sqlsrv_free_stmt( $stmt);

$spreadsheet_penelitian_dtps_mhs1 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./blank/sapto_penelitian_dtps_mhs.xlsx');
$worksheet_penelitian_dtps_mhs1 = $spreadsheet_penelitian_dtps_mhs1->getActiveSheet();

$worksheet_penelitian_dtps_mhs1->fromArray($data_penelitian_dtps_mhs1, NULL, 'A2');

$writer_penelitian_dtps_mhs1 = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_penelitian_dtps_mhs1, 'Xlsx');
$writer_penelitian_dtps_mhs1->save('./raw/sapto_penelitian_dtps_mhs.xlsx');

$spreadsheet_penelitian_dtps_mhs1->disconnectWorksheets();
unset($spreadsheet_penelitian_dtps_mhs1);

// Load Format Baru
$spreadsheet_penelitian_dtps_mhs2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_penelitian_dtps_mhs.xlsx');
$worksheet_penelitian_dtps_mhs2 = $spreadsheet_penelitian_dtps_mhs2->getActiveSheet();

// Formasi Array SAPTO
$array_penelitian_dtps_mhs = $worksheet_penelitian_dtps_mhs2->toArray();
$data_penelitian_dtps_mhs = [];

foreach($worksheet_penelitian_dtps_mhs2->getRowIterator() as $row_id => $row) {
    if($worksheet_penelitian_dtps_mhs2->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['nama_dosen'] = $array_penelitian_dtps_mhs[$row_id-1][1];
            $item['tema_penelitian'] = $array_penelitian_dtps_mhs[$row_id-1][2];
			$item['nama_mhs'] = $array_penelitian_dtps_mhs[$row_id-1][3];
			$item['judul_kegiatan'] = $array_penelitian_dtps_mhs[$row_id-1][4];
			$item['tahun'] = $array_penelitian_dtps_mhs[$row_id-1][5];
            $data_penelitian_dtps_mhs[] = $item;
        }
    }
}

$spreadsheet_penelitian_dtps_mhs2->disconnectWorksheets();
unset($spreadsheet_penelitian_dtps_mhs2);


/**
 * Tabel 6.b Penelitian DTPS yang Menjadi Rujukan Tema Tesis/Disertasi
 */

$sql_penelitian_dtps_rujukan_tesis = "SELECT id, nama_dosen, tema_penelitian, nama_mhs, judul_tesis_disertasi, tahun_penelitian, id_prodi FROM akreditasi.sapto_penelitian_dtps_rujukan_tesis WHERE id_prodi = '".$nama_prodi."'";
$stmt = sqlsrv_query( $conn, $sql_penelitian_dtps_rujukan_tesis );
if( $stmt === false) {
    die( print_r( sqlsrv_errors(), true) );
}

$data_penelitian_dtps_rujukan_tesis1 = [];

while( $row = sqlsrv_fetch_array( $stmt, SQLSRV_FETCH_NUMERIC) ) {
	 $data_array_penelitian_dtps_rujukan_tesis[] = array(
		$row[0], $row[1], $row[2], $row[3], $row[4], $row[5], $row[6]
	);
}

$data_penelitian_dtps_rujukan_tesis1 = array_merge($data_penelitian_dtps_rujukan_tesis1, $data_array_penelitian_dtps_rujukan_tesis); 

sqlsrv_free_stmt( $stmt);

$spreadsheet_penelitian_dtps_rujukan_tesis1 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./blank/sapto_penelitian_dtps_rujukan_tesis.xlsx');
$worksheet_penelitian_dtps_rujukan_tesis1 = $spreadsheet_penelitian_dtps_rujukan_tesis1->getActiveSheet();

$worksheet_penelitian_dtps_rujukan_tesis1->fromArray($data_penelitian_dtps_rujukan_tesis1, NULL, 'A2');

$writer_penelitian_dtps_rujukan_tesis1 = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_penelitian_dtps_rujukan_tesis1, 'Xlsx');
$writer_penelitian_dtps_rujukan_tesis1->save('./raw/sapto_penelitian_dtps_rujukan_tesis.xlsx');

$spreadsheet_penelitian_dtps_rujukan_tesis1->disconnectWorksheets();
unset($spreadsheet_penelitian_dtps_rujukan_tesis1);

// Load Format Baru
$spreadsheet_penelitian_dtps_rujukan_tesis2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_penelitian_dtps_rujukan_tesis.xlsx');
$worksheet_penelitian_dtps_rujukan_tesis2 = $spreadsheet_penelitian_dtps_rujukan_tesis2->getActiveSheet();

// May be change
$array_penelitian_dtps_rujukan_tesis = $worksheet_penelitian_dtps_rujukan_tesis2->toArray();
$data_penelitian_dtps_rujukan_tesis = [];

foreach($worksheet_penelitian_dtps_rujukan_tesis2->getRowIterator() as $row_id => $row) {
    if($worksheet_penelitian_dtps_rujukan_tesis2->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['nama_dosen'] = $array_penelitian_dtps_rujukan_tesis[$row_id-1][1];
            $item['tema_penelitian'] = $array_penelitian_dtps_rujukan_tesis[$row_id-1][2];
			$item['nama_mhs'] = $array_penelitian_dtps_rujukan_tesis[$row_id-1][3];
			$item['judul_tesis'] = $array_penelitian_dtps_rujukan_tesis[$row_id-1][4];
			$item['tahun'] = $array_penelitian_dtps_rujukan_tesis[$row_id-1][5];
            $data_penelitian_dtps_rujukan_tesis[] = $item;
        }
    }
}

$spreadsheet_penelitian_dtps_rujukan_tesis2->disconnectWorksheets();
unset($spreadsheet_penelitian_dtps_rujukan_tesis2);



/**
 * Set ke Template SAPTO
 */

$spreadsheet_aps = \PhpOffice\PhpSpreadsheet\IOFactory::load('./result/sapto_aps9 (F).xlsx');


// Penelitian DTPS yang Melibatkan Mahasiswa
$worksheet_aps21 = $spreadsheet_aps->getSheetByName('6a');
$worksheet_aps21->fromArray($data_penelitian_dtps_mhs, NULL, 'B6');

$highestRow_aps21 = $worksheet_aps21->getHighestRow();

$worksheet_aps21->getStyle('A6:F'.$highestRow_aps21)->applyFromArray($styleBorder);
$worksheet_aps21->getStyle('B6:F'.$highestRow_aps21)->applyFromArray($styleYellow);
$worksheet_aps21->getStyle('A6:A'.$highestRow_aps21)->applyFromArray($styleCenter);
$worksheet_aps21->getStyle('C6:F'.$highestRow_aps21)->applyFromArray($styleCenter);
$worksheet_aps21->getStyle('B6:F'.$highestRow_aps21)->getAlignment()->setWrapText(true);

foreach($worksheet_aps21->getRowDimensions() as $rd21) { 
    $rd21->setRowHeight(-1); 
}

for($row = 6; $row <= $highestRow_aps21; $row++) {
	$worksheet_aps21->setCellValue('A'.$row, $row-5);
}


// Penelitian DTPS yang Menjadi Rujukan Tema Tesis/Disertasi
$worksheet_aps22 = $spreadsheet_aps->getSheetByName('6b');
$worksheet_aps22->fromArray($data_penelitian_dtps_rujukan_tesis, NULL, 'B6');

$highestRow_aps22 = $worksheet_aps22->getHighestRow();

$worksheet_aps22->getStyle('A6:F'.$highestRow_aps22)->applyFromArray($styleBorder);
$worksheet_aps22->getStyle('B6:F'.$highestRow_aps22)->applyFromArray($styleYellow);
$worksheet_aps22->getStyle('A6:A'.$highestRow_aps22)->applyFromArray($styleCenter);
$worksheet_aps22->getStyle('C6:F'.$highestRow_aps22)->applyFromArray($styleCenter);
$worksheet_aps22->getStyle('B6:F'.$highestRow_aps22)->getAlignment()->setWrapText(true);

foreach($worksheet_aps22->getRowDimensions() as $rd22) { 
    $rd22->setRowHeight(-1); 
}

for($row = 6; $row <= $highestRow_aps22; $row++) {
	$worksheet_aps22->setCellValue('A'.$row, $row-5);
}


$writer_aps = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_aps, 'Xlsx');
$writer_aps->save('./result/sapto_aps9 (F).xlsx');

$spreadsheet_aps->disconnectWorksheets();
unset($spreadsheet_aps);

?>