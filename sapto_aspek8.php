<?php

/*

ASPEK 8: PENGABDIAN KEPADA MASYARAKAT

*/

require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

/* if ($argc < 2 )
{
    exit( "Usage: sapto_aspek8.php <id prodi sesuai di database>\n" );
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
 * Tabel 7 PkM DTPS yang Melibatkan Mahasiswa
 */

$sql_pkm_dtps_mhs = "SELECT id, nama, bidang_pengmas, anggota_mhs_nama, judul_pengmas, tahun, id_prodi FROM akreditasi.sapto_pkm_dtps_mhs WHERE id_prodi = '".$nama_prodi."'";
$stmt = sqlsrv_query( $conn, $sql_pkm_dtps_mhs );
if( $stmt === false) {
    die( print_r( sqlsrv_errors(), true) );
}

$data_pkm_dtps_mhs1 = [];

while( $row = sqlsrv_fetch_array( $stmt, SQLSRV_FETCH_NUMERIC) ) {
	 $data_array_pkm_dtps_mhs[] = array(
		$row[0], $row[1], $row[2], $row[3], $row[4], $row[5], $row[6]
	);
}

$data_pkm_dtps_mhs1 = array_merge($data_pkm_dtps_mhs1, $data_array_pkm_dtps_mhs); 

sqlsrv_free_stmt( $stmt);

$spreadsheet_pkm_dtps_mhs1 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./blank/sapto_pkm_dtps_mhs.xlsx');
$worksheet_pkm_dtps_mhs1 = $spreadsheet_pkm_dtps_mhs1->getActiveSheet();

$worksheet_pkm_dtps_mhs1->fromArray($data_pkm_dtps_mhs1, NULL, 'A2');

$writer_pkm_dtps_mhs1 = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_pkm_dtps_mhs1, 'Xlsx');
$writer_pkm_dtps_mhs1->save('./raw/sapto_pkm_dtps_mhs.xlsx');

$spreadsheet_pkm_dtps_mhs1->disconnectWorksheets();
unset($spreadsheet_pkm_dtps_mhs1);

// Load Format Baru
$spreadsheet_pkm_dtps_mhs2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_pkm_dtps_mhs.xlsx');
$worksheet_pkm_dtps_mhs2 = $spreadsheet_pkm_dtps_mhs2->getActiveSheet();

// Formasi Array SAPTO
$array_pkm_dtps_mhs = $worksheet_pkm_dtps_mhs2->toArray();
$data_pkm_dtps_mhs = [];

foreach($worksheet_pkm_dtps_mhs2->getRowIterator() as $row_id => $row) {
    if($worksheet_pkm_dtps_mhs2->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['nama_dosen'] = $array_pkm_dtps_mhs[$row_id-1][1];
            $item['tema_pkm'] = $array_pkm_dtps_mhs[$row_id-1][2];
			$item['nama_mhs'] = $array_pkm_dtps_mhs[$row_id-1][3];
			$item['judul_kegiatan'] = $array_pkm_dtps_mhs[$row_id-1][4];
			$item['tahun'] = $array_pkm_dtps_mhs[$row_id-1][5];
            $data_pkm_dtps_mhs[] = $item;
        }
    }
}

$spreadsheet_pkm_dtps_mhs2->disconnectWorksheets();
unset($spreadsheet_pkm_dtps_mhs2);


/**
 * Set ke Template SAPTO
 */

$spreadsheet_aps = \PhpOffice\PhpSpreadsheet\IOFactory::load('./result/sapto_aps9 (F).xlsx');


// PkM DTPS yang Melibatkan Mahasiswa
$worksheet_aps23 = $spreadsheet_aps->getSheetByName('7');
$worksheet_aps23->fromArray($data_pkm_dtps_mhs, NULL, 'B6');

$highestRow_aps23 = $worksheet_aps23->getHighestRow();

$worksheet_aps23->getStyle('A6:F'.$highestRow_aps23)->applyFromArray($styleBorder);
$worksheet_aps23->getStyle('B6:F'.$highestRow_aps23)->applyFromArray($styleYellow);
$worksheet_aps23->getStyle('A6:A'.$highestRow_aps23)->applyFromArray($styleCenter);
$worksheet_aps23->getStyle('C6:F'.$highestRow_aps23)->applyFromArray($styleCenter);
$worksheet_aps23->getStyle('B6:F'.$highestRow_aps23)->getAlignment()->setWrapText(true);

foreach($worksheet_aps23->getRowDimensions() as $rd23) { 
    $rd23->setRowHeight(-1); 
}

for($row = 6; $row <= $highestRow_aps23; $row++) {
	$worksheet_aps23->setCellValue('A'.$row, $row-5);
}


$writer_aps = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_aps, 'Xlsx');
$writer_aps->save('./result/sapto_aps9 (F).xlsx');

$spreadsheet_aps->disconnectWorksheets();
unset($spreadsheet_aps);

?>