<?php

/*

ASPEK 2: MAHASISWA

*/


require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

/* if ($argc < 2 )
{
    exit( "Usage: sapto_aspek2.php <id prodi sesuai di database>\n" );
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
 * Tabel 2.a Seleksi Mahasiswa
 */

$sql_seleksi_mhs = "SELECT id, prodi_id, tahun_akademik, daya_tampung, peminat, diterima, daftar_ulang, daftar_ulang_trf, jml_mahasiswa_reguler, jml_mahasiswa_trf FROM akreditasi.sapto_seleksi_mhs_baru WHERE prodi_id = '".$nama_prodi."' ORDER BY tahun_akademik DESC";
$stmt = sqlsrv_query( $conn, $sql_seleksi_mhs );
if( $stmt === false) {
    die( print_r( sqlsrv_errors(), true) );
}

$data_seleksi_mhs1 = [];

while( $row = sqlsrv_fetch_array( $stmt, SQLSRV_FETCH_NUMERIC) ) {
	 $data_array_seleksi_mhs[] = array(
		$row[0], $row[1], $row[2], $row[3], $row[4], $row[5], $row[6], $row[7], $row[8], $row[9]
	);
}

$data_seleksi_mhs1 = array_merge($data_seleksi_mhs1, $data_array_seleksi_mhs); 

sqlsrv_free_stmt($stmt);

$spreadsheet_seleksi_mhs1 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./blank/sapto_seleksi_mhs_baru.xlsx');
$worksheet_seleksi_mhs1 = $spreadsheet_seleksi_mhs1->getActiveSheet();

$worksheet_seleksi_mhs1->fromArray($data_seleksi_mhs1, NULL, 'A2');

$writer_seleksi_mhs1 = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_seleksi_mhs1, 'Xlsx');
$writer_seleksi_mhs1->save('./raw/sapto_seleksi_mhs_baru.xlsx');

$spreadsheet_seleksi_mhs1->disconnectWorksheets();
unset($spreadsheet_seleksi_mhs1);

$spreadsheet_seleksi_mhs = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_seleksi_mhs_baru.xlsx');
$worksheet_seleksi_mhs = $spreadsheet_seleksi_mhs->getActiveSheet();

$highestRow_seleksi_mhs = $worksheet_seleksi_mhs->getHighestRow();

$worksheet_seleksi_mhs->setAutoFilter('B1:J'.$highestRow_seleksi_mhs);
$autoFilter_seleksi_mhs = $worksheet_seleksi_mhs->getAutoFilter();
$columnFilter_seleksi_mhs = $autoFilter_seleksi_mhs->getColumn('C');
$columnFilter_seleksi_mhs->setFilterType(
    \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_FILTER
);
$columnFilter_seleksi_mhs->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        'TS-4'
    );
$columnFilter_seleksi_mhs->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        'TS-3'
    );
$columnFilter_seleksi_mhs->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        'TS-2'
    );
$columnFilter_seleksi_mhs->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        'TS-1'
    );
$columnFilter_seleksi_mhs->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        'TS'
    );

$autoFilter_seleksi_mhs->showHideRows();

$writer_seleksi_mhs = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_seleksi_mhs, 'Xlsx');
$writer_seleksi_mhs->save('./formatted/sapto_seleksi_mhs_baru (F).xlsx');

$spreadsheet_seleksi_mhs->disconnectWorksheets();
unset($spreadsheet_seleksi_mhs);

// Load Format Baru
$spreadsheet_seleksi_mhs2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_seleksi_mhs_baru (F).xlsx');
$worksheet_seleksi_mhs2 = $spreadsheet_seleksi_mhs2->getActiveSheet();

$array_seleksi_mhs = $worksheet_seleksi_mhs2->toArray();
$data_seleksi_mhs = [];

foreach($worksheet_seleksi_mhs2->getRowIterator() as $row_id => $row) {
    if($worksheet_seleksi_mhs->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['tahun_akademik'] = $array_seleksi_mhs[$row_id-1][2];
            $item['daya_tampung'] = $array_seleksi_mhs[$row_id-1][3];
			$item['pendaftar'] = $array_seleksi_mhs[$row_id-1][4];
			$item['lulus_seleksi'] = $array_seleksi_mhs[$row_id-1][5];
			$item['maba_reguler'] = $array_seleksi_mhs[$row_id-1][6];
			$item['maba_transfer'] = $array_seleksi_mhs[$row_id-1][7];
			$item['mhs_reguler'] = $array_seleksi_mhs[$row_id-1][8];
			$item['mhs_transfer'] = $array_seleksi_mhs[$row_id-1][9];
            $data_seleksi_mhs[] = $item;
        }
    }
}


$spreadsheet_seleksi_mhs2->disconnectWorksheets();
unset($spreadsheet_seleksi_mhs2);


/**
 * Set ke Template SAPTO
 */

$spreadsheet_aps = \PhpOffice\PhpSpreadsheet\IOFactory::load('./result/sapto_aps9 (F).xlsx');


// Seleksi Mahasiswa Baru
$worksheet_aps4 = $spreadsheet_aps->getSheetByName('2a');
$worksheet_aps4->fromArray($data_seleksi_mhs, NULL, 'A6');


$writer_aps = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_aps, 'Xlsx');
$writer_aps->save('./result/sapto_aps9 (F).xlsx');

$spreadsheet_aps->disconnectWorksheets();
unset($spreadsheet_aps);

?>