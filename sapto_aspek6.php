<?php

/*

ASPEK 6: PENDIDIKAN

*/

require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

/* if ($argc < 2 )
{
    exit( "Usage: sapto_aspek6.php <id prodi sesuai di database>\n" );
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
 * Tabel 5.a Kurikulum, Capaian Pembelajaran, dan Rencana Pembelajaran
 */

$sql_kurikulum = "SELECT id, semester, kode_mk, nama_mk, mk_kompetensi, sks_kuliah, sks_seminar, sks_praktikum, konversi_jam, sikap, pengetahuan, keterampilan_umum, keterampilan_khusus, dok_rp, penyelenggara, prodi_id, tahun FROM akreditasi.sapto_kurikulum_capaian_rencana WHERE prodi_id = '".$nama_prodi."'";
$stmt = sqlsrv_query( $conn, $sql_kurikulum );
if( $stmt === false) {
    die( print_r( sqlsrv_errors(), true) );
}

$data_kurikulum1 = [];

while( $row = sqlsrv_fetch_array( $stmt, SQLSRV_FETCH_NUMERIC) ) {
	 $data_array_kurikulum[] = array(
		$row[0], $row[1], $row[2], $row[3], $row[4], $row[5], $row[6], $row[7], $row[8], $row[9], $row[10], $row[11], $row[12], $row[13], $row[14], $row[15]
	);
}

$data_kurikulum1 = array_merge($data_kurikulum1, $data_array_kurikulum); 

sqlsrv_free_stmt( $stmt);

$spreadsheet_kurikulum1 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./blank/sapto_kurikulum_capaian_rencana.xlsx');
$worksheet_kurikulum1 = $spreadsheet_kurikulum1->getActiveSheet();

$worksheet_kurikulum1->fromArray($data_kurikulum1, NULL, 'A2');

$writer_kurikulum1 = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_kurikulum1, 'Xlsx');
$writer_kurikulum1->save('./raw/sapto_kurikulum_capaian_rencana.xlsx');

$spreadsheet_kurikulum1->disconnectWorksheets();
unset($spreadsheet_kurikulum1);

$spreadsheet_kurikulum = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_kurikulum_capaian_rencana.xlsx');

$worksheet_kurikulum = $spreadsheet_kurikulum->getActiveSheet();

$worksheet_kurikulum->insertNewColumnBefore('F', 1);
$worksheet_kurikulum->insertNewColumnBefore('O', 4);

$highestRow_kurikulum = $worksheet_kurikulum->getHighestRow();

for($row = 2;$row <= $highestRow_kurikulum; $row++) {
	$worksheet_kurikulum->setCellValue('F'.$row, '=IF(E'.$row.'=1;"V";"")');
	$worksheet_kurikulum->getCell('F'.$row)->getStyle()->setQuotePrefix(true);
	
	$worksheet_kurikulum->setCellValue('O'.$row, '=IF(K'.$row.'=1;"V";"")');
	$worksheet_kurikulum->getCell('O'.$row)->getStyle()->setQuotePrefix(true);
	
	$worksheet_kurikulum->setCellValue('P'.$row, '=IF(L'.$row.'=1;"V";"")');
	$worksheet_kurikulum->getCell('P'.$row)->getStyle()->setQuotePrefix(true);
	
	$worksheet_kurikulum->setCellValue('Q'.$row, '=IF(M'.$row.'=1;"V";"")');
	$worksheet_kurikulum->getCell('Q'.$row)->getStyle()->setQuotePrefix(true);
	
	$worksheet_kurikulum->setCellValue('R'.$row, '=IF(N'.$row.'=1;"V";"")');
	$worksheet_kurikulum->getCell('R'.$row)->getStyle()->setQuotePrefix(true);
}

$writer_kurikulum = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_kurikulum, 'Xls');
$writer_kurikulum->save('./formatted/sapto_kurikulum_capaian_rencana (F).xls');

$spreadsheet_kurikulum->disconnectWorksheets();
unset($spreadsheet_kurikulum);

// Load Format Baru
$spreadsheet_kurikulum2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_kurikulum_capaian_rencana (F).xls');
$worksheet_kurikulum2 = $spreadsheet_kurikulum2->getActiveSheet();

// Formasi Array SAPTO
$array_kurikulum = $worksheet_kurikulum2->toArray();
$data_kurikulum = [];

foreach($worksheet_kurikulum2->getRowIterator() as $row_id => $row) {
    if($worksheet_kurikulum2->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['semester'] = $array_kurikulum[$row_id-1][1];
            $item['kode_mk'] = $array_kurikulum[$row_id-1][2];
			$item['matkul'] = $array_kurikulum[$row_id-1][3];
			$item['mk_kompetensi'] = $array_kurikulum[$row_id-1][5];
			$item['sks_kuliah'] = $array_kurikulum[$row_id-1][6];
			$item['sks_seminar'] = $array_kurikulum[$row_id-1][7];
			$item['sks_praktikum'] = $array_kurikulum[$row_id-1][8];
			$item['konversi_jam'] = $array_kurikulum[$row_id-1][9];
			$item['capaian_sikap'] = $array_kurikulum[$row_id-1][14];
			$item['capaian_pengetahuan'] = $array_kurikulum[$row_id-1][15];
			$item['capaian_k_umum'] = $array_kurikulum[$row_id-1][16];
			$item['capaian_k_khusus'] = $array_kurikulum[$row_id-1][17];
			$item['dok_rps'] = $array_kurikulum[$row_id-1][18];
			$item['unit_penyelenggara'] = $array_kurikulum[$row_id-1][19];
            $data_kurikulum[] = $item;
        }
    }
}

$spreadsheet_kurikulum2->disconnectWorksheets();
unset($spreadsheet_kurikulum2);


/**
 * Tabel 5.b Integrasi Kegiatan Penelitan/PkM dalam Pembelajaran
 */

$sql_integrasi_penelitian = "SELECT id, judul_penelitian, nama_dosen, mata_kuliah, bentuk_integrasi, tahun_penelitian, prodi_id FROM akreditasi.sapto_integrasi_kegiatan_penelitian WHERE prodi_id = '".$nama_prodi."'";
$stmt = sqlsrv_query( $conn, $sql_integrasi_penelitian );
if( $stmt === false) {
    die( print_r( sqlsrv_errors(), true) );
}

$data_integrasi_penelitian1 = [];

while( $row = sqlsrv_fetch_array( $stmt, SQLSRV_FETCH_NUMERIC) ) {
	 $data_array_integrasi_penelitian[] = array(
		$row[0], $row[1], $row[2], $row[3], $row[4], $row[5], $row[6]
	);
}

$data_integrasi_penelitian1 = array_merge($data_integrasi_penelitian1, $data_array_integrasi_penelitian); 

sqlsrv_free_stmt( $stmt);

$spreadsheet_integrasi_penelitian1 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./blank/sapto_integrasi_kegiatan_penelitian.xlsx');
$worksheet_integrasi_penelitian1 = $spreadsheet_integrasi_penelitian1->getActiveSheet();

$worksheet_integrasi_penelitian1->fromArray($data_integrasi_penelitian1, NULL, 'A2');

$writer_integrasi_penelitian1 = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_integrasi_penelitian1, 'Xlsx');
$writer_integrasi_penelitian1->save('./raw/sapto_integrasi_kegiatan_penelitian.xlsx');

$spreadsheet_integrasi_penelitian1->disconnectWorksheets();
unset($spreadsheet_integrasi_penelitian1);

// Load Format Baru
$spreadsheet_integrasi_penelitian2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_integrasi_kegiatan_penelitian.xlsx');
$worksheet_integrasi_penelitian2 = $spreadsheet_integrasi_penelitian2->getActiveSheet();

// Formasi Array SAPTO
$array_integrasi_penelitian = $worksheet_integrasi_penelitian2->toArray();
$data_integrasi_penelitian = [];

foreach($worksheet_integrasi_penelitian2->getRowIterator() as $row_id => $row) {
    if($worksheet_integrasi_penelitian2->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['judul_penelitian'] = $array_integrasi_penelitian[$row_id-1][1];
            $item['nama_dosen'] = $array_integrasi_penelitian[$row_id-1][2];
			$item['matkul'] = $array_integrasi_penelitian[$row_id-1][3];
			$item['bentuk_integrasi'] = $array_integrasi_penelitian[$row_id-1][4];
			$item['tahun'] = $array_integrasi_penelitian[$row_id-1][5];
            $data_integrasi_penelitian[] = $item;
        }
    }
}

$spreadsheet_integrasi_penelitian2->disconnectWorksheets();
unset($spreadsheet_integrasi_penelitian2);


/**
 * Tabel 5.c Kepuasan Mahasiswa
 */

$sql_kepuasan_mahasiswa = "SELECT id, aspek, sangat_baik, baik, cukup, kurang, rencana_tindak_lanjut, prodi_id, tahun FROM akreditasi.sapto_integrasi_kegiatan_penelitian WHERE prodi_id = '".$nama_prodi."'";
$stmt = sqlsrv_query( $conn, $sql_kepuasan_mahasiswa );
if( $stmt === false) {
    die( print_r( sqlsrv_errors(), true) );
}

$data_kepuasan_mahasiswa1 = [];

while( $row = sqlsrv_fetch_array( $stmt, SQLSRV_FETCH_NUMERIC) ) {
	 $data_array_kepuasan_mahasiswa[] = array(
		$row[0], $row[1], $row[2], $row[3], $row[4], $row[5], $row[6], $row[7], $row[8]
	);
}

$data_kepuasan_mahasiswa1 = array_merge($data_kepuasan_mahasiswa1, $data_array_kepuasan_mahasiswa); 

sqlsrv_free_stmt( $stmt);

$spreadsheet_kepuasan_mahasiswa1 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./blank/sapto_kepuasan_mahasiswa.xlsx');
$worksheet_kepuasan_mahasiswa1 = $spreadsheet_kepuasan_mahasiswa1->getActiveSheet();

$worksheet_kepuasan_mahasiswa1->fromArray($data_kepuasan_mahasiswa1, NULL, 'A2');

$writer_kepuasan_mahasiswa1 = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_kepuasan_mahasiswa1, 'Xlsx');
$writer_kepuasan_mahasiswa1->save('./raw/sapto_kepuasan_mahasiswa.xlsx');

$spreadsheet_kepuasan_mahasiswa1->disconnectWorksheets();
unset($spreadsheet_kepuasan_mahasiswa1);

// Load Format Baru
$spreadsheet_kepuasan_mahasiswa2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_kepuasan_mahasiswa.xlsx');
$worksheet_kepuasan_mahasiswa2 = $spreadsheet_kepuasan_mahasiswa2->getActiveSheet();

// Formasi Array SAPTO
$array_kepuasan_mahasiswa = $worksheet_kepuasan_mahasiswa2->toArray();
$data_kepuasan_mahasiswa = [];

foreach($worksheet_kepuasan_mahasiswa2->getRowIterator() as $row_id => $row) {
    if($worksheet_kepuasan_mahasiswa2->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['aspek'] = $array_kepuasan_mahasiswa[$row_id-1][1];
            $item['puas_sangat_baik'] = $array_kepuasan_mahasiswa[$row_id-1][2];
			$item['puas_baik'] = $array_kepuasan_mahasiswa[$row_id-1][3];
			$item['puas_cukup'] = $array_kepuasan_mahasiswa[$row_id-1][4];
			$item['puas_kurang'] = $array_kepuasan_mahasiswa[$row_id-1][5];
			$item['rencana_lanjut'] = $array_kepuasan_mahasiswa[$row_id-1][6];
            $data_kepuasan_mahasiswa[] = $item;
        }
    }
}

$spreadsheet_kepuasan_mahasiswa2->disconnectWorksheets();
unset($spreadsheet_kepuasan_mahasiswa2);


/**
 * Set ke Template SAPTO
 */

$spreadsheet_aps = \PhpOffice\PhpSpreadsheet\IOFactory::load('./result/sapto_aps9 (F).xlsx');


// Kurikulum, Capaian Pembelajaran, dan Rencana Pembelajaran
$worksheet_aps17 = $spreadsheet_aps->getSheetByName('5a');
$worksheet_aps17->fromArray($data_kurikulum, NULL, 'B10');

$highestRow_aps17 = $worksheet_aps17->getHighestRow();

$worksheet_aps17->getStyle('A10:O'.$highestRow_aps17)->applyFromArray($styleBorder);
$worksheet_aps17->getStyle('B10:O'.$highestRow_aps17)->applyFromArray($styleYellow);
$worksheet_aps17->getStyle('A10:C'.$highestRow_aps17)->applyFromArray($styleCenter);
$worksheet_aps17->getStyle('E10:N'.$highestRow_aps17)->applyFromArray($styleCenter);
$worksheet_aps17->getStyle('B10:O'.$highestRow_aps17)->getAlignment()->setWrapText(true);

foreach($worksheet_aps17->getRowDimensions() as $rd17) { 
    $rd17->setRowHeight(-1); 
}

for($row = 10; $row <= $highestRow_aps17; $row++) {
	$worksheet_aps17->setCellValue('A'.$row, $row-9);
}


// Integrasi Kegiatan Penelitan/PkM dalam Pembelajaran
$worksheet_aps18 = $spreadsheet_aps->getSheetByName('5b');
$worksheet_aps18->fromArray($data_integrasi_penelitian, NULL, 'B5');

$highestRow_aps18 = $worksheet_aps18->getHighestRow();

$worksheet_aps18->getStyle('A5:F'.$highestRow_aps18)->applyFromArray($styleBorder);
$worksheet_aps18->getStyle('B5:F'.$highestRow_aps18)->applyFromArray($styleYellow);
$worksheet_aps18->getStyle('A5:A'.$highestRow_aps18)->applyFromArray($styleCenter);
$worksheet_aps18->getStyle('D5:F'.$highestRow_aps18)->applyFromArray($styleCenter);
$worksheet_aps18->getStyle('B5:F'.$highestRow_aps18)->getAlignment()->setWrapText(true);

foreach($worksheet_aps18->getRowDimensions() as $rd18) { 
    $rd18->setRowHeight(-1); 
}

for($row = 5; $row <= $highestRow_aps18; $row++) {
	$worksheet_aps18->setCellValue('A'.$row, $row-4);
}


// Kepuasan Mahasiswa
$worksheet_aps19 = $spreadsheet_aps->getSheetByName('5c');
$worksheet_aps19->fromArray($data_kepuasan_mahasiswa, NULL, 'B6');

foreach($worksheet_aps19->getRowDimensions() as $rd19) { 
    $rd19->setRowHeight(-1); 
}


$writer_aps = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_aps, 'Xlsx');
$writer_aps->save('./result/sapto_aps9 (F).xlsx');

$spreadsheet_aps->disconnectWorksheets();
unset($spreadsheet_aps);

?>