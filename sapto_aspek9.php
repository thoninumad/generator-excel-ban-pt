<?php

/*

ASPEK 9: LUARAN DHARMA PENDIDIKAN

*/


require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

/* if ($argc < 2 )
{
    exit( "Usage: sapto_aspek9.php <id prodi sesuai di database>\n" );
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
$nama_prodi = 'S-1 T.MESIN';
$nama_prodi2 = 15;

$serverName = "10.199.16.69";
$connectionInfo = array( "Database"=>"its-report", "UID"=>"sa", "PWD"=>"Akreditasi2019!");
$conn = sqlsrv_connect( $serverName, $connectionInfo );
if( $conn === false ) {
    die( print_r( sqlsrv_errors(), true));
}


/**
 * Tabel 8.a IPK Lulusan
 */

$sql_ipk_lulusan = "SELECT id, id_prodi, tahun_akademik, jumlah_lulusan, ipk_min, ipk_rata, ipk_max FROM akreditasi.sapto_ipk_lulusan WHERE id_prodi = '".$nama_prodi2."' ORDER BY tahun_akademik DESC";
$stmt = sqlsrv_query( $conn, $sql_ipk_lulusan );
if( $stmt === false) {
    die( print_r( sqlsrv_errors(), true) );
}

$data_ipk_lulusan1 = [];

while( $row = sqlsrv_fetch_array( $stmt, SQLSRV_FETCH_NUMERIC) ) {
	 $data_array_ipk_lulusan[] = array(
		$row[0], $row[1], $row[2], $row[3], $row[4], $row[5], $row[6]
	);
}

$data_ipk_lulusan1 = array_merge($data_ipk_lulusan1, $data_array_ipk_lulusan); 

sqlsrv_free_stmt( $stmt);

$spreadsheet_ipk_lulusan1 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./blank/sapto_ipk_lulusan.xlsx');
$worksheet_ipk_lulusan1 = $spreadsheet_ipk_lulusan1->getActiveSheet();

$worksheet_ipk_lulusan1->fromArray($data_ipk_lulusan1, NULL, 'A2');

$writer_ipk_lulusan1 = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_ipk_lulusan1, 'Xlsx');
$writer_ipk_lulusan1->save('./raw/sapto_ipk_lulusan.xlsx');

$spreadsheet_ipk_lulusan1->disconnectWorksheets();
unset($spreadsheet_ipk_lulusan1);

$spreadsheet_ipk_lulusan = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_ipk_lulusan.xlsx')
$worksheet_ipk_lulusan = $spreadsheet_ipk_lulusan->getActiveSheet();

$highestRow_ipk_lulusan = $worksheet_ipk_lulusan->getHighestRow();

$worksheet_ipk_lulusan->setAutoFilter('B1:G'.$highestRow_ipk_lulusan);
$autoFilter_ipk_lulusan = $worksheet_ipk_lulusan->getAutoFilter();
$columnFilter_ipk_lulusan = $autoFilter_ipk_lulusan->getColumn('C');
$columnFilter_ipk_lulusan->setFilterType(
    \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_FILTER
);
$columnFilter_ipk_lulusan->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        'TS-2'
    );
$columnFilter_ipk_lulusan->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        'TS-1'
    );
$columnFilter_ipk_lulusan->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        'TS'
    );

$autoFilter_ipk_lulusan->showHideRows();

$writer_ipk_lulusan = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_ipk_lulusan, 'Xlsx');
$writer_ipk_lulusan->save('./formatted/sapto_ipk_lulusan (F).xlsx');

$spreadsheet_ipk_lulusan->disconnectWorksheets();
unset($spreadsheet_ipk_lulusan);

// Load Format Baru
$spreadsheet_ipk_lulusan2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_ipk_lulusan (F).xlsx');
$worksheet_ipk_lulusan2 = $spreadsheet_ipk_lulusan2->getActiveSheet();

// Formasi Array SAPTO
$array_ipk_lulusan = $worksheet_ipk_lulusan2->toArray();
$data_ipk_lulusan = [];

foreach($worksheet_ipk_lulusan2->getRowIterator() as $row_id => $row) {
    if($worksheet_ipk_lulusan2->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['tahun_akademik'] = $array_ipk_lulusan[$row_id-1][2];
            $item['jml_lulusan'] = $array_ipk_lulusan[$row_id-1][3];
			$item['ipk_min'] = $array_ipk_lulusan[$row_id-1][4];
			$item['ipk_rata'] = $array_ipk_lulusan[$row_id-1][5];
			$item['ipk_maks'] = $array_ipk_lulusan[$row_id-1][6];
            $data_ipk_lulusan[] = $item;
        }
    }
}

$spreadsheet_ipk_lulusan2->disconnectWorksheets();
unset($spreadsheet_ipk_lulusan2);


/**
 * Tabel 8.b.1 Prestasi Akademik Mahasiswa
 */

$sql_prestasi_akademik = "SELECT id, id_prodi, nama_kegiatan, waktu_perolehan, tingkat, capaian, is_akademik FROM akreditasi.sapto_prestasi_mhs WHERE id_prodi = '".$nama_prodi2."' AND is_akademik = 1";
$stmt = sqlsrv_query( $conn, $sql_prestasi_akademik );
if( $stmt === false) {
    die( print_r( sqlsrv_errors(), true) );
}

$data_prestasi_akademik1 = [];

while( $row = sqlsrv_fetch_array( $stmt, SQLSRV_FETCH_NUMERIC) ) {
	 $data_array_prestasi_akademik[] = array(
		$row[0], $row[1], $row[2], $row[3], $row[4], $row[5], $row[6]
	);
}

$data_prestasi_akademik1 = array_merge($data_prestasi_akademik1, $data_array_prestasi_akademik); 

sqlsrv_free_stmt( $stmt);

$spreadsheet_prestasi_akademik1 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./blank/sapto_prestasi_mhs_akademik.xlsx');
$worksheet_prestasi_akademik1 = $spreadsheet_prestasi_akademik1->getActiveSheet();

$worksheet_prestasi_akademik1->fromArray($data_prestasi_akademik1, NULL, 'A2');

$writer_prestasi_akademik1 = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_prestasi_akademik1, 'Xlsx');
$writer_prestasi_akademik1->save('./raw/sapto_prestasi_mhs_akademik.xlsx');

$spreadsheet_prestasi_akademik1->disconnectWorksheets();
unset($spreadsheet_prestasi_akademik1);

$spreadsheet_prestasi_akademik = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_prestasi_mhs_akademik.xlsx');

$worksheet_prestasi_akademik = $spreadsheet_prestasi_akademik->getActiveSheet();

$worksheet_prestasi_akademik->insertNewColumnBefore('F', 3);

$highestRow_prestasi_akademik = $worksheet_prestasi_akademik->getHighestRow();

for($row = 2;$row <= $highestRow_prestasi_akademik; $row++) {
	$worksheet_prestasi_akademik->setCellValue('F'.$row, '=IF(E'.$row.'="Lokal";"V";"")');
	$worksheet_prestasi_akademik->getCell('F'.$row)->getStyle()->setQuotePrefix(true);
	
	$worksheet_prestasi_akademik->setCellValue('G'.$row, '=IF(F'.$row.'="Nasional";"V";"")');
	$worksheet_prestasi_akademik->getCell('G'.$row)->getStyle()->setQuotePrefix(true);
	
	$worksheet_prestasi_akademik->setCellValue('H'.$row, '=IF(G'.$row.'="Internasional";"V";"")');
	$worksheet_prestasi_akademik->getCell('H'.$row)->getStyle()->setQuotePrefix(true);
}

$writer_prestasi_akademik = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_prestasi_akademik, 'Xls');
$writer_prestasi_akademik->save('./formatted/sapto_prestasi_mhs_akademik (F).xls');

$spreadsheet_prestasi_akademik->disconnectWorksheets();
unset($spreadsheet_prestasi_akademik);

// Load Format Baru
$spreadsheet_prestasi_akademik2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_prestasi_mhs_akademik (F).xls');
$worksheet_prestasi_akademik2 = $spreadsheet_prestasi_akademik2->getActiveSheet();

// Formasi Array SAPTO
$array_prestasi_akademik = $worksheet_prestasi_akademik2->toArray();
$data_prestasi_akademik = [];

foreach($worksheet_prestasi_akademik2->getRowIterator() as $row_id => $row) {
    if($worksheet_prestasi_akademik2->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['nama_kegiatan'] = $array_prestasi_akademik[$row_id-1][2];
            $item['tahun'] = $array_prestasi_akademik[$row_id-1][3];
			$item['lokal'] = $array_prestasi_akademik[$row_id-1][5];
			$item['nasional'] = $array_prestasi_akademik[$row_id-1][6];
			$item['internasional'] = $array_prestasi_akademik[$row_id-1][7];
			$item['capaian'] = $array_prestasi_akademik[$row_id-1][8];
            $data_prestasi_akademik[] = $item;
        }
    }
}

$spreadsheet_prestasi_akademik2->disconnectWorksheets();
unset($spreadsheet_prestasi_akademik2);


/**
 * Tabel 8.b.2 Prestasi Non Akademik Mahasiswa
 */

$sql_prestasi_non_akademik = "SELECT id, id_prodi, nama_kegiatan, waktu_perolehan, tingkat, capaian, is_akademik FROM akreditasi.sapto_prestasi_mhs WHERE id_prodi = '".$nama_prodi2."' AND is_akademik = NULL";
$stmt = sqlsrv_query( $conn, $sql_prestasi_non_akademik );
if( $stmt === false) {
    die( print_r( sqlsrv_errors(), true) );
}

$data_prestasi_non_akademik1 = [];

while( $row = sqlsrv_fetch_array( $stmt, SQLSRV_FETCH_NUMERIC) ) {
	 $data_array_prestasi_non_akademik[] = array(
		$row[0], $row[1], $row[2], $row[3], $row[4], $row[5], $row[6]
	);
}

$data_prestasi_non_akademik1 = array_merge($data_prestasi_non_akademik1, $data_array_prestasi_non_akademik); 

sqlsrv_free_stmt( $stmt);

$spreadsheet_prestasi_non_akademik1 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./blank/sapto_prestasi_mhs_non_akademik.xlsx');
$worksheet_prestasi_non_akademik1 = $spreadsheet_prestasi_non_akademik1->getActiveSheet();

$worksheet_prestasi_non_akademik1->fromArray($data_prestasi_non_akademik1, NULL, 'A2');

$writer_prestasi_non_akademik1 = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_prestasi_non_akademik1, 'Xlsx');
$writer_prestasi_non_akademik1->save('./raw/sapto_prestasi_mhs_non_akademik.xlsx');

$spreadsheet_prestasi_non_akademik1->disconnectWorksheets();
unset($spreadsheet_prestasi_non_akademik1);

$spreadsheet_prestasi_non_akademik = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_prestasi_mhs_non_akademik.xlsx');

$worksheet_prestasi_non_akademik = $spreadsheet_prestasi_non_akademik->getActiveSheet();

$worksheet_prestasi_non_akademik->insertNewColumnBefore('F', 3);

$highestRow_prestasi_non_akademik = $worksheet_prestasi_non_akademik->getHighestRow();

for($row = 2;$row <= $highestRow_prestasi_non_akademik; $row++) {
	$worksheet_prestasi_non_akademik->setCellValue('F'.$row, '=IF(E'.$row.'="Lokal";"V";"")');
	$worksheet_prestasi_non_akademik->getCell('F'.$row)->getStyle()->setQuotePrefix(true);
	
	$worksheet_prestasi_non_akademik->setCellValue('G'.$row, '=IF(F'.$row.'="Nasional";"V";"")');
	$worksheet_prestasi_non_akademik->getCell('G'.$row)->getStyle()->setQuotePrefix(true);
	
	$worksheet_prestasi_non_akademik->setCellValue('H'.$row, '=IF(G'.$row.'="Internasional";"V";"")');
	$worksheet_prestasi_non_akademik->getCell('H'.$row)->getStyle()->setQuotePrefix(true);
}

$writer_prestasi_non_akademik = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_prestasi_non_akademik, 'Xls');
$writer_prestasi_non_akademik->save('./formatted/sapto_prestasi_mhs_non_akademik (F).xls');

$spreadsheet_prestasi_non_akademik->disconnectWorksheets();
unset($spreadsheet_prestasi_non_akademik);

// Load Format Baru
$spreadsheet_prestasi_non_akademik2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_prestasi_mhs_non_akademik (F).xls');
$worksheet_prestasi_non_akademik2 = $spreadsheet_prestasi_non_akademik2->getActiveSheet();

// Formasi Array SAPTO
$array_prestasi_non_akademik = $worksheet_prestasi_non_akademik2->toArray();
$data_prestasi_non_akademik = [];

foreach($worksheet_prestasi_non_akademik2->getRowIterator() as $row_id => $row) {
    if($worksheet_prestasi_non_akademik2->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['nama_kegiatan'] = $array_prestasi_non_akademik[$row_id-1][2];
            $item['tahun'] = $array_prestasi_non_akademik[$row_id-1][3];
			$item['lokal'] = $array_prestasi_non_akademik[$row_id-1][5];
			$item['nasional'] = $array_prestasi_non_akademik[$row_id-1][6];
			$item['internasional'] = $array_prestasi_non_akademik[$row_id-1][7];
			$item['capaian'] = $array_prestasi_non_akademik[$row_id-1][8];
            $data_prestasi_non_akademik[] = $item;
        }
    }
}

$spreadsheet_prestasi_non_akademik2->disconnectWorksheets();
unset($spreadsheet_prestasi_non_akademik2);


/**
 * Tabel 8.c Masa Studi Lulusan
 */

$sql_masa_studi = "SELECT a.id, a.prodi_id, a.jenjang, a.tahun_angkatan, a.jml_mhs_diterima, a.tahun_lulus, a.jml_lulusan, a.rata_lama_studi, b.nama_prodi FROM akreditasi.sapto_masa_studi_lulusan a INNER JOIN akreditasi.prodi_map b ON a.prodi_id = b.id WHERE b.nama_prodi = '".$nama_prodi."' ORDER BY a.tahun_angkatan DESC";
$stmt = sqlsrv_query( $conn, $sql_masa_studi );
if( $stmt === false) {
    die( print_r( sqlsrv_errors(), true) );
}

$data_masa_studi1 = [];

while( $row = sqlsrv_fetch_array( $stmt, SQLSRV_FETCH_NUMERIC) ) {
	 $data_array_masa_studi[] = array(
		$row[0], $row[1], $row[2], $row[3], $row[4], $row[5], $row[6], $row[7], $row[8]
	);
}

$data_masa_studi1 = array_merge($data_masa_studi1, $data_array_masa_studi); 

sqlsrv_free_stmt( $stmt);

$spreadsheet_masa_studi1 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./blank/sapto_masa_studi_lulusan.xlsx');
$worksheet_masa_studi1 = $spreadsheet_masa_studi1->getActiveSheet();

$worksheet_masa_studi1->fromArray($data_masa_studi1, NULL, 'A2');

$writer_masa_studi1 = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_masa_studi1, 'Xlsx');
$writer_masa_studi1->save('./raw/sapto_masa_studi_lulusan.xlsx');

$spreadsheet_masa_studi1->disconnectWorksheets();
unset($spreadsheet_masa_studi1);

$spreadsheet_masa_studi = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_masa_studi_lulusan.xlsx');
$worksheet_masa_studi = $spreadsheet_masa_studi->getActiveSheet();

$worksheet_masa_studi->insertNewColumnBefore('F', 7);

$worksheet_masa_studi->getCell('F1')->setValue('akhir TS-6');
$worksheet_masa_studi->getCell('G1')->setValue('akhir TS-5');
$worksheet_masa_studi->getCell('H1')->setValue('akhir TS-4');
$worksheet_masa_studi->getCell('I1')->setValue('akhir TS-3');
$worksheet_masa_studi->getCell('J1')->setValue('akhir TS-2');
$worksheet_masa_studi->getCell('K1')->setValue('akhir TS-1');
$worksheet_masa_studi->getCell('L1')->setValue('akhir TS');

$highestRow_masa_studi = $worksheet_masa_studi->getHighestRow();

for($row = 2;$row <= $highestRow_masa_studi; $row++) {
	$worksheet_masa_studi->setCellValue('H'.$row, '=IF(M'.$row.'=2016,N'.$row.',0)');
	$worksheet_masa_studi->getCell('H'.$row)->getStyle()->setQuotePrefix(true);
	
	$worksheet_masa_studi->setCellValue('I'.$row, '=IF(M'.$row.'=2017,N'.$row.',0)');
	$worksheet_masa_studi->getCell('I'.$row)->getStyle()->setQuotePrefix(true);
	
	$worksheet_masa_studi->setCellValue('J'.$row, '=IF(M'.$row.'=2018,N'.$row.',0)');
	$worksheet_masa_studi->getCell('J'.$row)->getStyle()->setQuotePrefix(true);
	
	$worksheet_masa_studi->setCellValue('K'.$row, '=IF(M'.$row.'=2019,N'.$row.',0)');
	$worksheet_masa_studi->getCell('K'.$row)->getStyle()->setQuotePrefix(true);
	
	$worksheet_masa_studi->setCellValue('L'.$row, '=IF(M'.$row.'=2020,N'.$row.',0)');
	$worksheet_masa_studi->getCell('L'.$row)->getStyle()->setQuotePrefix(true);
}

$writer_masa_studi = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_masa_studi, 'Xls');
$writer_masa_studi->save('./formatted/sapto_masa_studi_lulusan (F).xls');

$spreadsheet_masa_studi->disconnectWorksheets();
unset($spreadsheet_masa_studi);

// Load Format Baru
$spreadsheet_masa_studi2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_masa_studi_lulusan (F).xls');
$worksheet_masa_studi2 = $spreadsheet_masa_studi2->getActiveSheet();

// Formasi Array SAPTO
$array_masa_studi = $worksheet_masa_studi2->toArray();
$data_masa_studi = [];

foreach($worksheet_masa_studi2->getRowIterator() as $row_id => $row) {
    if($worksheet_masa_studi2->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
			if(substr($nama_prodi, 0, 3) == "S-1" || substr($nama_prodi, 0, 3) == "D-4" || substr($nama_prodi, 0, 3) == "S-3") {
				$item['tahun_masuk'] = $array_masa_studi[$row_id-1][3];
				$item['jml_mhs_diterima'] = $array_masa_studi[$row_id-1][4];
				$item['ts6_lulusan'] = $array_masa_studi[$row_id-1][5];
				$item['ts5_lulusan'] = $array_masa_studi[$row_id-1][6];
				$item['ts4_lulusan'] = $array_masa_studi[$row_id-1][7];
				$item['ts3_lulusan'] = $array_masa_studi[$row_id-1][8];
				$item['ts2_lulusan'] = $array_masa_studi[$row_id-1][9];
				$item['ts1_lulusan'] = $array_masa_studi[$row_id-1][10];
				$item['ts_lulusan'] = $array_masa_studi[$row_id-1][11];
				$item['jml_lulusan'] = $array_masa_studi[$row_id-1][13];
				$item['rata_masa_studi'] = $array_masa_studi[$row_id-1][14];
			} elseif(substr($nama_prodi, 0, 3) == "S-2") {
				$item['tahun_masuk'] = $array_masa_studi[$row_id-1][3];
				$item['jml_mhs_diterima'] = $array_masa_studi[$row_id-1][4];
				$item['ts3_lulusan'] = $array_masa_studi[$row_id-1][8];
				$item['ts2_lulusan'] = $array_masa_studi[$row_id-1][9];
				$item['ts1_lulusan'] = $array_masa_studi[$row_id-1][10];
				$item['ts_lulusan'] = $array_masa_studi[$row_id-1][11];
				$item['jml_lulusan'] = $array_masa_studi[$row_id-1][13];
				$item['rata_masa_studi'] = $array_masa_studi[$row_id-1][14];
			} elseif(substr($nama_prodi, 0, 3) == "D-3") {
				$item['tahun_masuk'] = $array_masa_studi[$row_id-1][3];
				$item['jml_mhs_diterima'] = $array_masa_studi[$row_id-1][4];
				$item['ts4_lulusan'] = $array_masa_studi[$row_id-1][7];
				$item['ts3_lulusan'] = $array_masa_studi[$row_id-1][8];
				$item['ts2_lulusan'] = $array_masa_studi[$row_id-1][9];
				$item['ts1_lulusan'] = $array_masa_studi[$row_id-1][10];
				$item['ts_lulusan'] = $array_masa_studi[$row_id-1][11];
				$item['jml_lulusan'] = $array_masa_studi[$row_id-1][13];
				$item['rata_masa_studi'] = $array_masa_studi[$row_id-1][14];
			}
			
            $data_masa_studi[] = $item;
        }
    }
}

$worksheet_masa_studi3 = new \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet($spreadsheet_masa_studi2, 'Sheet 2');
$spreadsheet_masa_studi2->addSheet($worksheet_masa_studi3);

$worksheet_masa_studi3 = $spreadsheet_masa_studi2->getSheetByName('Sheet 2');
$worksheet_masa_studi3->fromArray($data_masa_studi, NULL, 'A1');

$highestRow_masa_studi3 = $worksheet_masa_studi3->getHighestRow();

$row_jumlah = -1;

for($group = 1;$group <= 5; $group++) {

	// $ts6_lulusan = 0; $ts5_lulusan = 0; $ts4_lulusan = 0; $ts3_lulusan = 0; $ts2_lulusan = 0; $ts1_lulusan = 0;
	
	$row_jumlah += 2;
	$baris_awal = $row_jumlah;
	$ts4_lulusan = $worksheet_masa_studi3->getCell('E'.($row_jumlah))->getValue();
	$ts3_lulusan = $worksheet_masa_studi3->getCell('F'.($row_jumlah))->getValue();
	$ts2_lulusan = $worksheet_masa_studi3->getCell('G'.($row_jumlah))->getValue();
	$ts1_lulusan = $worksheet_masa_studi3->getCell('H'.($row_jumlah))->getValue();
	$ts_lulusan = $worksheet_masa_studi3->getCell('I'.($row_jumlah))->getValue(); 
	
	for($row = $row_jumlah;$row <= ($highestRow_masa_studi3+1); $row++) {
		if($worksheet_masa_studi3->getCell('A'.$row)->getValue() == $worksheet_masa_studi3->getCell('A'.($row+1))->getValue()) {
			$ts4_lulusan += $worksheet_masa_studi3->getCell('E'.($row+1))->getValue();
			$ts3_lulusan += $worksheet_masa_studi3->getCell('F'.($row+1))->getValue();
			$ts2_lulusan += $worksheet_masa_studi3->getCell('G'.($row+1))->getValue();
			$ts1_lulusan += $worksheet_masa_studi3->getCell('H'.($row+1))->getValue();
			$ts_lulusan += $worksheet_masa_studi3->getCell('I'.($row+1))->getValue();
			
			$row_jumlah++;
		} else {
			break;
		}
	}
	
	$worksheet_masa_studi3->insertNewRowBefore(($row_jumlah+1), 1);
	$worksheet_masa_studi3->setCellValue('B'.($row_jumlah+1), $worksheet_masa_studi3->getCell('B'.$row_jumlah)->getValue());
	$worksheet_masa_studi3->setCellValue('E'.($row_jumlah+1), $ts4_lulusan);
	$worksheet_masa_studi3->setCellValue('F'.($row_jumlah+1), $ts3_lulusan);
	$worksheet_masa_studi3->setCellValue('G'.($row_jumlah+1), $ts2_lulusan);
	$worksheet_masa_studi3->setCellValue('H'.($row_jumlah+1), $ts1_lulusan);	
	$worksheet_masa_studi3->setCellValue('I'.($row_jumlah+1), $ts_lulusan);
	$worksheet_masa_studi3->setCellValue('J'.($row_jumlah+1), '=SUM(E'.($row_jumlah+1).':I'.($row_jumlah+1).')');
	$worksheet_masa_studi3->setCellValue('K'.($row_jumlah+1), '=AVERAGE(K'.$baris_awal.':K'.$row_jumlah.')');
	
	${"masa_studi_angkatan".$group} = $worksheet_masa_studi3->rangeToArray('B'.($row_jumlah+1).':K'.($row_jumlah+1), NULL, TRUE, TRUE, TRUE);
}

$writer_masa_studi2 = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_masa_studi2, 'Xls');
$writer_masa_studi2->save('./formatted/sapto_masa_studi_lulusan (F).xls');

$spreadsheet_masa_studi2->disconnectWorksheets();
unset($spreadsheet_masa_studi2);



/**
 * Tabel 8.d.2 Kesesuaian Bidang Kerja Lulusan
 */

$sql_sesuai_kerja = "SELECT id, tahun_lulus, jml_lulusan, lulusan_terlacak, sesuai_rendah, sesuai_sedang, sesuai_tinggi, prodi_id FROM akreditasi.sapto_kesesuaian_kerja_lulusan WHERE prodi_id = '".$nama_prodi2."' ORDER BY tahun_lulus DESC";
$stmt = sqlsrv_query( $conn, $sql_sesuai_kerja );
if( $stmt === false) {
    die( print_r( sqlsrv_errors(), true) );
}

$data_sesuai_kerja1 = [];

while( $row = sqlsrv_fetch_array( $stmt, SQLSRV_FETCH_NUMERIC) ) {
	 $data_array_sesuai_kerja[] = array(
		$row[0], $row[1], $row[2], $row[3], $row[4], $row[5], $row[6], $row[7]
	);
}

$data_sesuai_kerja1 = array_merge($data_sesuai_kerja1, $data_array_sesuai_kerja); 

sqlsrv_free_stmt( $stmt);

$spreadsheet_sesuai_kerja1 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./blank/sapto_kesesuaian_kerja_lulusan.xlsx');
$worksheet_sesuai_kerja1 = $spreadsheet_sesuai_kerja1->getActiveSheet();

$worksheet_sesuai_kerja1->fromArray($data_sesuai_kerja1, NULL, 'A2');

$writer_sesuai_kerja1 = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_sesuai_kerja1, 'Xlsx');
$writer_sesuai_kerja1->save('./raw/sapto_kesesuaian_kerja_lulusan.xlsx');

$spreadsheet_sesuai_kerja1->disconnectWorksheets();
unset($spreadsheet_sesuai_kerja1);

$spreadsheet_sesuai_kerja = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_kesesuaian_kerja_lulusan.xlsx');
$worksheet_sesuai_kerja = $spreadsheet_sesuai_kerja->getActiveSheet();

$highestRow_sesuai_kerja = $worksheet_sesuai_kerja->getHighestRow();

$sesuai_kerja_ts = intval(date("Y"));
$sesuai_kerja_ts1 = intval(date("Y", strtotime("-1 year")));
$sesuai_kerja_ts2 = intval(date("Y", strtotime("-2 year")));

$worksheet_sesuai_kerja->setAutoFilter('B1:H'.$highestRow_sesuai_kerja);
$autoFilter_sesuai_kerja = $worksheet_sesuai_kerja->getAutoFilter();
$columnFilter_sesuai_kerja = $autoFilter_sesuai_kerja->getColumn('B');
$columnFilter_sesuai_kerja->setFilterType(
    \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_FILTER
);
$columnFilter_sesuai_kerja->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        $sesuai_kerja_ts2
    );
$columnFilter_sesuai_kerja->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        $sesuai_kerja_ts1
    );
$columnFilter_sesuai_kerja->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        $sesuai_kerja_ts
    );


$autoFilter_sesuai_kerja->showHideRows();

$writer_sesuai_kerja = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_sesuai_kerja, 'Xlsx');
$writer_sesuai_kerja->save('./formatted/sapto_kesesuaian_kerja_lulusan (F).xlsx');

$spreadsheet_sesuai_kerja->disconnectWorksheets();
unset($spreadsheet_sesuai_kerja);

// Load Format Baru
$spreadsheet_sesuai_kerja2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_kesesuaian_kerja_lulusan (F).xlsx');
$worksheet_sesuai_kerja2 = $spreadsheet_sesuai_kerja2->getActiveSheet();

// Formasi Array SAPTO
$array_sesuai_kerja = $worksheet_sesuai_kerja2->toArray();
$data_sesuai_kerja = [];

foreach($worksheet_sesuai_kerja2->getRowIterator() as $row_id => $row) {
    if($worksheet_sesuai_kerja2->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['tahun_lulus'] = $array_sesuai_kerja[$row_id-1][1];
            $item['jml_lulusan'] = $array_sesuai_kerja[$row_id-1][2];
			$item['jml_terlacak'] = $array_sesuai_kerja[$row_id-1][3];
			$item['sesuai_kurang'] = $array_sesuai_kerja[$row_id-1][4];
			$item['sesuai_sedang'] = $array_sesuai_kerja[$row_id-1][5];
			$item['sesuai_tinggi'] = $array_sesuai_kerja[$row_id-1][6];
            $data_sesuai_kerja[] = $item;
        }
    }
}

$spreadsheet_sesuai_kerja2->disconnectWorksheets();
unset($spreadsheet_sesuai_kerja2);


/**
 * Tabel 8.e.1 Tempat Kerja Lulusan
 */

$sql_tempat_kerja = "SELECT id, tahun_lulus, jml_lulusan, lulusan_terlacak, lokal, nasional, internasional, prodi_id FROM akreditasi.sapto_tempat_kerja_lulusan WHERE prodi_id = '".$nama_prodi2."' ORDER BY tahun_lulus DESC";
$stmt = sqlsrv_query( $conn, $sql_tempat_kerja );
if( $stmt === false) {
    die( print_r( sqlsrv_errors(), true) );
}

$data_tempat_kerja1 = [];

while( $row = sqlsrv_fetch_array( $stmt, SQLSRV_FETCH_NUMERIC) ) {
	 $data_array_tempat_kerja[] = array(
		$row[0], $row[1], $row[2], $row[3], $row[4], $row[5], $row[6], $row[7]
	);
}

$data_tempat_kerja1 = array_merge($data_tempat_kerja1, $data_array_sesuai_kerja); 

sqlsrv_free_stmt( $stmt);

$spreadsheet_tempat_kerja1 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./blank/sapto_tempat_kerja_lulusan.xlsx');
$worksheet_tempat_kerja1 = $spreadsheet_tempat_kerja1->getActiveSheet();

$worksheet_tempat_kerja1->fromArray($data_tempat_kerja1, NULL, 'A2');

$writer_tempat_kerja1 = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_tempat_kerja1, 'Xlsx');
$writer_tempat_kerja1->save('./raw/sapto_tempat_kerja_lulusan.xlsx');

$spreadsheet_tempat_kerja1->disconnectWorksheets();
unset($spreadsheet_tempat_kerja1);

$spreadsheet_tempat_kerja = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_tempat_kerja_lulusan.xlsx');
$worksheet_tempat_kerja = $spreadsheet_tempat_kerja->getActiveSheet();

$highestRow_tempat_kerja = $worksheet_tempat_kerja->getHighestRow();

$tempat_kerja_ts = intval(date("Y"));
$tempat_kerja_ts1 = intval(date("Y", strtotime("-1 year")));
$tempat_kerja_ts2 = intval(date("Y", strtotime("-2 year")));

$worksheet_tempat_kerja->setAutoFilter('B1:H'.$highestRow_sesuai_kerja);
$autoFilter_tempat_kerja = $worksheet_tempat_kerja->getAutoFilter();
$columnFilter_tempat_kerja = $autoFilter_tempat_kerja->getColumn('B');
$columnFilter_tempat_kerja->setFilterType(
    \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_FILTER
);
$columnFilter_tempat_kerja->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        $tempat_kerja_ts2
    );
$columnFilter_tempat_kerja->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        $tempat_kerja_ts1
    );
$columnFilter_tempat_kerja->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        $tempat_kerja_ts
    );


$autoFilter_tempat_kerja->showHideRows();

$writer_tempat_kerja = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_tempat_kerja, 'Xlsx');
$writer_tempat_kerja->save('./formatted/sapto_tempat_kerja_lulusan (F).xlsx');

$spreadsheet_tempat_kerja->disconnectWorksheets();
unset($spreadsheet_tempat_kerja);

// Load Format Baru
$spreadsheet_tempat_kerja2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_kesesuaian_tempat_lulusan (F).xlsx');
$worksheet_tempat_kerja2 = $spreadsheet_tempat_kerja2->getActiveSheet();

// Formasi Array SAPTO
$array_tempat_kerja = $worksheet_tempat_kerja2->toArray();
$data_tempat_kerja = [];

foreach($worksheet_tempat_kerja2->getRowIterator() as $row_id => $row) {
    if($worksheet_tempat_kerja2->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['tahun_lulus'] = $array_tempat_kerja[$row_id-1][1];
            $item['jml_lulusan'] = $array_tempat_kerja[$row_id-1][2];
			$item['jml_terlacak'] = $array_tempat_kerja[$row_id-1][3];
			$item['kerja_lokal'] = $array_tempat_kerja[$row_id-1][4];
			$item['kerja_nasional'] = $array_tempat_kerja[$row_id-1][5];
			$item['kerja_internasional'] = $array_tempat_kerja[$row_id-1][6];
            $data_tempat_kerja[] = $item;
        }
    }
}

$spreadsheet_tempat_kerja2->disconnectWorksheets();
unset($spreadsheet_tempat_kerja2);



/**
 * Referensi Tabel 8.e.2 Kepuasan Pengguna Lulusan
 */

$sql_ref_kepuasan_lulusan = "SELECT id, tahun_lulus, jml_lulusan, jml_tanggapan, prodi_id FROM akreditasi.sapto_ref_kepuasan_pengguna_lulusan WHERE prodi_id = '".$nama_prodi2."' ORDER BY tahun_lulus DESC";
$stmt = sqlsrv_query( $conn, $sql_ref_kepuasan_lulusan );
if( $stmt === false) {
    die( print_r( sqlsrv_errors(), true) );
}

$data_ref_kepuasan_lulusan1 = [];

while( $row = sqlsrv_fetch_array( $stmt, SQLSRV_FETCH_NUMERIC) ) {
	 $data_array_ref_kepuasan_lulusan[] = array(
		$row[0], $row[1], $row[2], $row[3], $row[4]
	);
}

$data_ref_kepuasan_lulusan1 = array_merge($data_ref_kepuasan_lulusan1, $data_array_ref_kepuasan_lulusan); 

sqlsrv_free_stmt( $stmt);

$spreadsheet_ref_kepuasan_lulusan1 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./blank/sapto_ref_kepuasan_pengguna_lulusan.xlsx');
$worksheet_ref_kepuasan_lulusan1 = $spreadsheet_ref_kepuasan_lulusan1->getActiveSheet();

$worksheet_ref_kepuasan_lulusan1->fromArray($data_ref_kepuasan_lulusan1, NULL, 'A2');

$writer_ref_kepuasan_lulusan1 = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_ref_kepuasan_lulusan1, 'Xlsx');
$writer_ref_kepuasan_lulusan1->save('./raw/sapto_ref_kepuasan_pengguna_lulusan.xlsx');

$spreadsheet_ref_kepuasan_lulusan1->disconnectWorksheets();
unset($spreadsheet_ref_kepuasan_lulusan1);

$spreadsheet_ref_kepuasan_lulusan = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_ref_kepuasan_pengguna_lulusan.xlsx');
$worksheet_ref_kepuasan_lulusan = $spreadsheet_ref_kepuasan_lulusan->getActiveSheet();

$highestRow_ref_kepuasan_lulusan = $worksheet_ref_kepuasan_lulusan->getHighestRow();

$ref_kepuasan_lulusan_ts = intval(date("Y"));
$ref_kepuasan_lulusan_ts1 = intval(date("Y", strtotime("-1 year")));
$ref_kepuasan_lulusan_ts2 = intval(date("Y", strtotime("-2 year")));

$worksheet_ref_kepuasan_lulusan->setAutoFilter('B1:E'.$highestRow_sesuai_kerja);
$autoFilter_ref_kepuasan_lulusan = $worksheet_ref_kepuasan_lulusan->getAutoFilter();
$columnFilter_ref_kepuasan_lulusan = $autoFilter_ref_kepuasan_lulusan->getColumn('B');
$columnFilter_ref_kepuasan_lulusan->setFilterType(
    \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_FILTER
);
$columnFilter_ref_kepuasan_lulusan->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        $ref_kepuasan_lulusan_ts2
    );
$columnFilter_ref_kepuasan_lulusan->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        $ref_kepuasan_lulusan_ts1
    );
$columnFilter_ref_kepuasan_lulusan->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        $ref_kepuasan_lulusan_ts
    );


$autoFilter_ref_kepuasan_lulusan->showHideRows();

$writer_ref_kepuasan_lulusan = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_ref_kepuasan_lulusan, 'Xlsx');
$writer_ref_kepuasan_lulusan->save('./formatted/sapto_ref_kepuasan_pengguna_lulusan (F).xlsx');

$spreadsheet_ref_kepuasan_lulusan->disconnectWorksheets();
unset($spreadsheet_ref_kepuasan_lulusan);

// Load Format Baru
$spreadsheet_ref_kepuasan_lulusan2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_ref_kepuasan_pengguna_lulusan (F).xlsx');
$worksheet_ref_kepuasan_lulusan2 = $spreadsheet_ref_kepuasan_lulusan2->getActiveSheet();

// Formasi Array SAPTO
$array_ref_kepuasan_lulusan = $worksheet_ref_kepuasan_lulusan2->toArray();
$data_ref_kepuasan_lulusan = [];

foreach($worksheet_ref_kepuasan_lulusan2->getRowIterator() as $row_id => $row) {
    if($worksheet_ref_kepuasan_lulusan2->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['tahun_lulus'] = $array_ref_kepuasan_lulusan[$row_id-1][1];
            $item['jml_lulusan'] = $array_ref_kepuasan_lulusan[$row_id-1][2];
			$item['jml_tanggapan'] = $array_ref_kepuasan_lulusan[$row_id-1][3];;
            $data_ref_kepuasan_lulusan[] = $item;
        }
    }
}

$spreadsheet_ref_kepuasan_lulusan2->disconnectWorksheets();
unset($spreadsheet_ref_kepuasan_lulusan2);



/**
 * Set ke Template SAPTO
 */

$spreadsheet_aps = \PhpOffice\PhpSpreadsheet\IOFactory::load('./result/sapto_aps9 (F).xlsx');


// IPK Lulusan
$worksheet_aps24 = $spreadsheet_aps->getSheetByName('8a');
$worksheet_aps24->fromArray($data_ipk_lulusan, NULL, 'B6');


// Prestasi Akademik Mahasiswa
$worksheet_aps25 = $spreadsheet_aps->getSheetByName('8b1');
$worksheet_aps25->fromArray($data_prestasi_akademik, NULL, 'B10');

$highestRow_aps25 = $worksheet_aps25->getHighestRow();

$worksheet_aps25->getStyle('A10:G'.$highestRow_aps25)->applyFromArray($styleBorder);
$worksheet_aps25->getStyle('B10:G'.$highestRow_aps25)->applyFromArray($styleYellow);
$worksheet_aps25->getStyle('A10:A'.$highestRow_aps25)->applyFromArray($styleCenter);
$worksheet_aps25->getStyle('D10:F'.$highestRow_aps25)->applyFromArray($styleCenter);
$worksheet_aps25->getStyle('B10:G'.$highestRow_aps25)->getAlignment()->setWrapText(true);

foreach($worksheet_aps25->getRowDimensions() as $rd25) { 
    $rd25->setRowHeight(-1); 
}

for($row = 10; $row <= $highestRow_aps25; $row++) {
	$worksheet_aps25->setCellValue('A'.$row, $row-9);
}


// Prestasi Non Akademik Mahasiswa
$worksheet_aps26 = $spreadsheet_aps->getSheetByName('8b2');
$worksheet_aps26->fromArray($data_prestasi_non_akademik, NULL, 'B11');

$highestRow_aps26 = $worksheet_aps26->getHighestRow();

$worksheet_aps26->getStyle('A11:G'.$highestRow_aps26)->applyFromArray($styleBorder);
$worksheet_aps26->getStyle('B11:G'.$highestRow_aps26)->applyFromArray($styleYellow);
$worksheet_aps26->getStyle('A11:A'.$highestRow_aps26)->applyFromArray($styleCenter);
$worksheet_aps26->getStyle('D11:F'.$highestRow_aps26)->applyFromArray($styleCenter);
$worksheet_aps26->getStyle('B11:G'.$highestRow_aps26)->getAlignment()->setWrapText(true);

foreach($worksheet_aps26->getRowDimensions() as $rd26) { 
    $rd26->setRowHeight(-1); 
}

for($row = 11; $row <= $highestRow_aps26; $row++) {
	$worksheet_aps26->setCellValue('A'.$row, $row-10);
}


// Kesesuaian Bidang Kerja Lulusan
$worksheet_aps27 = $spreadsheet_aps->getSheetByName('8d2');
$worksheet_aps27->fromArray($data_sesuai_kerja, NULL, 'A7');

$highestRow_aps27 = $worksheet_aps27->getHighestRow();

$worksheet_aps27->setCellValue('B10', '=SUM(B7:B9)');
$worksheet_aps27->setCellValue('C10', '=SUM(C7:C9)');
$worksheet_aps27->setCellValue('D10', '=SUM(D7:D9)');
$worksheet_aps27->setCellValue('E10', '=SUM(E7:E9)');
$worksheet_aps27->setCellValue('F10', '=SUM(F7:F9)');


// Tempat Kerja Lulusan
$worksheet_aps28 = $spreadsheet_aps->getSheetByName('8e1');
$worksheet_aps28->fromArray($data_tempat_kerja, NULL, 'A7');

$highestRow_aps28 = $worksheet_aps28->getHighestRow();

$worksheet_aps28->setCellValue('B10', '=SUM(B7:B9)');
$worksheet_aps28->setCellValue('C10', '=SUM(C7:C9)');
$worksheet_aps28->setCellValue('D10', '=SUM(D7:D9)');
$worksheet_aps28->setCellValue('E10', '=SUM(E7:E9)');
$worksheet_aps28->setCellValue('F10', '=SUM(F7:F9)');


// Referensi Kepuasan Pengguna Lulusan
$worksheet_aps45 = $spreadsheet_aps->getSheetByName('Ref 8e2');
$worksheet_aps45->fromArray($data_ref_kepuasan_lulusan, NULL, 'B7');

$highestRow_aps45 = $worksheet_aps45->getHighestRow();


// Kepuasan Pengguna Lulusan
$worksheet_aps29 = $spreadsheet_aps->getSheetByName('8e2');
$worksheet_aps29->fromArray($data_kepuasan_lulusan, NULL, 'B7');

$worksheet_aps29->setCellValue('C14', '=SUM(C7:C13)');
$worksheet_aps29->setCellValue('D14', '=SUM(D7:D13)');
$worksheet_aps29->setCellValue('E14', '=SUM(E7:E13)');
$worksheet_aps29->setCellValue('F14', '=SUM(F7:F13)');

foreach($worksheet_aps29->getRowDimensions() as $rd29) { 
    $rd29->setRowHeight(-1); 
}


$writer_aps = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_aps, 'Xlsx');
$writer_aps->save('./result/sapto_aps9 (F).xlsx');

$spreadsheet_aps->disconnectWorksheets();
unset($spreadsheet_aps);

?>