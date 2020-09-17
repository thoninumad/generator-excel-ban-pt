<?php

/*

ASPEK 3: PROFIL DOSEN

*/

require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

/* if ($argc < 2 )
{
    exit( "Usage: sapto_aspek3.php <id prodi sesuai di database>\n" );
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
 * Tabel 3.a.1 Dosen Tetap Perguruan Tinggi
 */

$sql_dosen_tetap = "SELECT nama, nidn_nidk, pend_s2, pend_s3, pendidikan_bidang, kesesuaian_kompetensi_inti_ps, nama_fungsional, no_sertifikat_pendidik, matkul, kesesuaian_bidang_keahlian_dengan_mk, id_homebase_dosen FROM akreditasi.sapto_dosen_tetap_matkul WHERE id_homebase_dosen = '".$nama_prodi."'";
$stmt = sqlsrv_query( $conn, $sql_dosen_tetap );
if( $stmt === false) {
    die( print_r( sqlsrv_errors(), true) );
}

$data_dosen_tetap1 = [];

while( $row = sqlsrv_fetch_array( $stmt, SQLSRV_FETCH_NUMERIC) ) {
	 $data_array_dosen_tetap[] = array(
		$row[0], $row[1], $row[2], $row[3], $row[4], $row[5], $row[6], $row[7], $row[8], $row[9], $row[10]
	);
}

$data_dosen_tetap1 = array_merge($data_dosen_tetap1, $data_array_dosen_tetap); 

sqlsrv_free_stmt( $stmt);

$spreadsheet_dosen_tetap1 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./blank/sapto_dosen_tetap_matkul.xlsx');
$worksheet_dosen_tetap1 = $spreadsheet_dosen_tetap1->getActiveSheet();

$worksheet_dosen_tetap1->fromArray($data_dosen_tetap1, NULL, 'A2');

$writer_dosen_tetap1 = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_dosen_tetap1, 'Xlsx');
$writer_dosen_tetap1->save('./raw/sapto_dosen_tetap_matkul.xlsx');

$spreadsheet_dosen_tetap1->disconnectWorksheets();
unset($spreadsheet_dosen_tetap1);

$spreadsheet_dosen_tetap = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_dosen_tetap_matkul.xlsx');

$worksheet_dosen_tetap = $spreadsheet_dosen_tetap->getActiveSheet();

$worksheet_dosen_tetap->insertNewColumnBefore('G', 1);
$worksheet_dosen_tetap->insertNewColumnBefore('J', 1);
$worksheet_dosen_tetap->insertNewColumnBefore('M', 2);

$highestRow_dosen_tetap = $worksheet_dosen_tetap->getHighestRow();

for($row = 2;$row <= $highestRow_dosen_tetap; $row++) {
	$worksheet_dosen_tetap->setCellValue('G'.$row, '=IF(F'.$row.'=1;"V";"")');
	$worksheet_dosen_tetap->getCell('G'.$row)->getStyle()->setQuotePrefix(true);
	
	$worksheet_dosen_tetap->setCellValue('M'.$row, '=IF(L'.$row.'=1;"V";"")');
	$worksheet_dosen_tetap->getCell('M'.$row)->getStyle()->setQuotePrefix(true);
}

$writer_dosen_tetap = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_dosen_tetap, 'Xls');
$writer_dosen_tetap->save('./formatted/sapto_dosen_tetap_matkul (F).xls');

$spreadsheet_dosen_tetap->disconnectWorksheets();
unset($spreadsheet_dosen_tetap);

// Load Format Baru
$spreadsheet_dosen_tetap2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_dosen_tetap_matkul (F).xls');
$worksheet_dosen_tetap2 = $spreadsheet_dosen_tetap2->getActiveSheet();

// Formasi Array SAPTO
$array_dosen_tetap = $worksheet_dosen_tetap2->toArray();
$data_dosen_tetap = [];

foreach($worksheet_dosen_tetap2->getRowIterator() as $row_id => $row) {
    if($worksheet_dosen_tetap2->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['nama_dosen'] = $array_dosen_tetap[$row_id-1][0];
            $item['nidn_nidk'] = $array_dosen_tetap[$row_id-1][1];
			$item['pendidikan_s2'] = $array_dosen_tetap[$row_id-1][2];
			$item['pendidikan_s3'] = $array_dosen_tetap[$row_id-1][3];
			$item['bidang_keahlian'] = $array_dosen_tetap[$row_id-1][4];
			$item['sesuai_kompetensi_inti'] = $array_dosen_tetap[$row_id-1][6];
			$item['jabatan_akademik'] = $array_dosen_tetap[$row_id-1][7];
			$item['sertifikat_pendidik'] = $array_dosen_tetap[$row_id-1][8];
			$item['sertifikat_kompetensi'] = $array_dosen_tetap[$row_id-1][9];
			$item['matkul_ps'] = $array_dosen_tetap[$row_id-1][10];
			$item['sesuai_bidang_keahlian'] = $array_dosen_tetap[$row_id-1][12];
			$item['matkul_ps_lain'] = $array_dosen_tetap[$row_id-1][13];
            $data_dosen_tetap[] = $item;
        }
    }
}

$spreadsheet_dosen_tetap2->disconnectWorksheets();
unset($spreadsheet_dosen_tetap2);


/**
 * Tabel 3.a.2 Dosen Pembimbing Utama TA
 */

$sql_dosen_pembimbing = "SELECT b.nama, a.tahun_akademik, a.id_mapping, a.jumlah, a.rata_prodi, a.rata_total, a.ps_key FROM akreditasi.sapto_dosen_pembimbing_nilai_rata a INNER JOIN akreditasi.dosen b ON a.nip_akademik = b.nip_akademik WHERE a.ps_key = '".$nama_prodi."' ORDER BY b.nama";
$stmt = sqlsrv_query( $conn, $sql_dosen_pembimbing );
if( $stmt === false) {
    die( print_r( sqlsrv_errors(), true) );
}

$data_dosen_pembimbing1 = [];

while( $row = sqlsrv_fetch_array( $stmt, SQLSRV_FETCH_NUMERIC) ) {
	 $data_array_dosen_pembimbing[] = array(
		$row[0], $row[1], $row[2], $row[3], $row[4], $row[5], $row[6]
	);
}

$data_dosen_pembimbing1 = array_merge($data_dosen_pembimbing1, $data_array_dosen_pembimbing); 

sqlsrv_free_stmt( $stmt);

$spreadsheet_dosen_pembimbing1 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./blank/sapto_dosen_pembimbing_nilai_rata.xlsx');
$worksheet_dosen_pembimbing1 = $spreadsheet_dosen_pembimbing1->getActiveSheet();

$worksheet_dosen_pembimbing1->fromArray($data_dosen_pembimbing1, NULL, 'A2');

$writer_dosen_pembimbing1 = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_dosen_pembimbing1, 'Xlsx');
$writer_dosen_pembimbing1->save('./raw/sapto_dosen_pembimbing_nilai_rata.xlsx');

$spreadsheet_dosen_pembimbing1->disconnectWorksheets();
unset($spreadsheet_dosen_pembimbing1);

$spreadsheet_dosen_pembimbing = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_dosen_pembimbing_nilai_rata.xlsx');

$worksheet_dosen_pembimbing = $spreadsheet_dosen_pembimbing->getActiveSheet();

$worksheet_dosen_pembimbing->insertNewColumnBefore('E', 8);

$worksheet_dosen_pembimbing->getCell('E1')->setValue('TS-2 Prodi');
$worksheet_dosen_pembimbing->getCell('F1')->setValue('TS-1 Prodi');
$worksheet_dosen_pembimbing->getCell('G1')->setValue('TS Prodi');
$worksheet_dosen_pembimbing->getCell('H1')->setValue('Rata Prodi');
$worksheet_dosen_pembimbing->getCell('I1')->setValue('TS-2 Prodi Lain');
$worksheet_dosen_pembimbing->getCell('J1')->setValue('TS-1 Prodi Lain');
$worksheet_dosen_pembimbing->getCell('K1')->setValue('TS Prodi Lain');
$worksheet_dosen_pembimbing->getCell('L1')->setValue('Rata Prodi Lain');

$highestRow_dosen_pembimbing = $worksheet_dosen_pembimbing->getHighestRow();

for($row = 2;$row <= $highestRow_dosen_pembimbing; $row++) {
	if($nama_prodi == $worksheet_dosen_pembimbing->getCell('C'.$row)->getValue()) {
		$worksheet_dosen_pembimbing->setCellValue('E'.$row, '=IF(B'.$row.'="TS-2",D'.$row.',0)');
		$worksheet_dosen_pembimbing->getCell('E'.$row)->getStyle()->setQuotePrefix(true);
		
		$worksheet_dosen_pembimbing->setCellValue('F'.$row, '=IF(B'.$row.'="TS-1",D'.$row.',0)');
		$worksheet_dosen_pembimbing->getCell('F'.$row)->getStyle()->setQuotePrefix(true);
		
		$worksheet_dosen_pembimbing->setCellValue('G'.$row, '=IF(B'.$row.'="TS",D'.$row.',0)');
		$worksheet_dosen_pembimbing->getCell('G'.$row)->getStyle()->setQuotePrefix(true);
		
		$worksheet_dosen_pembimbing->setCellValue('H'.$row, $worksheet_dosen_pembimbing->getCell('M'.$row)->getValue());
	} else {
		$worksheet_dosen_pembimbing->setCellValue('I'.$row, '=IF(B'.$row.'="TS-2",D'.$row.',0)');
		$worksheet_dosen_pembimbing->getCell('I'.$row)->getStyle()->setQuotePrefix(true);
		
		$worksheet_dosen_pembimbing->setCellValue('J'.$row, '=IF(B'.$row.'="TS-1",D'.$row.',0)');
		$worksheet_dosen_pembimbing->getCell('J'.$row)->getStyle()->setQuotePrefix(true);
		
		$worksheet_dosen_pembimbing->setCellValue('K'.$row, '=IF(B'.$row.'="TS",D'.$row.',0)');
		$worksheet_dosen_pembimbing->getCell('K'.$row)->getStyle()->setQuotePrefix(true);
	
		$worksheet_dosen_pembimbing->setCellValue('L'.$row, $worksheet_dosen_pembimbing->getCell('M'.$row)->getValue());
	}
}

$worksheet_dosen_pembimbing->setAutoFilter('A1:O'.$highestRow_dosen_pembimbing);
$autoFilter_dosen_pembimbing = $worksheet_dosen_pembimbing->getAutoFilter();
$columnFilter_dosen_pembimbing = $autoFilter_dosen_pembimbing->getColumn('B');
$columnFilter_dosen_pembimbing->setFilterType(
    \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_FILTER
);
$columnFilter_dosen_pembimbing->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        'TS-2'
    );
$columnFilter_dosen_pembimbing->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        'TS-1'
    );
$columnFilter_dosen_pembimbing->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        'TS'
    );

$autoFilter_dosen_pembimbing->showHideRows();

$writer_dosen_pembimbing = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_dosen_pembimbing, 'Xls');
$writer_dosen_pembimbing->save('./formatted/sapto_dosen_pembimbing_nilai_rata (F).xls');

$spreadsheet_dosen_pembimbing->disconnectWorksheets();
unset($spreadsheet_dosen_pembimbing);

// Load Format Baru
$spreadsheet_dosen_pembimbing2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_dosen_pembimbing_nilai_rata (F).xls');
$worksheet_dosen_pembimbing2 = $spreadsheet_dosen_pembimbing2->getActiveSheet();

// Formasi Array SAPTO
$array_dosen_pembimbing = $worksheet_dosen_pembimbing2->toArray();
$data_dosen_pembimbing = [];

foreach($worksheet_dosen_pembimbing2->getRowIterator() as $row_id => $row) {
    if($worksheet_dosen_pembimbing2->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['nama_dosen'] = $array_dosen_pembimbing[$row_id-1][0];
            $item['ts2_prodi'] = $array_dosen_pembimbing[$row_id-1][4];
			$item['ts1_prodi'] = $array_dosen_pembimbing[$row_id-1][5];
			$item['ts_prodi'] = $array_dosen_pembimbing[$row_id-1][6];
			$item['rata_prodi'] = $array_dosen_pembimbing[$row_id-1][7];
			$item['ts2_prodi_lain'] = $array_dosen_pembimbing[$row_id-1][8];
			$item['ts1_prodi_lain'] = $array_dosen_pembimbing[$row_id-1][9];
			$item['ts_prodi_lain'] = $array_dosen_pembimbing[$row_id-1][10];
			$item['rata_prodi_lain'] = $array_dosen_pembimbing[$row_id-1][11];
			$item['rata_total'] = $array_dosen_pembimbing[$row_id-1][13];
            $data_dosen_pembimbing[] = $item;
        }
    }
}

$worksheet_dosen_pembimbing3 = new \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet($spreadsheet_dosen_pembimbing2, 'Sheet 2');
$spreadsheet_dosen_pembimbing2->addSheet($worksheet_dosen_pembimbing3);

$worksheet_dosen_pembimbing3 = $spreadsheet_dosen_pembimbing2->getSheetByName('Sheet 2');
$worksheet_dosen_pembimbing3->fromArray($data_dosen_pembimbing, NULL, 'A1');

$highestRow_dosen_pembimbing3 = $worksheet_dosen_pembimbing3->getHighestRow();

$row_jumlah_pembimbing = -1;

for($group = 1;$group <= 100; $group++) {
	
	$row_jumlah_pembimbing += 2;
	$baris_awal_pembimbing = $row_jumlah_pembimbing;
	
	$ts2_prodi = $worksheet_dosen_pembimbing3->getCell('B'.($row_jumlah_pembimbing))->getValue(); 
	$ts1_prodi = $worksheet_dosen_pembimbing3->getCell('C'.($row_jumlah_pembimbing))->getValue();
	$ts_prodi = $worksheet_dosen_pembimbing3->getCell('D'.($row_jumlah_pembimbing))->getValue();
	$ts2_prodi_lain = $worksheet_dosen_pembimbing3->getCell('F'.($row_jumlah_pembimbing))->getValue();
	$ts1_prodi_lain = $worksheet_dosen_pembimbing3->getCell('G'.($row_jumlah_pembimbing))->getValue();
	$ts_prodi_lain = $worksheet_dosen_pembimbing3->getCell('H'.($row_jumlah_pembimbing))->getValue();
	
	for($row = $row_jumlah_pembimbing;$row <= ($highestRow_dosen_pembimbing3+1); $row++) {
		if($worksheet_dosen_pembimbing3->getCell('A'.$row)->getValue() == $worksheet_dosen_pembimbing3->getCell('A'.($row+1))->getValue()) {
			$ts2_prodi += $worksheet_dosen_pembimbing3->getCell('B'.($row+1))->getValue();
			$ts1_prodi += $worksheet_dosen_pembimbing3->getCell('C'.($row+1))->getValue();
			$ts_prodi += $worksheet_dosen_pembimbing3->getCell('D'.($row+1))->getValue();
			$ts2_prodi_lain += $worksheet_dosen_pembimbing3->getCell('F'.($row+1))->getValue();
			$ts1_prodi_lain += $worksheet_dosen_pembimbing3->getCell('G'.($row+1))->getValue();
			$ts_prodi_lain += $worksheet_dosen_pembimbing3->getCell('H'.($row+1))->getValue();
			
			$row_jumlah_pembimbing++;
		} else {
			break;
		}
	}
	
	$worksheet_dosen_pembimbing3->insertNewRowBefore(($row_jumlah_pembimbing+1), 1);
	$worksheet_dosen_pembimbing3->setCellValue('A'.($row_jumlah_pembimbing+1), $worksheet_dosen_pembimbing3->getCell('A'.$row_jumlah_pembimbing)->getValue());
	$worksheet_dosen_pembimbing3->setCellValue('B'.($row_jumlah_pembimbing+1), $ts2_prodi);
	$worksheet_dosen_pembimbing3->setCellValue('C'.($row_jumlah_pembimbing+1), $ts1_prodi);
	$worksheet_dosen_pembimbing3->setCellValue('D'.($row_jumlah_pembimbing+1), $ts_prodi);
	$worksheet_dosen_pembimbing3->setCellValue('E'.($row_jumlah_pembimbing+1), '=IF(ISERROR(AVERAGE(B'.($row_jumlah_pembimbing+1).':D'.($row_jumlah_pembimbing+1).')),"",AVERAGE(B'.($row_jumlah_pembimbing+1).':D'.($row_jumlah_pembimbing+1).'))');
	$worksheet_dosen_pembimbing3->setCellValue('F'.($row_jumlah_pembimbing+1), $ts2_prodi_lain);	
	$worksheet_dosen_pembimbing3->setCellValue('G'.($row_jumlah_pembimbing+1), $ts1_prodi_lain);
	$worksheet_dosen_pembimbing3->setCellValue('H'.($row_jumlah_pembimbing+1), $ts_prodi_lain);
	$worksheet_dosen_pembimbing3->setCellValue('I'.($row_jumlah_pembimbing+1), '=IF(ISERROR(AVERAGE(F'.($row_jumlah_pembimbing+1).':H'.($row_jumlah_pembimbing+1).')),"",AVERAGE(F'.($row_jumlah_pembimbing+1).':H'.($row_jumlah_pembimbing+1).'))');
	$worksheet_dosen_pembimbing3->setCellValue('J'.($row_jumlah_pembimbing+1), $worksheet_dosen_pembimbing3->getCell('J'.$row_jumlah_pembimbing)->getValue());
	
	${"dosen_pembimbing".$group} = $worksheet_dosen_pembimbing3->rangeToArray('A'.($row_jumlah_pembimbing+1).':J'.($row_jumlah_pembimbing+1), NULL, TRUE, TRUE, TRUE);
}

$writer_dosen_pembimbing2 = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_dosen_pembimbing2, 'Xls');
$writer_dosen_pembimbing2->save('./formatted/sapto_dosen_pembimbing_nilai_rata (F).xls');

$spreadsheet_dosen_pembimbing2->disconnectWorksheets();
unset($spreadsheet_dosen_pembimbing2); 



/**
 * Tabel 3.a.3 EWMP DTPS
 */

$sql_ewmp = "SELECT id, nama_dosen, dtps, ps, ps_lain_dalam, ps_lain_luar, penelitian, pkm, tugas_tambahan, jml_sks, rata_sks_smt, id_prodi FROM akreditasi.sapto_ewmp_dosen_tetap WHERE id_prodi = '".$nama_prodi."'";
$stmt = sqlsrv_query( $conn, $sql_ewmp );
if( $stmt === false) {
    die( print_r( sqlsrv_errors(), true) );
}

$data_ewmp1 = [];

while( $row = sqlsrv_fetch_array( $stmt, SQLSRV_FETCH_NUMERIC) ) {
	 $data_array_ewmp[] = array(
		$row[0], $row[1], $row[2], $row[3], $row[4], $row[5], $row[6], $row[7], $row[8], $row[9], $row[10], $row[11]
	);
}

$data_ewmp1 = array_merge($data_ewmp1, $data_array_ewmp); 

sqlsrv_free_stmt( $stmt);

$spreadsheet_ewmp1 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./blank/sapto_ewmp_dosen_tetap.xlsx');
$worksheet_ewmp1 = $spreadsheet_ewmp1->getActiveSheet();

$worksheet_ewmp1->fromArray($data_ewmp1, NULL, 'A2');

$writer_ewmp1 = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_ewmp1, 'Xlsx');
$writer_ewmp1->save('./raw/sapto_ewmp_dosen_tetap.xlsx');

$spreadsheet_ewmp1->disconnectWorksheets();
unset($spreadsheet_ewmp1);

$spreadsheet_ewmp = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_ewmp_dosen_tetap.xlsx');

$worksheet_ewmp = $spreadsheet_ewmp->getActiveSheet();

$worksheet_ewmp->insertNewColumnBefore('D', 1);

$highestRow_ewmp = $worksheet_ewmp->getHighestRow();

for($row = 2;$row <= $highestRow_ewmp; $row++) {
	$worksheet_ewmp->setCellValue('D'.$row, '=IF(C'.$row.'=1;"V";"")');
	$worksheet_ewmp->getCell('D'.$row)->getStyle()->setQuotePrefix(true);
}

$writer_ewmp = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_ewmp, 'Xls');
$writer_ewmp->save('./formatted/sapto_ewmp_dosen_tetap (F).xls');

$spreadsheet_ewmp->disconnectWorksheets();
unset($spreadsheet_ewmp);

// Load Format Baru
$spreadsheet_ewmp2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_ewmp_dosen_tetap (F).xls');
$worksheet_ewmp2 = $spreadsheet_ewmp2->getActiveSheet();

// Formasi Array SAPTO
$array_ewmp = $worksheet_ewmp2->toArray();
$data_ewmp = [];

foreach($worksheet_ewmp2->getRowIterator() as $row_id => $row) {
    if($worksheet_ewmp2->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['nama_dosen'] = $array_ewmp[$row_id-1][1];
            $item['dtps'] = $array_ewmp[$row_id-1][3];
			$item['ewmp_ps'] = $array_ewmp[$row_id-1][4];
			$item['ewmp_ps_lain'] = $array_ewmp[$row_id-1][5];
			$item['ewmp_ps_luar'] = $array_ewmp[$row_id-1][6];
			$item['ewmp_penelitian'] = $array_ewmp[$row_id-1][7];
			$item['ewmp_pkm'] = $array_ewmp[$row_id-1][8];
			$item['ewmp_tambahan'] = $array_ewmp[$row_id-1][9];
			$item['jumlah'] = $array_ewmp[$row_id-1][10];
			$item['rata2'] = $array_ewmp[$row_id-1][11];
            $data_ewmp[] = $item;
        }
    }
}

$spreadsheet_ewmp2->disconnectWorksheets();
unset($spreadsheet_ewmp2);


/**
 * Tabel 3.a.4 Dosen Tidak Tetap
 */

$sql_dosen_tidaktetap = "SELECT nama, nidn_nidk, pend_s2, pend_s3, pendidikan_bidang, nama_fungsional, no_sertifikat_pendidik, matkul, kesesuaian_bidang_keahlian_dengan_mk, id_homebase_dosen FROM akreditasi.sapto_dosen_tidaktetap_matkul WHERE id_homebase_dosen = '".$nama_prodi."'";
$stmt = sqlsrv_query( $conn, $sql_dosen_tidaktetap );
if( $stmt === false) {
    die( print_r( sqlsrv_errors(), true) );
}

$data_dosen_tidaktetap1 = [];

while( $row = sqlsrv_fetch_array( $stmt, SQLSRV_FETCH_NUMERIC) ) {
	 $data_array_dosen_tidaktetap[] = array(
		$row[0], $row[1], $row[2], $row[3], $row[4], $row[5], $row[6], $row[7], $row[8], $row[9]
	);
}

$data_dosen_tidaktetap1 = array_merge($data_dosen_tidaktetap1, $data_array_dosen_tidaktetap); 

sqlsrv_free_stmt( $stmt);

$spreadsheet_dosen_tidaktetap1 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./blank/sapto_dosen_tidaktetap_matkul.xlsx');
$worksheet_dosen_tidaktetap1 = $spreadsheet_dosen_tidaktetap1->getActiveSheet();

$worksheet_dosen_tidaktetap1->fromArray($data_dosen_tidaktetap1, NULL, 'A2');

$writer_dosen_tidaktetap1 = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_dosen_tidaktetap1, 'Xlsx');
$writer_dosen_tidaktetap1->save('./raw/sapto_dosen_tidaktetap_matkul.xlsx');

$spreadsheet_dosen_tidaktetap1->disconnectWorksheets();
unset($spreadsheet_dosen_tidaktetap1);

$spreadsheet_dosen_tidaktetap = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_dosen_tidaktetap_matkul.xlsx');

$worksheet_dosen_tidaktetap = $spreadsheet_dosen_tidaktetap->getActiveSheet();

$worksheet_dosen_tidaktetap->insertNewColumnBefore('E', 1);
$worksheet_dosen_tidaktetap->insertNewColumnBefore('I', 1);
$worksheet_dosen_tidaktetap->insertNewColumnBefore('L', 1);

$worksheet_dosen_tidaktetap->getCell('E1')->setValue('Pendidikan Pasca Sarjana');
$worksheet_dosen_tidaktetap->getCell('I1')->setValue('Sertifikat Kompetensi/Profesi/Industri');
$worksheet_dosen_tidaktetap->getCell('L1')->setValue('Kesesuaian Bidang Keahlian dengan Mata Kuliah yang Diampu');

$highestRow_dosen_tidaktetap = $worksheet_dosen_tidaktetap->getHighestRow();

for($row = 2;$row <= $highestRow_dosen_tidaktetap; $row++) {
	$worksheet_dosen_tidaktetap->setCellValue('E'.$row, '=IF(D'.$row.'<>"";D'.$row.';C'.$row.')');
	$worksheet_dosen_tidaktetap->getCell('E'.$row)->getStyle()->setQuotePrefix(true);
	
	$worksheet_dosen_tidaktetap->setCellValue('L'.$row, '=IF(K'.$row.'=1;"V";"")');
	$worksheet_dosen_tidaktetap->getCell('L'.$row)->getStyle()->setQuotePrefix(true);
}

$writer_dosen_tidaktetap = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_dosen_tidaktetap, 'Xls');
$writer_dosen_tidaktetap->save('./formatted/sapto_dosen_tidaktetap_matkul (F).xls');

$spreadsheet_dosen_tidaktetap->disconnectWorksheets();
unset($spreadsheet_dosen_tidaktetap);

// Load Format Baru
$spreadsheet_dosen_tidaktetap2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_dosen_tidaktetap_matkul (F).xls');
$worksheet_dosen_tidaktetap2 = $spreadsheet_dosen_tidaktetap2->getActiveSheet();

// Formasi Array SAPTO
$array_dosen_tidaktetap = $worksheet_dosen_tidaktetap2->toArray();
$data_dosen_tidaktetap = [];

foreach($worksheet_dosen_tidaktetap2->getRowIterator() as $row_id => $row) {
    if($worksheet_dosen_tidaktetap2->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['nama_dosen'] = $array_dosen_tidaktetap[$row_id-1][0];
            $item['nidn_nidk'] = $array_dosen_tidaktetap[$row_id-1][1];
			$item['pendidikan_pasca'] = $array_dosen_tidaktetap[$row_id-1][4];
			$item['bidang_keahlian'] = $array_dosen_tidaktetap[$row_id-1][5];
			$item['jabatan_akademik'] = $array_dosen_tidaktetap[$row_id-1][6];
			$item['sertifikat_pendidik'] = $array_dosen_tidaktetap[$row_id-1][7];
			$item['sertifikat_kompetensi'] = $array_dosen_tidaktetap[$row_id-1][8];
			$item['matkul_ps'] = $array_dosen_tidaktetap[$row_id-1][9];
			$item['sesuai_bidang_keahlian'] = $array_dosen_tidaktetap[$row_id-1][11];
            $data_dosen_tidaktetap[] = $item;
        }
    }
}

$spreadsheet_dosen_tidaktetap2->disconnectWorksheets();
unset($spreadsheet_dosen_tidaktetap2);


/**
 * Tabel 3.a.5 Dosen Industri/Praktisi
 */

$sql_dosen_industri = "SELECT a.id, a.nip_kepegawaian, a.nama, b.nidn_nidk, a.perusahaan_industri, b.pend_s2, b.pend_s3, b.pendidikan_bidang, a.matkul_diampu, a.sks, b.id_homebase_dosen FROM akreditasi.sapto_dosen_industri a INNER JOIN akreditasi.sapto_dosen_tidaktetap_matkul b ON a.nip_kepegawaian = b.nip_kepegawaian WHERE b.id_homebase_dosen = '".$nama_prodi."'";
$stmt = sqlsrv_query( $conn, $sql_dosen_industri );
if( $stmt === false) {
    die( print_r( sqlsrv_errors(), true) );
}

$data_dosen_industri1 = [];

while( $row = sqlsrv_fetch_array( $stmt, SQLSRV_FETCH_NUMERIC) ) {
	 $data_array_dosen_industri[] = array(
		$row[0], $row[1], $row[2], $row[3], $row[4], $row[5], $row[6], $row[7], $row[8], $row[9], $row[10], $row[11]
	);
}

$data_dosen_industri1 = array_merge($data_dosen_industri1, $data_array_dosen_industri); 

sqlsrv_free_stmt( $stmt);

$spreadsheet_dosen_industri1 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./blank/sapto_dosen_industri.xlsx');
$worksheet_dosen_industri1 = $spreadsheet_dosen_industri1->getActiveSheet();

$worksheet_dosen_industri1->fromArray($data_dosen_industri1, NULL, 'A2');

$writer_dosen_industri1 = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_dosen_industri1, 'Xlsx');
$writer_dosen_industri1->save('./raw/sapto_dosen_industri.xlsx');

$spreadsheet_dosen_industri1->disconnectWorksheets();
unset($spreadsheet_industri1);

$spreadsheet_dosen_industri = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_dosen_industri.xlsx');

$worksheet_dosen_industri = $spreadsheet_dosen_industri->getActiveSheet();

$worksheet_dosen_industri->insertNewColumnBefore('H', 1);
$worksheet_dosen_industri->insertNewColumnBefore('J', 1);

$worksheet_dosen_industri->getCell('H1')->setValue('Pendidikan Tertinggi');
$worksheet_dosen_industri->getCell('J1')->setValue('Sertifikat Kompetensi/Profesi/Industri');

$highestRow_dosen_industri = $worksheet_dosen_industri->getHighestRow();

for($row = 2;$row <= $highestRow_dosen_industri; $row++) {
	$worksheet_dosen_industri->setCellValue('H'.$row, '=IF(G'.$row.'<>"";G'.$row.';F'.$row.')');
	$worksheet_dosen_industri->getCell('H'.$row)->getStyle()->setQuotePrefix(true);
}

$writer_dosen_industri = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_dosen_industri, 'Xls');
$writer_dosen_industri->save('./formatted/sapto_dosen_industri (F).xls');

$spreadsheet_dosen_industri->disconnectWorksheets();
unset($spreadsheet_dosen_industri);

// Load Format Baru
$spreadsheet_dosen_industri2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_dosen_industri (F).xlsx');
$worksheet_dosen_industri2 = $spreadsheet_dosen_industri2->getActiveSheet();

// Formasi Array SAPTO
$array_dosen_industri = $worksheet_dosen_industri2->toArray();
$data_dosen_industri = [];

foreach($worksheet_dosen_industri2->getRowIterator() as $row_id => $row) {
    if($worksheet_dosen_industri2->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['nama_dosen'] = $array_dosen_industri[$row_id-1][2];
            $item['nidn_nidk'] = $array_dosen_industri[$row_id-1][3];
			$item['perusahaan'] = $array_dosen_industri[$row_id-1][4];
			$item['pendidikan_tertinggi'] = $array_dosen_industri[$row_id-1][7];
			$item['bidang_keahlian'] = $array_dosen_industri[$row_id-1][8];
			$item['sertifikat_kompetensi'] = $array_dosen_industri[$row_id-1][9];
			$item['matkul_ps'] = $array_dosen_industri[$row_id-1][10];
			$item['sks'] = $array_dosen_industri[$row_id-1][11];
            $data_dosen_industri[] = $item;
        }
    }
}

$spreadsheet_dosen_industri2->disconnectWorksheets();
unset($spreadsheet_dosen_industri2);


/**
 * Set ke Template SAPTO
 */

$spreadsheet_aps = \PhpOffice\PhpSpreadsheet\IOFactory::load('./result/sapto_aps9 (F).xlsx');


// Dosen Tetap Perguruan Tinggi
$worksheet_aps5 = $spreadsheet_aps->getSheetByName('3a1');
$worksheet_aps5->fromArray($data_dosen_tetap, NULL, 'B14');

$highestRow_aps5 = $worksheet_aps5->getHighestRow();
$worksheet_aps5->getStyle('A14:M'.$highestRow_aps5)->applyFromArray($styleBorder);
$worksheet_aps5->getStyle('B14:M'.$highestRow_aps5)->applyFromArray($styleYellow);
$worksheet_aps5->getStyle('A14:A'.$highestRow_aps5)->applyFromArray($styleCenter);
$worksheet_aps5->getStyle('G14:J'.$highestRow_aps5)->applyFromArray($styleCenter);
$worksheet_aps5->getStyle('L14:L'.$highestRow_aps5)->applyFromArray($styleCenter);
$worksheet_aps5->getStyle('B14:M'.$highestRow_aps5)->getAlignment()->setWrapText(true);

foreach($worksheet_aps5->getRowDimensions() as $rd5) { 
    $rd5->setRowHeight(-1); 
}

$worksheet_aps5->getStyle('C14:C'.$highestRow_aps5)->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_NUMBER);
$worksheet_aps5->getStyle('I14:J'.$highestRow_aps5)->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_NUMBER);

for($row = 14; $row <= $highestRow_aps5; $row++) {
	$worksheet_aps5->setCellValue('A'.$row, $row-13);
}


// Dosen Pembimbing Utama TA
$worksheet_aps39 = $spreadsheet_aps->getSheetByName('3a2');

for($group = 1; $group <= 100; $group++) {
	$worksheet_aps39->fromArray(${"dosen_pembimbing".$group}, NULL, 'B'.($group+6));
}
	
$highestRow_aps39 = $worksheet_aps39->getHighestRow();

$worksheet_aps39->getStyle('A7:K'.$highestRow_aps39)->applyFromArray($styleBorder);
$worksheet_aps39->getStyle('B7:K'.$highestRow_aps39)->applyFromArray($styleYellow);
$worksheet_aps39->getStyle('A7:A'.$highestRow_aps39)->applyFromArray($styleCenter);
$worksheet_aps39->getStyle('C7:K'.$highestRow_aps39)->applyFromArray($styleCenter);
$worksheet_aps39->getStyle('A3:K'.$highestRow_aps39)->getAlignment()->setWrapText(true);

$worksheet_aps39->getStyle('F7:F'.$highestRow_aps39)->getNumberFormat()->setFormatCode('0.00'); 
$worksheet_aps39->getStyle('J7:J'.$highestRow_aps39)->getNumberFormat()->setFormatCode('0.00');
$worksheet_aps39->getStyle('K7:K'.$highestRow_aps39)->getNumberFormat()->setFormatCode('0.00');  

foreach($worksheet_aps39->getRowDimensions() as $rd39) { 
    $rd39->setRowHeight(-1); 
}

for($row = 7; $row <= $highestRow_aps39; $row++) {
	$worksheet_aps39->setCellValue('A'.$row, $row-6);
}


// EWMP DTPS
$worksheet_aps6 = $spreadsheet_aps->getSheetByName('3a3');
$worksheet_aps6->fromArray($data_ewmp, NULL, 'B11');

$highestRow_aps6 = $worksheet_aps6->getHighestRow();
$worksheet_aps6->getStyle('A11:K'.$highestRow_aps6)->applyFromArray($styleBorder);
$worksheet_aps6->getStyle('B11:K'.$highestRow_aps6)->applyFromArray($styleYellow);
$worksheet_aps6->getStyle('A11:A'.$highestRow_aps6)->applyFromArray($styleCenter);
$worksheet_aps6->getStyle('C11:K'.$highestRow_aps6)->applyFromArray($styleCenter);
$worksheet_aps6->getStyle('B11:K'.$highestRow_aps6)->getAlignment()->setWrapText(true);

foreach($worksheet_aps6->getRowDimensions() as $rd6) { 
    $rd6->setRowHeight(-1); 
}

for($row = 11; $row <= $highestRow_aps6; $row++) {
	$worksheet_aps6->setCellValue('A'.$row, $row-10);
}


// Dosen Tidak Tetap
$worksheet_aps7 = $spreadsheet_aps->getSheetByName('3a4');
$worksheet_aps7->fromArray($data_dosen_tidaktetap, NULL, 'B13');

$highestRow_aps7 = $worksheet_aps7->getHighestRow();
$worksheet_aps7->getStyle('A13:J'.$highestRow_aps7)->applyFromArray($styleBorder);
$worksheet_aps7->getStyle('B13:J'.$highestRow_aps7)->applyFromArray($styleYellow);
$worksheet_aps7->getStyle('A13:A'.$highestRow_aps7)->applyFromArray($styleCenter);
$worksheet_aps7->getStyle('C13:J'.$highestRow_aps7)->applyFromArray($styleCenter);
$worksheet_aps7->getStyle('B13:J'.$highestRow_aps7)->getAlignment()->setWrapText(true);

foreach($worksheet_aps7->getRowDimensions() as $rd7) { 
    $rd7->setRowHeight(-1); 
}

$worksheet_aps7->getStyle('C13:C'.$highestRow_aps7)->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_NUMBER);
$worksheet_aps7->getStyle('G13:H'.$highestRow_aps7)->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_NUMBER);

for($row = 13; $row <= $highestRow_aps7; $row++) {
	$worksheet_aps7->setCellValue('A'.$row, $row-12);
}


// Dosen Industri/Praktisi
$worksheet_aps41 = $spreadsheet_aps->getSheetByName('3a5');
$worksheet_aps41->fromArray($data_dosen_industri, NULL, 'B6');

$highestRow_aps41 = $worksheet_aps41->getHighestRow();
$worksheet_aps41->getStyle('A6:I'.$highestRow_aps41)->applyFromArray($styleBorder);
$worksheet_aps41->getStyle('B6:I'.$highestRow_aps41)->applyFromArray($styleYellow);
$worksheet_aps41->getStyle('A6:A'.$highestRow_aps41)->applyFromArray($styleCenter);
$worksheet_aps41->getStyle('C6:I'.$highestRow_aps41)->applyFromArray($styleCenter);
$worksheet_aps41->getStyle('B6:I'.$highestRow_aps41)->getAlignment()->setWrapText(true);

foreach($worksheet_aps41->getRowDimensions() as $rd41) { 
    $rd41->setRowHeight(-1); 
}

$worksheet_aps41->getStyle('C6:C'.$highestRow_aps41)->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_NUMBER);
$worksheet_aps41->getStyle('G6:G'.$highestRow_aps41)->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_NUMBER);

for($row = 6; $row <= $highestRow_aps41; $row++) {
	$worksheet_aps41->setCellValue('A'.$row, $row-5);
}


$writer_aps = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_aps, 'Xlsx');
$writer_aps->save('./result/sapto_aps9 (F).xlsx');

$spreadsheet_aps->disconnectWorksheets();
unset($spreadsheet_aps);

?>