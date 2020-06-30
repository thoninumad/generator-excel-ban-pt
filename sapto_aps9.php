<?php

require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

/* if ($argc < 2 )
{
    exit( "Usage: sapto_aps9.php <nama prodi sesuai di database>\n" );
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
$nama_prodi2 = 'S-1 TEKNIK MESIN  - T.Mesin - Fakultas Teknologi Industri dan Rekayasa Sistem';

/**
 * Tabel 1.1 Kerjasama Tridharma Pendidikan
 */

$spreadsheet_tridharma_pendidikan = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_tridharma_pendidikan.xls');

$worksheet_tridharma_pendidikan = $spreadsheet_tridharma_pendidikan->getActiveSheet();

$worksheet_tridharma_pendidikan->insertNewColumnBefore('C', 3);

$highestRow_tridharma_pendidikan = $worksheet_tridharma_pendidikan->getHighestRow();

for($row = 2;$row <= $highestRow_tridharma_pendidikan; $row++) {
	$worksheet_tridharma_pendidikan->setCellValue('C'.$row, '=IF(F'.$row.'="Internasional";"V";"")');
	$worksheet_tridharma_pendidikan->getCell('C'.$row)->getStyle()->setQuotePrefix(true);
	
	$worksheet_tridharma_pendidikan->setCellValue('D'.$row, '=IF(F'.$row.'="Nasional";"V";"")');
	$worksheet_tridharma_pendidikan->getCell('D'.$row)->getStyle()->setQuotePrefix(true);
	
	$worksheet_tridharma_pendidikan->setCellValue('E'.$row, '=IF(F'.$row.'="Lokal";"V";"")');
	$worksheet_tridharma_pendidikan->getCell('E'.$row)->getStyle()->setQuotePrefix(true);
}

$worksheet_tridharma_pendidikan->setAutoFilter('B1:M'.$highestRow_tridharma_pendidikan);
$autoFilter_tridharma_pendidikan = $worksheet_tridharma_pendidikan->getAutoFilter();
$columnFilter_tridharma_pendidikan = $autoFilter_tridharma_pendidikan->getColumn('L');
$columnFilter_tridharma_pendidikan->setFilterType(
    \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_FILTER
);
$columnFilter_tridharma_pendidikan->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        $nama_prodi2
    );

$autoFilter_tridharma_pendidikan->showHideRows();

$writer_tridharma_pendidikan = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_tridharma_pendidikan, 'Xls');
$writer_tridharma_pendidikan->save('./formatted/sapto_tridharma_pendidikan (F).xls');

$spreadsheet_tridharma_pendidikan->disconnectWorksheets();
unset($spreadsheet_tridharma_pendidikan);

// Load Format Baru
$spreadsheet_tridharma_pendidikan2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_tridharma_pendidikan (F).xls');
$worksheet_tridharma_pendidikan2 = $spreadsheet_tridharma_pendidikan2->getActiveSheet();

// May be change
$array_tridharma_pendidikan = $worksheet_tridharma_pendidikan2->toArray();
$data_tridharma_pendidikan = [];

foreach($worksheet_tridharma_pendidikan2->getRowIterator() as $row_id => $row) {
    if($worksheet_tridharma_pendidikan2->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['lembaga_mitra'] = $array_tridharma_pendidikan[$row_id-1][1];
            $item['lokal'] = $array_tridharma_pendidikan[$row_id-1][2];
			$item['nasional'] = $array_tridharma_pendidikan[$row_id-1][3];
			$item['internasional'] = $array_tridharma_pendidikan[$row_id-1][4];
			$item['judul_kegiatan'] = $array_tridharma_pendidikan[$row_id-1][6];
			$item['manfaat'] = $array_tridharma_pendidikan[$row_id-1][7];
			$item['durasi'] = $array_tridharma_pendidikan[$row_id-1][8];
			$item['bukti'] = $array_tridharma_pendidikan[$row_id-1][9];
			$item['tahun_berakhir'] = $array_tridharma_pendidikan[$row_id-1][10];
            $data_tridharma_pendidikan[] = $item;
        }
    }
}

$spreadsheet_tridharma_pendidikan2->disconnectWorksheets();
unset($spreadsheet_tridharma_pendidikan2);


/**
 * Tabel 1.2 Kerjasama Tridharma Penelitian
 */

$spreadsheet_tridharma_penelitian = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_tridharma_penelitian.xls');

$worksheet_tridharma_penelitian = $spreadsheet_tridharma_penelitian->getActiveSheet();

$worksheet_tridharma_penelitian->insertNewColumnBefore('C', 3);

$highestRow_tridharma_penelitian = $worksheet_tridharma_penelitian->getHighestRow();

for($row = 2;$row <= $highestRow_tridharma_penelitian; $row++) {
	$worksheet_tridharma_penelitian->setCellValue('C'.$row, '=IF(F'.$row.'="Internasional";"V";"")');
	$worksheet_tridharma_penelitian->getCell('C'.$row)->getStyle()->setQuotePrefix(true);
	
	$worksheet_tridharma_penelitian->setCellValue('D'.$row, '=IF(F'.$row.'="Nasional";"V";"")');
	$worksheet_tridharma_penelitian->getCell('D'.$row)->getStyle()->setQuotePrefix(true);
	
	$worksheet_tridharma_penelitian->setCellValue('E'.$row, '=IF(F'.$row.'="Lokal";"V";"")');
	$worksheet_tridharma_penelitian->getCell('E'.$row)->getStyle()->setQuotePrefix(true);
}

$worksheet_tridharma_penelitian->setAutoFilter('B1:M'.$highestRow_tridharma_penelitian);
$autoFilter_tridharma_penelitian = $worksheet_tridharma_penelitian->getAutoFilter();
$columnFilter_tridharma_penelitian = $autoFilter_tridharma_penelitian->getColumn('L');
$columnFilter_tridharma_penelitian->setFilterType(
    \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_FILTER
);
$columnFilter_tridharma_penelitian->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        $nama_prodi2
    );

$autoFilter_tridharma_penelitian->showHideRows();

$writer_tridharma_penelitian = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_tridharma_penelitian, 'Xls');
$writer_tridharma_penelitian->save('./formatted/sapto_tridharma_penelitian (F).xls');

$spreadsheet_tridharma_penelitian->disconnectWorksheets();
unset($spreadsheet_tridharma_penelitian);

// Load Format Baru
$spreadsheet_tridharma_penelitian2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_tridharma_penelitian (F).xls');
$worksheet_tridharma_penelitian2 = $spreadsheet_tridharma_penelitian2->getActiveSheet();

// May be change
$array_tridharma_penelitian = $worksheet_tridharma_penelitian2->toArray();
$data_tridharma_penelitian = [];

foreach($worksheet_tridharma_penelitian2->getRowIterator() as $row_id => $row) {
    if($worksheet_tridharma_penelitian2->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['lembaga_mitra'] = $array_tridharma_penelitian[$row_id-1][1];
            $item['lokal'] = $array_tridharma_penelitian[$row_id-1][2];
			$item['nasional'] = $array_tridharma_penelitian[$row_id-1][3];
			$item['internasional'] = $array_tridharma_penelitian[$row_id-1][4];
			$item['judul_kegiatan'] = $array_tridharma_penelitian[$row_id-1][6];
			$item['manfaat'] = $array_tridharma_penelitian[$row_id-1][7];
			$item['durasi'] = $array_tridharma_penelitian[$row_id-1][8];
			$item['bukti'] = $array_tridharma_penelitian[$row_id-1][9];
			$item['tahun_berakhir'] = $array_tridharma_penelitian[$row_id-1][10];
            $data_tridharma_penelitian[] = $item;
        }
    }
}

$spreadsheet_tridharma_penelitian2->disconnectWorksheets();
unset($spreadsheet_tridharma_penelitian2);


/**
 * Tabel 1.3 Kerjasama Tridharma PkM
 */

$spreadsheet_tridharma_pkm = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_tridharma_pengmas.xls');

$worksheet_tridharma_pkm = $spreadsheet_tridharma_pkm->getActiveSheet();

$worksheet_tridharma_pkm->insertNewColumnBefore('C', 3);

$highestRow_tridharma_pkm = $worksheet_tridharma_pkm->getHighestRow();

for($row = 2;$row <= $highestRow_tridharma_pkm; $row++) {
	$worksheet_tridharma_pkm->setCellValue('C'.$row, '=IF(F'.$row.'="Internasional";"V";"")');
	$worksheet_tridharma_pkm->getCell('C'.$row)->getStyle()->setQuotePrefix(true);
	
	$worksheet_tridharma_pkm->setCellValue('D'.$row, '=IF(F'.$row.'="Nasional";"V";"")');
	$worksheet_tridharma_pkm->getCell('D'.$row)->getStyle()->setQuotePrefix(true);
	
	$worksheet_tridharma_pkm->setCellValue('E'.$row, '=IF(F'.$row.'="Lokal";"V";"")');
	$worksheet_tridharma_pkm->getCell('E'.$row)->getStyle()->setQuotePrefix(true);
}

$worksheet_tridharma_pkm->setAutoFilter('B1:M'.$highestRow_tridharma_pkm);
$autoFilter_tridharma_pkm = $worksheet_tridharma_pkm->getAutoFilter();
$columnFilter_tridharma_pkm = $autoFilter_tridharma_pkm->getColumn('L');
$columnFilter_tridharma_pkm->setFilterType(
    \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_FILTER
);
$columnFilter_tridharma_pkm->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        $nama_prodi2
    );

$autoFilter_tridharma_pkm->showHideRows();

$writer_tridharma_pkm = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_tridharma_pkm, 'Xls');
$writer_tridharma_pkm->save('./formatted/sapto_tridharma_pkm (F).xls');

$spreadsheet_tridharma_pkm->disconnectWorksheets();
unset($spreadsheet_tridharma_pkm);

// Load Format Baru
$spreadsheet_tridharma_pkm2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_tridharma_pkm (F).xls');
$worksheet_tridharma_pkm2 = $spreadsheet_tridharma_pkm2->getActiveSheet();

// May be change
$array_tridharma_pkm = $worksheet_tridharma_pkm2->toArray();
$data_tridharma_pkm = [];

foreach($worksheet_tridharma_pkm2->getRowIterator() as $row_id => $row) {
    if($worksheet_tridharma_pkm2->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['lembaga_mitra'] = $array_tridharma_pkm[$row_id-1][1];
            $item['lokal'] = $array_tridharma_pkm[$row_id-1][2];
			$item['nasional'] = $array_tridharma_pkm[$row_id-1][3];
			$item['internasional'] = $array_tridharma_pkm[$row_id-1][4];
			$item['judul_kegiatan'] = $array_tridharma_pkm[$row_id-1][6];
			$item['manfaat'] = $array_tridharma_pkm[$row_id-1][7];
			$item['durasi'] = $array_tridharma_pkm[$row_id-1][8];
			$item['bukti'] = $array_tridharma_pkm[$row_id-1][9];
			$item['tahun_berakhir'] = $array_tridharma_pkm[$row_id-1][10];
            $data_tridharma_pkm[] = $item;
        }
    }
}

$spreadsheet_tridharma_pkm2->disconnectWorksheets();
unset($spreadsheet_tridharma_pkm2);


/**
 * Tabel 2.a Seleksi Mahasiswa Baru
 */

$spreadsheet_seleksi_mhs = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_seleksi_mhs_baru.xlsx');

$worksheet_seleksi_mhs = $spreadsheet_seleksi_mhs->getActiveSheet();

$highestRow_seleksi_mhs = $worksheet_seleksi_mhs->getHighestRow();

$worksheet_seleksi_mhs->setAutoFilter('B1:L'.$highestRow_seleksi_mhs);
$autoFilter_seleksi_mhs = $worksheet_seleksi_mhs->getAutoFilter();
$columnFilter_seleksi_mhs = $autoFilter_seleksi_mhs->getColumn('B');
$columnFilter_seleksi_mhs2 = $autoFilter_seleksi_mhs->getColumn('E');
$columnFilter_seleksi_mhs->setFilterType(
    \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_FILTER
);
$columnFilter_seleksi_mhs->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        $nama_prodi
    );
$columnFilter_seleksi_mhs2->setFilterType(
    \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_FILTER
);
$columnFilter_seleksi_mhs2->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        'TS-4'
    );
$columnFilter_seleksi_mhs2->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        'TS-3'
    );
$columnFilter_seleksi_mhs2->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        'TS-2'
    );
$columnFilter_seleksi_mhs2->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        'TS-1'
    );
$columnFilter_seleksi_mhs2->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        'TS'
    );

$autoFilter_seleksi_mhs->showHideRows();

// May be change
$array_seleksi_mhs = $worksheet_seleksi_mhs->toArray();
$data_seleksi_mhs = [];

foreach($worksheet_seleksi_mhs->getRowIterator() as $row_id => $row) {
    if($worksheet_seleksi_mhs->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['tahun_akademik'] = $array_seleksi_mhs[$row_id-1][4];
            $item['daya_tampung'] = $array_seleksi_mhs[$row_id-1][5];
			$item['pendaftar'] = $array_seleksi_mhs[$row_id-1][6];
			$item['lulus_seleksi'] = $array_seleksi_mhs[$row_id-1][7];
			$item['maba_reguler'] = $array_seleksi_mhs[$row_id-1][8];
			$item['maba_transfer'] = $array_seleksi_mhs[$row_id-1][9];
			$item['mhs_reguler'] = $array_seleksi_mhs[$row_id-1][10];
			$item['mhs_transfer'] = $array_seleksi_mhs[$row_id-1][11];
            $data_seleksi_mhs[] = $item;
        }
    }
}

$worksheet_seleksi_mhs2 = new \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet($spreadsheet_seleksi_mhs, 'Sheet 2');
$spreadsheet_seleksi_mhs->addSheet($worksheet_seleksi_mhs2);

$worksheet_seleksi_mhs2 = $spreadsheet_seleksi_mhs->getSheetByName('Sheet 2');
$worksheet_seleksi_mhs2->fromArray($data_seleksi_mhs, NULL, 'A1');

$writer_seleksi_mhs = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_seleksi_mhs, 'Xlsx');
$writer_seleksi_mhs->save('./formatted/sapto_seleksi_mhs_baru (F).xlsx');

$spreadsheet_seleksi_mhs->disconnectWorksheets();
unset($spreadsheet_seleksi_mhs);

// Load Format Baru
$spreadsheet_seleksi_mhs2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_seleksi_mhs_baru (F).xlsx');
$worksheet_seleksi_mhs3 = $spreadsheet_seleksi_mhs2->getSheetByName('Sheet 2');

// May be change
$ts4_seleksi_mhs = $worksheet_seleksi_mhs3->rangeToArray('B5:H5', NULL, TRUE, TRUE, TRUE);
$ts3_seleksi_mhs = $worksheet_seleksi_mhs3->rangeToArray('B4:H4', NULL, TRUE, TRUE, TRUE);
$ts2_seleksi_mhs = $worksheet_seleksi_mhs3->rangeToArray('B3:H3', NULL, TRUE, TRUE, TRUE);
$ts1_seleksi_mhs = $worksheet_seleksi_mhs3->rangeToArray('B2:H2', NULL, TRUE, TRUE, TRUE);
$ts_seleksi_mhs = $worksheet_seleksi_mhs3->rangeToArray('B1:H1', NULL, TRUE, TRUE, TRUE);

$spreadsheet_seleksi_mhs2->disconnectWorksheets();
unset($spreadsheet_seleksi_mhs2);


/**
 * Tabel 2.b Mahasiswa Asing
 */

$spreadsheet_mhs_asing = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_mhs_asing.xlsx');

$worksheet_mhs_asing = $spreadsheet_mhs_asing->getActiveSheet();

$worksheet_mhs_asing->insertNewColumnBefore('G', 3);
$worksheet_mhs_asing->insertNewColumnBefore('K', 3);

$worksheet_mhs_asing->getCell('G1')->setValue('TS-2 Aktif');
$worksheet_mhs_asing->getCell('H1')->setValue('TS-1 Aktif');
$worksheet_mhs_asing->getCell('I1')->setValue('TS Aktif');
$worksheet_mhs_asing->getCell('K1')->setValue('TS-2 Full Time');
$worksheet_mhs_asing->getCell('L1')->setValue('TS-1 Full Time');
$worksheet_mhs_asing->getCell('M1')->setValue('TS Full Time');
$worksheet_mhs_asing->getCell('O1')->setValue('TS-2 Part Time');
$worksheet_mhs_asing->getCell('P1')->setValue('TS-1 Part Time');
$worksheet_mhs_asing->getCell('Q1')->setValue('TS Part Time');

$highestRow_mhs_asing = $worksheet_mhs_asing->getHighestRow();

for($row = 2;$row <= $highestRow_mhs_asing; $row++) {
	$worksheet_mhs_asing->setCellValue('G'.$row, '=IF(D'.$row.'=2018,F'.$row.',0)');
	$worksheet_mhs_asing->getCell('G'.$row)->getStyle()->setQuotePrefix(true);
	
	$worksheet_mhs_asing->setCellValue('H'.$row, '=IF(D'.$row.'=2019,F'.$row.',0)');
	$worksheet_mhs_asing->getCell('H'.$row)->getStyle()->setQuotePrefix(true);
	
	$worksheet_mhs_asing->setCellValue('I'.$row, '=IF(D'.$row.'=2020,F'.$row.',0)');
	$worksheet_mhs_asing->getCell('I'.$row)->getStyle()->setQuotePrefix(true);
	
	$worksheet_mhs_asing->setCellValue('K'.$row, '=IF(D'.$row.'=2018,J'.$row.',0)');
	$worksheet_mhs_asing->getCell('K'.$row)->getStyle()->setQuotePrefix(true);
	
	$worksheet_mhs_asing->setCellValue('L'.$row, '=IF(D'.$row.'=2019,J'.$row.',0)');
	$worksheet_mhs_asing->getCell('L'.$row)->getStyle()->setQuotePrefix(true);
	
	$worksheet_mhs_asing->setCellValue('M'.$row, '=IF(D'.$row.'=2020,J'.$row.',0)');
	$worksheet_mhs_asing->getCell('M'.$row)->getStyle()->setQuotePrefix(true);
	
	$worksheet_mhs_asing->setCellValue('O'.$row, '=IF(D'.$row.'=2018,N'.$row.',0)');
	$worksheet_mhs_asing->getCell('O'.$row)->getStyle()->setQuotePrefix(true);
	
	$worksheet_mhs_asing->setCellValue('P'.$row, '=IF(D'.$row.'=2019,N'.$row.',0)');
	$worksheet_mhs_asing->getCell('P'.$row)->getStyle()->setQuotePrefix(true);
	
	$worksheet_mhs_asing->setCellValue('Q'.$row, '=IF(D'.$row.'=2020,N'.$row.',0)');
	$worksheet_mhs_asing->getCell('Q'.$row)->getStyle()->setQuotePrefix(true);
}

$worksheet_mhs_asing->setAutoFilter('A1:Q'.$highestRow_mhs_asing);
$autoFilter_mhs_asing = $worksheet_mhs_asing->getAutoFilter();
$columnFilter_mhs_asing = $autoFilter_mhs_asing->getColumn('B');
$columnFilter_mhs_asing2 = $autoFilter_mhs_asing->getColumn('E');
$columnFilter_mhs_asing->setFilterType(
    \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_FILTER
);
$columnFilter_mhs_asing->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
       $nama_prodi
    );
$columnFilter_mhs_asing2->setFilterType(
    \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_FILTER
);
$columnFilter_mhs_asing2->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        'TS-2'
    );
$columnFilter_mhs_asing2->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        'TS-1'
    );
$columnFilter_mhs_asing2->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        'TS'
    );

$autoFilter_mhs_asing->showHideRows();

$writer_mhs_asing = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_mhs_asing, 'Xls');
$writer_mhs_asing->save('./formatted/sapto_mhs_asing (F).xls');

$spreadsheet_mhs_asing->disconnectWorksheets();
unset($spreadsheet_mhs_asing);

// Load Format Baru
$spreadsheet_mhs_asing2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_mhs_asing (F).xls');
$worksheet_mhs_asing2 = $spreadsheet_mhs_asing2->getActiveSheet();

// May be change
$array_mhs_asing = $worksheet_mhs_asing2->toArray();
$data_mhs_asing = [];

foreach($worksheet_mhs_asing2->getRowIterator() as $row_id => $row) {
    if($worksheet_mhs_asing2->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['nama_prodi'] = $array_mhs_asing[$row_id-1][1];
            $item['ts2_aktif'] = $array_mhs_asing[$row_id-1][6];
			$item['ts1_aktif'] = $array_mhs_asing[$row_id-1][7];
			$item['ts_aktif'] = $array_mhs_asing[$row_id-1][8];
			$item['ts2_full_time'] = $array_mhs_asing[$row_id-1][10];
			$item['ts1_full_time'] = $array_mhs_asing[$row_id-1][11];
			$item['ts_full_time'] = $array_mhs_asing[$row_id-1][12];
			$item['ts2_part_time'] = $array_mhs_asing[$row_id-1][14];
			$item['ts1_part_time'] = $array_mhs_asing[$row_id-1][15];
			$item['ts_part_time'] = $array_mhs_asing[$row_id-1][16];
            $data_mhs_asing[] = $item;
        }
    }
}

$worksheet_mhs_asing3 = new \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet($spreadsheet_mhs_asing2, 'Sheet 2');
$spreadsheet_mhs_asing2->addSheet($worksheet_mhs_asing3);

$worksheet_mhs_asing3 = $spreadsheet_mhs_asing2->getSheetByName('Sheet 2');
$worksheet_mhs_asing3->fromArray($data_mhs_asing, NULL, 'A1');

$highestRow_mhs_asing3 = $worksheet_mhs_asing3->getHighestRow();

$row_jumlah_mhs_asing = -1;

for($group = 1;$group <= 3; $group++) {
	
	$row_jumlah_mhs_asing += 2;
	$baris_awal_mhs_asing = $row_jumlah_mhs_asing;
	
	$ts2_aktif = $worksheet_mhs_asing3->getCell('B'.($row_jumlah_mhs_asing))->getValue(); 
	$ts1_aktif = $worksheet_mhs_asing3->getCell('C'.($row_jumlah_mhs_asing))->getValue();
	$ts_aktif = $worksheet_mhs_asing3->getCell('D'.($row_jumlah_mhs_asing))->getValue();
	$ts2_full_time = $worksheet_mhs_asing3->getCell('E'.($row_jumlah_mhs_asing))->getValue();
	$ts1_full_time = $worksheet_mhs_asing3->getCell('F'.($row_jumlah_mhs_asing))->getValue();
	$ts_full_time = $worksheet_mhs_asing3->getCell('G'.($row_jumlah_mhs_asing))->getValue();
	$ts2_part_time = $worksheet_mhs_asing3->getCell('H'.($row_jumlah_mhs_asing))->getValue();
	$ts1_part_time = $worksheet_mhs_asing3->getCell('I'.($row_jumlah_mhs_asing))->getValue();
	$ts_part_time = $worksheet_mhs_asing3->getCell('J'.($row_jumlah_mhs_asing))->getValue();
	
	for($row = $row_jumlah_mhs_asing;$row <= ($highestRow_mhs_asing3+1); $row++) {
		if($worksheet_mhs_asing3->getCell('A'.$row)->getValue() == $worksheet_mhs_asing3->getCell('A'.($row+1))->getValue()) {
			$ts2_aktif += $worksheet_mhs_asing3->getCell('B'.($row+1))->getValue();
			$ts1_aktif += $worksheet_mhs_asing3->getCell('C'.($row+1))->getValue();
			$ts_aktif += $worksheet_mhs_asing3->getCell('D'.($row+1))->getValue();
			$ts2_full_time += $worksheet_mhs_asing3->getCell('E'.($row+1))->getValue();
			$ts1_full_time += $worksheet_mhs_asing3->getCell('F'.($row+1))->getValue();
			$ts_full_time += $worksheet_mhs_asing3->getCell('G'.($row+1))->getValue();
			$ts2_part_time += $worksheet_mhs_asing3->getCell('H'.($row+1))->getValue();
			$ts1_part_time += $worksheet_mhs_asing3->getCell('I'.($row+1))->getValue();
			$ts_part_time += $worksheet_mhs_asing3->getCell('J'.($row+1))->getValue();
			
			$row_jumlah_mhs_asing++;
		} else {
			break;
		}
	}
	
	$worksheet_mhs_asing3->insertNewRowBefore(($row_jumlah_mhs_asing+1), 1);
	$worksheet_mhs_asing3->setCellValue('A'.($row_jumlah_mhs_asing+1), $worksheet_mhs_asing3->getCell('A'.$row_jumlah_mhs_asing)->getValue());
	$worksheet_mhs_asing3->setCellValue('B'.($row_jumlah_mhs_asing+1), $ts2_aktif);
	$worksheet_mhs_asing3->setCellValue('C'.($row_jumlah_mhs_asing+1), $ts1_aktif);
	$worksheet_mhs_asing3->setCellValue('D'.($row_jumlah_mhs_asing+1), $ts_aktif);
	$worksheet_mhs_asing3->setCellValue('E'.($row_jumlah_mhs_asing+1), $ts2_full_time);	
	$worksheet_mhs_asing3->setCellValue('F'.($row_jumlah_mhs_asing+1), $ts1_full_time);
	$worksheet_mhs_asing3->setCellValue('G'.($row_jumlah_mhs_asing+1), $ts_full_time);
	$worksheet_mhs_asing3->setCellValue('H'.($row_jumlah_mhs_asing+1), $ts2_part_time);	
	$worksheet_mhs_asing3->setCellValue('I'.($row_jumlah_mhs_asing+1), $ts1_part_time);
	$worksheet_mhs_asing3->setCellValue('J'.($row_jumlah_mhs_asing+1), $ts_part_time);
	
	${"mhs_asing".$group} = $worksheet_mhs_asing3->rangeToArray('A'.($row_jumlah_mhs_asing+1).':J'.($row_jumlah_mhs_asing+1), NULL, TRUE, TRUE, TRUE);
}

$writer_mhs_asing2 = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_mhs_asing2, 'Xls');
$writer_mhs_asing2->save('./formatted/sapto_mhs_asing (F).xls');

$spreadsheet_mhs_asing2->disconnectWorksheets();
unset($spreadsheet_mhs_asing2);


/**
 * Tabel 3.a.1 Dosen Tetap Perguruan Tinggi
 */

$spreadsheet_dosen_tetap = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_dosen_tetap_matkul.xlsx');

$worksheet_dosen_tetap = $spreadsheet_dosen_tetap->getActiveSheet();

$worksheet_dosen_tetap->insertNewColumnBefore('J', 2);
$worksheet_dosen_tetap->insertNewColumnBefore('O', 2);

$highestRow_dosen_tetap = $worksheet_dosen_tetap->getHighestRow();

for($row = 2;$row <= $highestRow_dosen_tetap; $row++) {
	$worksheet_dosen_tetap->setCellValue('J'.$row, '=IF(I'.$row.'=1;"V";"")');
	$worksheet_dosen_tetap->getCell('J'.$row)->getStyle()->setQuotePrefix(true);
	
	$worksheet_dosen_tetap->setCellValue('O'.$row, '=IF(N'.$row.'=1;"V";"")');
	$worksheet_dosen_tetap->getCell('O'.$row)->getStyle()->setQuotePrefix(true);
}

$worksheet_dosen_tetap->setAutoFilter('B1:Y'.$highestRow_dosen_tetap);
$autoFilter_dosen_tetap = $worksheet_dosen_tetap->getAutoFilter();
$columnFilter_dosen_tetap = $autoFilter_dosen_tetap->getColumn('B');
$columnFilter_dosen_tetap->setFilterType(
    \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_FILTER
);
$columnFilter_dosen_tetap->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        15
    );

$autoFilter_dosen_tetap->showHideRows();

$writer_dosen_tetap = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_dosen_tetap, 'Xls');
$writer_dosen_tetap->save('./formatted/sapto_dosen_tetap_matkul (F).xls');

$spreadsheet_dosen_tetap->disconnectWorksheets();
unset($spreadsheet_dosen_tetap);

// Load Format Baru
$spreadsheet_dosen_tetap2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_dosen_tetap_matkul (F).xls');
$worksheet_dosen_tetap2 = $spreadsheet_dosen_tetap2->getActiveSheet();

// May be change
$array_dosen_tetap = $worksheet_dosen_tetap2->toArray();
$data_dosen_tetap = [];

foreach($worksheet_dosen_tetap2->getRowIterator() as $row_id => $row) {
    if($worksheet_dosen_tetap2->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['nama_dosen'] = $array_dosen_tetap[$row_id-1][2];
            $item['nidn_nidk'] = $array_dosen_tetap[$row_id-1][4];
			$item['pendidikan_s2'] = $array_dosen_tetap[$row_id-1][5];
			$item['pendidikan_s3'] = $array_dosen_tetap[$row_id-1][6];
			$item['bidang_keahlian'] = $array_dosen_tetap[$row_id-1][7];
			$item['sesuai_kompetensi_inti'] = $array_dosen_tetap[$row_id-1][14];
			$item['jabatan_akademik'] = $array_dosen_tetap[$row_id-1][11];
			$item['sertifikat_pendidik'] = $array_dosen_tetap[$row_id-1][12];
			$item['sertifikat_kompetensi'] = $array_dosen_tetap[$row_id-1][10];
			$item['matkul_ps'] = $array_dosen_tetap[$row_id-1][16];
			$item['sesuai_bidang_keahlian'] = $array_dosen_tetap[$row_id-1][9];
			$item['matkul_ps_lain'] = $array_dosen_tetap[$row_id-1][15];
            $data_dosen_tetap[] = $item;
        }
    }
}

$spreadsheet_dosen_tetap2->disconnectWorksheets();
unset($spreadsheet_dosen_tetap2);


/**
 * Tabel 3.a.2 Dosen Pembimbing Utama TA
 */

$spreadsheet_dosen_pembimbing = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_dosen_pembimbing_nilai_rata.xlsx');

$worksheet_dosen_pembimbing = $spreadsheet_dosen_pembimbing->getActiveSheet();

$worksheet_dosen_pembimbing->insertNewColumnBefore('F', 8);

$worksheet_dosen_pembimbing->getCell('F1')->setValue('TS-2 Prodi');
$worksheet_dosen_pembimbing->getCell('G1')->setValue('TS-1 Prodi');
$worksheet_dosen_pembimbing->getCell('H1')->setValue('TS Prodi');
$worksheet_dosen_pembimbing->getCell('I1')->setValue('Rata Prodi');
$worksheet_dosen_pembimbing->getCell('J1')->setValue('TS-2 Prodi Lain');
$worksheet_dosen_pembimbing->getCell('K1')->setValue('TS-1 Prodi Lain');
$worksheet_dosen_pembimbing->getCell('L1')->setValue('TS Prodi Lain');
$worksheet_dosen_pembimbing->getCell('M1')->setValue('Rata Prodi Lain');

$highestRow_dosen_pembimbing = $worksheet_dosen_pembimbing->getHighestRow();

for($row = 2;$row <= $highestRow_dosen_pembimbing; $row++) {
	if($nama_prodi == $worksheet_dosen_pembimbing->getCell('E'.$row)->getValue()) {
		$worksheet_dosen_pembimbing->setCellValue('F'.$row, '=IF(D'.$row.'=2018,P'.$row.',0)');
		$worksheet_dosen_pembimbing->getCell('F'.$row)->getStyle()->setQuotePrefix(true);
		
		$worksheet_dosen_pembimbing->setCellValue('G'.$row, '=IF(D'.$row.'=2019,P'.$row.',0)');
		$worksheet_dosen_pembimbing->getCell('G'.$row)->getStyle()->setQuotePrefix(true);
		
		$worksheet_dosen_pembimbing->setCellValue('H'.$row, '=IF(D'.$row.'=2020,P'.$row.',0)');
		$worksheet_dosen_pembimbing->getCell('H'.$row)->getStyle()->setQuotePrefix(true);
		
		$worksheet_dosen_pembimbing->setCellValue('I'.$row, $worksheet_dosen_pembimbing->getCell('Q'.$row)->getValue());
	} else {
		$worksheet_dosen_pembimbing->setCellValue('J'.$row, '=IF(D'.$row.'=2018,P'.$row.',0)');
		$worksheet_dosen_pembimbing->getCell('J'.$row)->getStyle()->setQuotePrefix(true);
		
		$worksheet_dosen_pembimbing->setCellValue('K'.$row, '=IF(D'.$row.'=2019,P'.$row.',0)');
		$worksheet_dosen_pembimbing->getCell('K'.$row)->getStyle()->setQuotePrefix(true);
		
		$worksheet_dosen_pembimbing->setCellValue('L'.$row, '=IF(D'.$row.'=2020,P'.$row.',0)');
		$worksheet_dosen_pembimbing->getCell('L'.$row)->getStyle()->setQuotePrefix(true);
	
		$worksheet_dosen_pembimbing->setCellValue('M'.$row, $worksheet_dosen_pembimbing->getCell('Q'.$row)->getValue());
	}
}

$worksheet_dosen_pembimbing->setAutoFilter('A1:R'.$highestRow_dosen_pembimbing);
$autoFilter_dosen_pembimbing = $worksheet_dosen_pembimbing->getAutoFilter();
$columnFilter_dosen_pembimbing = $autoFilter_dosen_pembimbing->getColumn('C');
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

// May be change
$array_dosen_pembimbing = $worksheet_dosen_pembimbing2->toArray();
$data_dosen_pembimbing = [];

foreach($worksheet_dosen_pembimbing2->getRowIterator() as $row_id => $row) {
    if($worksheet_dosen_pembimbing2->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['nama_dosen'] = $array_dosen_pembimbing[$row_id-1][0];
            $item['ts2_prodi'] = $array_dosen_pembimbing[$row_id-1][5];
			$item['ts1_prodi'] = $array_dosen_pembimbing[$row_id-1][6];
			$item['ts_prodi'] = $array_dosen_pembimbing[$row_id-1][7];
			$item['rata_prodi'] = $array_dosen_pembimbing[$row_id-1][8];
			$item['ts2_prodi_lain'] = $array_dosen_pembimbing[$row_id-1][9];
			$item['ts1_prodi_lain'] = $array_dosen_pembimbing[$row_id-1][10];
			$item['ts_prodi_lain'] = $array_dosen_pembimbing[$row_id-1][11];
			$item['rata_prodi_lain'] = $array_dosen_pembimbing[$row_id-1][12];
			$item['rata_total'] = $array_dosen_pembimbing[$row_id-1][17];
            $data_dosen_pembimbing[] = $item;
        }
    }
}

$worksheet_dosen_pembimbing3 = new \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet($spreadsheet_dosen_pembimbing2, 'Sheet 2');
$spreadsheet_dosen_pembimbing2->addSheet($worksheet_dosen_pembimbing3);

$worksheet_dosen_pembimbing3 = $spreadsheet_dosen_pembimbing2->getSheetByName('Sheet 2');
$worksheet_dosen_pembimbing3->fromArray($data_dosen_pembimbing, NULL, 'A1');

$highestRow_dosen_pembimbing3 = $worksheet_dosen_pembimbing3->getHighestRow();

$worksheet_dosen_pembimbing3->setCellValue('L1', '=SUMPRODUCT((A1:A'.$highestRow_dosen_pembimbing3.'<>"")/COUNTIF(A1:A'.$highestRow_dosen_pembimbing3.',A1:A'.$highestRow_dosen_pembimbing3.'&""))');

$total_dosen_pembimbing = $worksheet_dosen_pembimbing3->getCell('L1')->getValue();
$row_jumlah_pembimbing = -1;

for($group = 1;$group <= 88; $group++) {
	
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

$spreadsheet_ewmp = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_ewmp_dosen_tetap.xlsx');

$worksheet_ewmp = $spreadsheet_ewmp->getActiveSheet();

$worksheet_ewmp->insertNewColumnBefore('D', 1);

$highestRow_ewmp = $worksheet_ewmp->getHighestRow();

for($row = 2;$row <= $highestRow_ewmp; $row++) {
	$worksheet_ewmp->setCellValue('D'.$row, '=IF(C'.$row.'=1;"V";"")');
	$worksheet_ewmp->getCell('D'.$row)->getStyle()->setQuotePrefix(true);
}

$worksheet_ewmp->setAutoFilter('B1:N'.$highestRow_ewmp);
$autoFilter_ewmp = $worksheet_ewmp->getAutoFilter();
$columnFilter_ewmp = $autoFilter_ewmp->getColumn('M');
$columnFilter_ewmp->setFilterType(
    \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_FILTER
);
$columnFilter_ewmp->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        $nama_prodi2
    );

$autoFilter_ewmp->showHideRows();

$writer_ewmp = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_ewmp, 'Xls');
$writer_ewmp->save('./formatted/sapto_ewmp_dosen_tetap (F).xls');

$spreadsheet_ewmp->disconnectWorksheets();
unset($spreadsheet_ewmp);

// Load Format Baru
$spreadsheet_ewmp2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_ewmp_dosen_tetap (F).xls');
$worksheet_ewmp2 = $spreadsheet_ewmp2->getActiveSheet();

// May be change
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

$spreadsheet_dosen_tidaktetap = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_dosen_tidaktetap_matkul.xlsx');

$worksheet_dosen_tidaktetap = $spreadsheet_dosen_tidaktetap->getActiveSheet();

$worksheet_dosen_tidaktetap->insertNewColumnBefore('G', 1);
$worksheet_dosen_tidaktetap->insertNewColumnBefore('L', 1);
$worksheet_dosen_tidaktetap->insertNewColumnBefore('O', 1);

$worksheet_dosen_tidaktetap->getCell('G1')->setValue('Pendidikan Pasca Sarjana');
$worksheet_dosen_tidaktetap->getCell('L1')->setValue('Sertifikat Kompetensi/Profesi/Industri');
$worksheet_dosen_tidaktetap->getCell('O1')->setValue('Kesesuaian Bidang Keahlian dengan Mata Kuliah yang Diampu');

$highestRow_dosen_tidaktetap = $worksheet_dosen_tidaktetap->getHighestRow();

for($row = 2;$row <= $highestRow_dosen_tidaktetap; $row++) {
	$worksheet_dosen_tidaktetap->setCellValue('G'.$row, '=IF(F'.$row.'<>"";F'.$row.';E'.$row.')');
	$worksheet_dosen_tidaktetap->getCell('G'.$row)->getStyle()->setQuotePrefix(true);
	
	$worksheet_dosen_tidaktetap->setCellValue('O'.$row, '=IF(I'.$row.'=1;"V";"")');
	$worksheet_dosen_tidaktetap->getCell('O'.$row)->getStyle()->setQuotePrefix(true);
}

$worksheet_dosen_tidaktetap->setAutoFilter('A1:X'.$highestRow_dosen_tidaktetap);
$autoFilter_dosen_tidaktetap = $worksheet_dosen_tidaktetap->getAutoFilter();
$columnFilter_dosen_tidaktetap = $autoFilter_dosen_tidaktetap->getColumn('A');
$columnFilter_dosen_tidaktetap->setFilterType(
    \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_FILTER
);
$columnFilter_dosen_tidaktetap->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        $nama_prodi
    );

$autoFilter_dosen_tidaktetap->showHideRows();

$writer_dosen_tidaktetap = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_dosen_tidaktetap, 'Xls');
$writer_dosen_tidaktetap->save('./formatted/sapto_dosen_tidaktetap_matkul (F).xls');

$spreadsheet_dosen_tidaktetap->disconnectWorksheets();
unset($spreadsheet_dosen_tidaktetap);

// Load Format Baru
$spreadsheet_dosen_tidaktetap2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_dosen_tidaktetap_matkul (F).xls');
$worksheet_dosen_tidaktetap2 = $spreadsheet_dosen_tidaktetap2->getActiveSheet();

// May be change
$array_dosen_tidaktetap = $worksheet_dosen_tidaktetap2->toArray();
$data_dosen_tidaktetap = [];

foreach($worksheet_dosen_tidaktetap2->getRowIterator() as $row_id => $row) {
    if($worksheet_dosen_tidaktetap2->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['nama_dosen'] = $array_dosen_tidaktetap[$row_id-1][1];
            $item['nidn_nidk'] = $array_dosen_tidaktetap[$row_id-1][3];
			$item['pendidikan_pasca'] = $array_dosen_tidaktetap[$row_id-1][6];
			$item['bidang_keahlian'] = $array_dosen_tidaktetap[$row_id-1][7];
			$item['jabatan_akademik'] = $array_dosen_tidaktetap[$row_id-1][9];
			$item['sertifikat_pendidik'] = $array_dosen_tidaktetap[$row_id-1][10];
			$item['sertifikat_kompetensi'] = $array_dosen_tidaktetap[$row_id-1][11];
			$item['matkul_ps'] = $array_dosen_tidaktetap[$row_id-1][13];
			$item['sesuai_bidang_keahlian'] = $array_dosen_tidaktetap[$row_id-1][14];
            $data_dosen_tidaktetap[] = $item;
        }
    }
}

$spreadsheet_dosen_tidaktetap2->disconnectWorksheets();
unset($spreadsheet_dosen_tidaktetap2);


/**
 * Tabel 3.a.5 Dosen Industri/Praktisi
 */

$spreadsheet_dosen_industri = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_dosen_industri.xlsx');

$worksheet_dosen_industri = $spreadsheet_dosen_industri->getActiveSheet();

$worksheet_dosen_industri->insertNewColumnBefore('E', 2);
$worksheet_dosen_industri->insertNewColumnBefore('J', 3);

$worksheet_dosen_industri->getCell('E1')->setValue('Perusahaan/Industri');
$worksheet_dosen_industri->getCell('F1')->setValue('Pendidikan Tertinggi');
$worksheet_dosen_industri->getCell('J1')->setValue('Sertifikat Profesi/ Kompetensi/ Industri');
$worksheet_dosen_industri->getCell('K1')->setValue('Mata Kuliah yang Diampu');
$worksheet_dosen_industri->getCell('L1')->setValue('Bobot Kredit (sks)');

$highestRow_dosen_industri = $worksheet_dosen_industri->getHighestRow();

$worksheet_dosen_industri->setAutoFilter('A1:O'.$highestRow_dosen_industri);
$autoFilter_dosen_industri = $worksheet_dosen_industri->getAutoFilter();
$columnFilter_dosen_industri = $autoFilter_dosen_industri->getColumn('M');
$columnFilter_dosen_industri->setFilterType(
    \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_FILTER
);
$columnFilter_dosen_industri->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        1
    );

$autoFilter_dosen_industri->showHideRows();

$writer_dosen_industri = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_dosen_industri, 'Xlsx');
$writer_dosen_industri->save('./formatted/sapto_dosen_industri (F).xlsx');

$spreadsheet_dosen_industri->disconnectWorksheets();
unset($spreadsheet_dosen_industri);

// Load Format Baru
$spreadsheet_dosen_industri2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_dosen_industri (F).xlsx');
$worksheet_dosen_industri2 = $spreadsheet_dosen_industri2->getActiveSheet();

// May be change
$array_dosen_industri = $worksheet_dosen_industri2->toArray();
$data_dosen_industri = [];

foreach($worksheet_dosen_industri2->getRowIterator() as $row_id => $row) {
    if($worksheet_dosen_industri2->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['nama_dosen'] = $array_dosen_industri[$row_id-1][2];
            $item['nidn_nidk'] = $array_dosen_industri[$row_id-1][3];
			$item['perusahaan'] = $array_dosen_industri[$row_id-1][4];
			$item['pendidikan_tertinggi'] = $array_dosen_industri[$row_id-1][5];
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
 * Tabel 3.b.1 Pengakuan/Rekognisi Dosen
 */

$spreadsheet_rekognisi = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_rekognisi_dosen.xls');

$worksheet_rekognisi = $spreadsheet_rekognisi->getActiveSheet();

$worksheet_rekognisi->removeColumn('D');

$worksheet_rekognisi->insertNewColumnBefore('F', 4);

$worksheet_rekognisi->getCell('F1')->setValue('Wilayah');
$worksheet_rekognisi->getCell('G1')->setValue('Nasional');
$worksheet_rekognisi->getCell('H1')->setValue('Internasional');

$highestRow_rekognisi = $worksheet_rekognisi->getHighestRow();

$array_tahun_rekognisi = $worksheet_rekognisi->rangeToArray('L1:L'.$highestRow_rekognisi, NULL, TRUE, TRUE, TRUE);
$worksheet_rekognisi->fromArray($array_tahun_rekognisi, NULL, 'I1');

for($row = 2;$row <= $highestRow_rekognisi; $row++) {
	$worksheet_rekognisi->setCellValue('F'.$row, '=IF(K'.$row.'="ITS";"V";"")');
	$worksheet_rekognisi->getCell('F'.$row)->getStyle()->setQuotePrefix(true);
	
	$worksheet_rekognisi->setCellValue('G'.$row, '=IF(K'.$row.'="Nasional";"V";"")');
	$worksheet_rekognisi->getCell('G'.$row)->getStyle()->setQuotePrefix(true);
	
	$worksheet_rekognisi->setCellValue('H'.$row, '=IF(K'.$row.'="International";"V";"")');
	$worksheet_rekognisi->getCell('H'.$row)->getStyle()->setQuotePrefix(true);
}

$worksheet_rekognisi->setAutoFilter('B1:L'.$highestRow_rekognisi);
$autoFilter_rekognisi = $worksheet_rekognisi->getAutoFilter();
$columnFilter_rekognisi = $autoFilter_rekognisi->getColumn('B');
$columnFilter_rekognisi->setFilterType(
    \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_FILTER
);
$columnFilter_rekognisi->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        $nama_prodi2
    );

$autoFilter_rekognisi->showHideRows();

$writer_rekognisi = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_rekognisi, 'Xls');
$writer_rekognisi->save('./formatted/sapto_rekognisi_dosen (F).xls');

$spreadsheet_rekognisi->disconnectWorksheets();
unset($spreadsheet_rekognisi);

// Load Format Baru
$spreadsheet_rekognisi2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_rekognisi_dosen (F).xls');
$worksheet_rekognisi2 = $spreadsheet_rekognisi2->getActiveSheet();

// May be change
$array_rekognisi = $worksheet_rekognisi2->toArray();
$data_rekognisi = [];

foreach($worksheet_rekognisi2->getRowIterator() as $row_id => $row) {
    if($worksheet_rekognisi2->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['nama_dosen'] = $array_rekognisi[$row_id-1][2];
            $item['bidang_keahlian'] = $array_rekognisi[$row_id-1][3];
			$item['rekognisi'] = $array_rekognisi[$row_id-1][4];
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

$spreadsheet_penelitian_dtps = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_penelitian_dtps.xls');

$worksheet_penelitian_dtps = $spreadsheet_penelitian_dtps->getActiveSheet();

$highestRow_penelitian_dtps = $worksheet_penelitian_dtps->getHighestRow();

$worksheet_penelitian_dtps->setAutoFilter('B1:H'.$highestRow_penelitian_dtps);
$autoFilter_penelitian_dtps = $worksheet_penelitian_dtps->getAutoFilter();
$columnFilter_penelitian_dtps = $autoFilter_penelitian_dtps->getColumn('G');
$columnFilter_penelitian_dtps->setFilterType(
    \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_FILTER
);
$columnFilter_penelitian_dtps->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        $nama_prodi2
    );

$autoFilter_penelitian_dtps->showHideRows();

$writer_penelitian_dtps = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_penelitian_dtps, 'Xls');
$writer_penelitian_dtps->save('./formatted/sapto_penelitian_dtps (F).xls');

$spreadsheet_penelitian_dtps->disconnectWorksheets();
unset($spreadsheet_penelitian_dtps);

// Load Format Baru
$spreadsheet_penelitian_dtps2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_penelitian_dtps (F).xls');
$worksheet_penelitian_dtps2 = $spreadsheet_penelitian_dtps2->getActiveSheet();

// May be change
$array_penelitian_dtps = $worksheet_penelitian_dtps2->toArray();
$data_penelitian_dtps = [];

foreach($worksheet_penelitian_dtps2->getRowIterator() as $row_id => $row) {
    if($worksheet_penelitian_dtps2->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['sumber_dana'] = $array_penelitian_dtps[$row_id-1][1];
            $item['judul_ts2'] = $array_penelitian_dtps[$row_id-1][2];
			$item['judul_ts1'] = $array_penelitian_dtps[$row_id-1][3];
			$item['judul_ts'] = $array_penelitian_dtps[$row_id-1][4];
			$item['jumlah'] = $array_penelitian_dtps[$row_id-1][5];
            $data_penelitian_dtps[] = $item;
        }
    }
}

$spreadsheet_penelitian_dtps2->disconnectWorksheets();
unset($spreadsheet_penelitian_dtps2);


/**
 * Tabel 3.b.3 PkM DTPS
 */

$spreadsheet_pkm_dtps = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_pkm_dtps.xlsx');

$worksheet_pkm_dtps = $spreadsheet_pkm_dtps->getActiveSheet();

$highestRow_pkm_dtps = $worksheet_pkm_dtps->getHighestRow();

$worksheet_pkm_dtps->setAutoFilter('B1:H'.$highestRow_pkm_dtps);
$autoFilter_pkm_dtps = $worksheet_pkm_dtps->getAutoFilter();
$columnFilter_pkm_dtps = $autoFilter_pkm_dtps->getColumn('G');
$columnFilter_pkm_dtps->setFilterType(
    \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_FILTER
);
$columnFilter_pkm_dtps->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        $nama_prodi2
    );

$autoFilter_pkm_dtps->showHideRows();

$writer_pkm_dtps = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_pkm_dtps, 'Xlsx');
$writer_pkm_dtps->save('./formatted/sapto_pkm_dtps (F).xlsx');

$spreadsheet_pkm_dtps->disconnectWorksheets();
unset($spreadsheet_pkm_dtps);

// Load Format Baru
$spreadsheet_pkm_dtps2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_pkm_dtps (F).xlsx');
$worksheet_pkm_dtps2 = $spreadsheet_pkm_dtps2->getActiveSheet();

// May be change
$array_pkm_dtps = $worksheet_pkm_dtps2->toArray();
$data_pkm_dtps = [];

foreach($worksheet_pkm_dtps2->getRowIterator() as $row_id => $row) {
    if($worksheet_pkm_dtps2->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['sumber_dana'] = $array_pkm_dtps[$row_id-1][1];
            $item['judul_ts2'] = $array_pkm_dtps[$row_id-1][2];
			$item['judul_ts1'] = $array_pkm_dtps[$row_id-1][3];
			$item['judul_ts'] = $array_pkm_dtps[$row_id-1][4];
			$item['jumlah'] = $array_pkm_dtps[$row_id-1][5];
            $data_pkm_dtps[] = $item;
        }
    }
}

$spreadsheet_pkm_dtps2->disconnectWorksheets();
unset($spreadsheet_pkm_dtps2);


/**
 * Tabel 3.b.4 Publikasi Ilmiah DTPS
 */

$spreadsheet_publikasi_ilmiah_dtps = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_publikasi_ilmiah_dtps.xlsx');

$worksheet_publikasi_ilmiah_dtps = $spreadsheet_publikasi_ilmiah_dtps->getActiveSheet();

$highestRow_publikasi_ilmiah_dtps = $worksheet_publikasi_ilmiah_dtps->getHighestRow();

$worksheet_publikasi_ilmiah_dtps->setAutoFilter('B1:H'.$highestRow_publikasi_ilmiah_dtps);
$autoFilter_publikasi_ilmiah_dtps = $worksheet_publikasi_ilmiah_dtps->getAutoFilter();
$columnFilter_publikasi_ilmiah_dtps = $autoFilter_publikasi_ilmiah_dtps->getColumn('G');
$columnFilter_publikasi_ilmiah_dtps->setFilterType(
    \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_FILTER
);
$columnFilter_publikasi_ilmiah_dtps->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        $nama_prodi2
    );

$autoFilter_publikasi_ilmiah_dtps->showHideRows();

$writer_publikasi_ilmiah_dtps = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_publikasi_ilmiah_dtps, 'Xlsx');
$writer_publikasi_ilmiah_dtps->save('./formatted/sapto_publikasi_ilmiah_dtps (F).xlsx');

$spreadsheet_publikasi_ilmiah_dtps->disconnectWorksheets();
unset($spreadsheet_publikasi_ilmiah_dtps);

// Load Format Baru
$spreadsheet_publikasi_ilmiah_dtps2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_publikasi_ilmiah_dtps (F).xlsx');
$worksheet_publikasi_ilmiah_dtps2 = $spreadsheet_publikasi_ilmiah_dtps2->getActiveSheet();

// May be change
$array_publikasi_ilmiah_dtps = $worksheet_publikasi_ilmiah_dtps2->toArray();
$data_publikasi_ilmiah_dtps = [];

foreach($worksheet_publikasi_ilmiah_dtps2->getRowIterator() as $row_id => $row) {
    if($worksheet_publikasi_ilmiah_dtps2->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['sumber_dana'] = $array_publikasi_ilmiah_dtps[$row_id-1][1];
            $item['judul_ts2'] = $array_publikasi_ilmiah_dtps[$row_id-1][2];
			$item['judul_ts1'] = $array_publikasi_ilmiah_dtps[$row_id-1][3];
			$item['judul_ts'] = $array_publikasi_ilmiah_dtps[$row_id-1][4];
			$item['jumlah'] = $array_publikasi_ilmiah_dtps[$row_id-1][5];
            $data_publikasi_ilmiah_dtps[] = $item;
        }
    }
}

$spreadsheet_publikasi_ilmiah_dtps2->disconnectWorksheets();
unset($spreadsheet_publikasi_ilmiah_dtps2);


/**
 * Tabel 3.b.4 Pagelaran/Pameran/Presentasi/Publikasi Ilmiah DTPS
 */

/* $spreadsheet_pagelaran_dtps = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_pagelaran_presentasi_publikasi_dtps.xlsx');

$worksheet_pagelaran_dtps = $spreadsheet_pagelaran_dtps->getActiveSheet();

$highestRow_pagelaran_dtps = $worksheet_pagelaran_dtps->getHighestRow();

$worksheet_pagelaran_dtps->setAutoFilter('B1:H'.$highestRow_pagelaran_dtps);
$autoFilter_pagelaran_dtps = $worksheet_pagelaran_dtps->getAutoFilter();
$columnFilter_pagelaran_dtps= $autoFilter_pagelaran_dtps->getColumn('G');
$columnFilter_pagelaran_dtps->setFilterType(
    \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_FILTER
);
$columnFilter_pagelaran_dtps->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        $nama_prodi2
    );

$autoFilter_pagelaran_dtps->showHideRows();

$writer_pagelaran_dtps = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_pagelaran_dtps, 'Xlsx');
$writer_pagelaran_dtps->save('./formatted/sapto_pagelaran_presentasi_publikasi_dtps (F).xlsx');

$spreadsheet_pagelaran_dtps->disconnectWorksheets();
unset($spreadsheet_pagelaran_dtps);

// Load Format Baru
$spreadsheet_pagelaran_dtps2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_pagelaran_presentasi_publikasi_dtps (F).xlsx');
$worksheet_pagelaran_dtps2 = $spreadsheet_pagelaran_dtps2->getActiveSheet();

// May be change
$array_pagelaran_dtps = $worksheet_pagelaran_dtps2->toArray();
$data_pagelaran_dtps = [];

foreach($worksheet_pagelaran_dtps2->getRowIterator() as $row_id => $row) {
    if($worksheet_pagelaran_dtps2->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['sumber_dana'] = $array_pagelaran_dtps[$row_id-1][1];
            $item['judul_ts2'] = $array_pagelaran_dtps[$row_id-1][2];
			$item['judul_ts1'] = $array_pagelaran_dtps[$row_id-1][3];
			$item['judul_ts'] = $array_pagelaran_dtps[$row_id-1][4];
			$item['jumlah'] = $array_pagelaran_dtps[$row_id-1][5];
            $data_pagelaran_dtps[] = $item;
        }
    }
}

$spreadsheet_pagelaran_dtps2->disconnectWorksheets();
unset($spreadsheet_pagelaran_dtps2);
 */

/**
 * Tabel 3.b.5 Luaran Penelitian/PkM Lainnya oleh DTPS - HKI (Paten, Paten Sederhana)
 */

$spreadsheet_luaran_penelitian_dtps_1 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_luaran_penelitian_dtps.xls');

$worksheet_luaran_penelitian_dtps_1 = $spreadsheet_luaran_penelitian_dtps_1->getActiveSheet();

$highestRow_luaran_penelitian_dtps_1 = $worksheet_luaran_penelitian_dtps_1->getHighestRow();

$worksheet_luaran_penelitian_dtps_1->setAutoFilter('B1:E'.$highestRow_luaran_penelitian_dtps_1);
$autoFilter_luaran_penelitian_dtps_1 = $worksheet_luaran_penelitian_dtps_1->getAutoFilter();
$columnFilter_luaran_penelitian_dtps_1 = $autoFilter_luaran_penelitian_dtps_1->getColumn('E');
$columnFilter_luaran_penelitian_dtps_1->setFilterType(
    \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_FILTER
);
$columnFilter_luaran_penelitian_dtps_1->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        $nama_prodi2
    );

$autoFilter_luaran_penelitian_dtps_1->showHideRows();

$writer_luaran_penelitian_dtps_1 = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_luaran_penelitian_dtps_1, 'Xls');
$writer_luaran_penelitian_dtps_1->save('./formatted/sapto_luaran_penelitian_dtps (F).xls');

$spreadsheet_luaran_penelitian_dtps_1->disconnectWorksheets();
unset($spreadsheet_luaran_penelitian_dtps_1);

// Load Format Baru
$spreadsheet_luaran_penelitian_dtps_12 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_luaran_penelitian_dtps (F).xls');
$worksheet_luaran_penelitian_dtps_12 = $spreadsheet_luaran_penelitian_dtps_12->getActiveSheet();

// May be change
$array_luaran_penelitian_dtps_1 = $worksheet_luaran_penelitian_dtps_12->toArray();
$data_luaran_penelitian_dtps_1 = [];

foreach($worksheet_luaran_penelitian_dtps_12->getRowIterator() as $row_id => $row) {
    if($worksheet_luaran_penelitian_dtps_12->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['luaran_penelitian'] = $array_luaran_penelitian_dtps_1[$row_id-1][1];
            $item['tahun'] = $array_luaran_penelitian_dtps_1[$row_id-1][2];
			$item['keterangan'] = $array_luaran_penelitian_dtps_1[$row_id-1][3];
            $data_luaran_penelitian_dtps_1[] = $item;
        }
    }
}

$spreadsheet_luaran_penelitian_dtps_12->disconnectWorksheets();
unset($spreadsheet_luaran_penelitian_dtps_12);


/**
 * Tabel 3.b.5 Luaran Penelitian/PkM Lainnya oleh DTPS - HKI (Hak Cipta, Desain Produk Industri, dll.)
 */

$spreadsheet_luaran_penelitian_dtps_2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_luaran_penelitian_dtps_2.xls');

$worksheet_luaran_penelitian_dtps_2 = $spreadsheet_luaran_penelitian_dtps_2->getActiveSheet();

$highestRow_luaran_penelitian_dtps_2 = $worksheet_luaran_penelitian_dtps_2->getHighestRow();

$worksheet_luaran_penelitian_dtps_2->setAutoFilter('B1:E'.$highestRow_luaran_penelitian_dtps_2);
$autoFilter_luaran_penelitian_dtps_2 = $worksheet_luaran_penelitian_dtps_2->getAutoFilter();
$columnFilter_luaran_penelitian_dtps_2 = $autoFilter_luaran_penelitian_dtps_2->getColumn('E');
$columnFilter_luaran_penelitian_dtps_2->setFilterType(
    \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_FILTER
);
$columnFilter_luaran_penelitian_dtps_2->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        $nama_prodi2
    );

$autoFilter_luaran_penelitian_dtps_2->showHideRows();

$writer_luaran_penelitian_dtps_2 = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_luaran_penelitian_dtps_2, 'Xls');
$writer_luaran_penelitian_dtps_2->save('./formatted/sapto_luaran_penelitian_dtps_2 (F).xls');

$spreadsheet_luaran_penelitian_dtps_2->disconnectWorksheets();
unset($spreadsheet_luaran_penelitian_dtps_2);

// Load Format Baru
$spreadsheet_luaran_penelitian_dtps_22 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_luaran_penelitian_dtps_2 (F).xls');
$worksheet_luaran_penelitian_dtps_22 = $spreadsheet_luaran_penelitian_dtps_22->getActiveSheet();

// May be change
$array_luaran_penelitian_dtps_2 = $worksheet_luaran_penelitian_dtps_22->toArray();
$data_luaran_penelitian_dtps_2 = [];

foreach($worksheet_luaran_penelitian_dtps_22->getRowIterator() as $row_id => $row) {
    if($worksheet_luaran_penelitian_dtps_22->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['luaran_penelitian'] = $array_luaran_penelitian_dtps_2[$row_id-1][1];
            $item['tahun'] = $array_luaran_penelitian_dtps_2[$row_id-1][2];
			$item['keterangan'] = $array_luaran_penelitian_dtps_2[$row_id-1][3];
            $data_luaran_penelitian_dtps_2[] = $item;
        }
    }
}

$spreadsheet_luaran_penelitian_dtps_22->disconnectWorksheets();
unset($spreadsheet_luaran_penelitian_dtps_22);


/**
 * Tabel 3.b.5 Luaran Penelitian/PkM Lainnya oleh DTPS - Teknologi Tepat Guna, Produk, Karya Seni, Rekayasa Sosial
 */

$spreadsheet_luaran_penelitian_dtps_3 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_luaran_penelitian_dtps_3.xls');

$worksheet_luaran_penelitian_dtps_3 = $spreadsheet_luaran_penelitian_dtps_3->getActiveSheet();

$highestRow_luaran_penelitian_dtps_3 = $worksheet_luaran_penelitian_dtps_3->getHighestRow();

$worksheet_luaran_penelitian_dtps_3->setAutoFilter('B1:E'.$highestRow_luaran_penelitian_dtps_3);
$autoFilter_luaran_penelitian_dtps_3 = $worksheet_luaran_penelitian_dtps_3->getAutoFilter();
$columnFilter_luaran_penelitian_dtps_3 = $autoFilter_luaran_penelitian_dtps_3->getColumn('E');
$columnFilter_luaran_penelitian_dtps_3->setFilterType(
    \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_FILTER
);
$columnFilter_luaran_penelitian_dtps_3->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        $nama_prodi2
    );

$autoFilter_luaran_penelitian_dtps_3->showHideRows();

$writer_luaran_penelitian_dtps_3 = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_luaran_penelitian_dtps_3, 'Xls');
$writer_luaran_penelitian_dtps_3->save('./formatted/sapto_luaran_penelitian_dtps_3 (F).xls');

$spreadsheet_luaran_penelitian_dtps_3->disconnectWorksheets();
unset($spreadsheet_luaran_penelitian_dtps_3);

// Load Format Baru
$spreadsheet_luaran_penelitian_dtps_32 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_luaran_penelitian_dtps_3 (F).xls');
$worksheet_luaran_penelitian_dtps_32 = $spreadsheet_luaran_penelitian_dtps_32->getActiveSheet();

// May be change
$array_luaran_penelitian_dtps_3 = $worksheet_luaran_penelitian_dtps_32->toArray();
$data_luaran_penelitian_dtps_3 = [];

foreach($worksheet_luaran_penelitian_dtps_32->getRowIterator() as $row_id => $row) {
    if($worksheet_luaran_penelitian_dtps_32->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['luaran_penelitian'] = $array_luaran_penelitian_dtps_3[$row_id-1][1];
            $item['tahun'] = $array_luaran_penelitian_dtps_3[$row_id-1][2];
			$item['keterangan'] = $array_luaran_penelitian_dtps_3[$row_id-1][3];
            $data_luaran_penelitian_dtps_3[] = $item;
        }
    }
}

$spreadsheet_luaran_penelitian_dtps_32->disconnectWorksheets();
unset($spreadsheet_luaran_penelitian_dtps_32);


/**
 * Tabel 3.b.5 Luaran Penelitian/PkM Lainnya oleh DTPS - Buku Ber-ISBN, Book Chapter
 */

$spreadsheet_luaran_penelitian_dtps_4 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_buku_ber_isbn.xls');

$worksheet_luaran_penelitian_dtps_4 = $spreadsheet_luaran_penelitian_dtps_4->getActiveSheet();

$highestRow_luaran_penelitian_dtps_4 = $worksheet_luaran_penelitian_dtps_4->getHighestRow();

$worksheet_luaran_penelitian_dtps_4->setAutoFilter('B1:F'.$highestRow_luaran_penelitian_dtps_4);
$autoFilter_luaran_penelitian_dtps_4 = $worksheet_luaran_penelitian_dtps_4->getAutoFilter();
$columnFilter_luaran_penelitian_dtps_4 = $autoFilter_luaran_penelitian_dtps_4->getColumn('F');
$columnFilter_luaran_penelitian_dtps_4->setFilterType(
    \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_FILTER
);
$columnFilter_luaran_penelitian_dtps_4->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        $nama_prodi2
    );

$autoFilter_luaran_penelitian_dtps_4->showHideRows();

$writer_luaran_penelitian_dtps_4 = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_luaran_penelitian_dtps_4, 'Xls');
$writer_luaran_penelitian_dtps_4->save('./formatted/sapto_buku_ber_isbn (F).xls');

$spreadsheet_luaran_penelitian_dtps_4->disconnectWorksheets();
unset($spreadsheet_luaran_penelitian_dtps_4);

// Load Format Baru
$spreadsheet_luaran_penelitian_dtps_42 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_buku_ber_isbn (F).xls');
$worksheet_luaran_penelitian_dtps_42 = $spreadsheet_luaran_penelitian_dtps_42->getActiveSheet();

// May be change
$array_luaran_penelitian_dtps_4 = $worksheet_luaran_penelitian_dtps_42->toArray();
$data_luaran_penelitian_dtps_4 = [];

foreach($worksheet_luaran_penelitian_dtps_42->getRowIterator() as $row_id => $row) {
    if($worksheet_luaran_penelitian_dtps_42->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['luaran_penelitian'] = $array_luaran_penelitian_dtps_4[$row_id-1][1];
            $item['tahun'] = $array_luaran_penelitian_dtps_4[$row_id-1][2];
			$item['keterangan'] = $array_luaran_penelitian_dtps_4[$row_id-1][3];
            $data_luaran_penelitian_dtps_4[] = $item;
        }
    }
}

$spreadsheet_luaran_penelitian_dtps_42->disconnectWorksheets();
unset($spreadsheet_luaran_penelitian_dtps_42);


/**
 * Tabel 3.b.6 Karya Ilmiah DTPS yang Disitasi
 */
$spreadsheet_karya_dtps_disitasi = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_karya_ilmiah_disitasi.xls');

$worksheet_karya_dtps_disitasi = $spreadsheet_karya_dtps_disitasi->getActiveSheet();

$highestRow_karya_dtps_disitasi = $worksheet_karya_dtps_disitasi->getHighestRow();

$worksheet_karya_dtps_disitasi->setAutoFilter('B1:F'.$highestRow_karya_dtps_disitasi);
$autoFilter_karya_dtps_disitasi = $worksheet_karya_dtps_disitasi->getAutoFilter();
$columnFilter_karya_dtps_disitasi = $autoFilter_karya_dtps_disitasi->getColumn('E');
$columnFilter_karya_dtps_disitasi->setFilterType(
    \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_FILTER
);
$columnFilter_karya_dtps_disitasi->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        $nama_prodi2
    );

$autoFilter_karya_dtps_disitasi->showHideRows();

$writer_karya_dtps_disitasi = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_karya_dtps_disitasi, 'Xls');
$writer_karya_dtps_disitasi->save('./formatted/sapto_karya_ilmiah_disitasi (F).xls');

$spreadsheet_karya_dtps_disitasi->disconnectWorksheets();
unset($spreadsheet_karya_dtps_disitasi);

// Load Format Baru
$spreadsheet_karya_dtps_disitasi2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_karya_ilmiah_disitasi (F).xls');
$worksheet_karya_dtps_disitasi2 = $spreadsheet_karya_dtps_disitasi2->getActiveSheet();

// May be change
$array_karya_dtps_disitasi = $worksheet_karya_dtps_disitasi2->toArray();
$data_karya_dtps_disitasi = [];

foreach($worksheet_karya_dtps_disitasi2->getRowIterator() as $row_id => $row) {
    if($worksheet_karya_dtps_disitasi2->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['nama_dosen'] = $array_karya_dtps_disitasi[$row_id-1][1];
            $item['judul_artikel'] = $array_karya_dtps_disitasi[$row_id-1][2];
			$item['jml_sitasi'] = $array_karya_dtps_disitasi[$row_id-1][3];
            $data_karya_dtps_disitasi[] = $item;
        }
    }
}

$spreadsheet_karya_dtps_disitasi2->disconnectWorksheets();
unset($spreadsheet_karya_dtps_disitasi2);


/**
 * Tabel 3.b.7 Produk/Jasa DTPS yang Diadopsi oleh Industri/Masyarakat
 */

/* $spreadsheet_produk_dtps_diadopsi = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_produk_jasa_dtps_masyarakat.xlsx');

$worksheet_produk_dtps_diadopsi = $spreadsheet_produk_dtps_diadopsi->getActiveSheet();

$highestRow_produk_dtps_diadopsi = $worksheet_produk_dtps_diadopsi->getHighestRow();

$worksheet_produk_dtps_diadopsi->setAutoFilter('B1:G'.$highestRow_karya_dtps_disitasi);
$autoFilter_produk_dtps_diadopsi = $worksheet_produk_dtps_diadopsi->getAutoFilter();
$columnFilter_produk_dtps_diadopsi = $autoFilter_produk_dtps_diadopsi->getColumn('F');
$columnFilter_produk_dtps_diadopsi->setFilterType(
    \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_FILTER
);
$columnFilter_produk_dtps_diadopsi->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        $nama_prodi2
    );

$autoFilter_produk_dtps_diadopsi->showHideRows();

$writer_produk_dtps_diadopsi = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_produk_dtps_diadopsi, 'Xlsx');
$writer_produk_dtps_diadopsi->save('./formatted/sapto_produk_jasa_dtps_masyarakat (F).xlsx');

$spreadsheet_produk_dtps_diadopsi->disconnectWorksheets();
unset($spreadsheet_produk_dtps_diadopsi);

// Load Format Baru
$spreadsheet_produk_dtps_diadopsi2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_produk_jasa_dtps_masyarakat (F).xlsx');
$worksheet_produk_dtps_diadopsi2 = $spreadsheet_produk_dtps_diadopsi2->getActiveSheet();

// May be change
$array_produk_dtps_diadopsi = $worksheet_produk_dtps_diadopsi2->toArray();
$data_produk_dtps_diadopsi = [];

foreach($worksheet_produk_dtps_diadopsi2->getRowIterator() as $row_id => $row) {
    if($worksheet_produk_dtps_diadopsi2->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['nama_dosen'] = $array_produk_dtps_diadopsi[$row_id-1][1];
            $item['nama_produk'] = $array_produk_dtps_diadopsi[$row_id-1][2];
			$item['desk_produk'] = $array_produk_dtps_diadopsi[$row_id-1][3];
			$item['bukti'] = $array_produk_dtps_diadopsi[$row_id-1][4];
            $data_produk_dtps_diadopsi[] = $item;
        }
    }
}

$spreadsheet_produk_dtps_diadopsi2->disconnectWorksheets();
unset($spreadsheet_produk_dtps_diadopsi2); */


/**
 * Tabel 4 Penggunaan Dana
 */

$spreadsheet_penggunaan_dana = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_penggunaan_dana.xlsx');

$worksheet_penggunaan_dana = $spreadsheet_penggunaan_dana->getActiveSheet();

$worksheet_penggunaan_dana->insertNewColumnBefore('D', 4);
$worksheet_penggunaan_dana->insertNewColumnBefore('K', 1);

$worksheet_penggunaan_dana->getCell('G1')->setValue('Rata-rata');
$worksheet_penggunaan_dana->getCell('K1')->setValue('Rata-rata');

$highestRow_penggunaan_dana = $worksheet_penggunaan_dana->getHighestRow();

$array_upps_dana = $worksheet_penggunaan_dana->rangeToArray('L1:N'.$highestRow_penggunaan_dana, NULL, TRUE, TRUE, TRUE);
$worksheet_penggunaan_dana->fromArray($array_upps_dana, NULL, 'D1');

$worksheet_penggunaan_dana->setAutoFilter('B1:K'.$highestRow_penggunaan_dana);
$autoFilter_penggunaan_dana = $worksheet_penggunaan_dana->getAutoFilter();
$columnFilter_penggunaan_dana = $autoFilter_penggunaan_dana->getColumn('B');
$columnFilter_penggunaan_dana->setFilterType(
    \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_FILTER
);
$columnFilter_penggunaan_dana->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        $nama_prodi
    );

$autoFilter_penggunaan_dana->showHideRows();

$array_penggunaan_dana = $worksheet_penggunaan_dana->toArray();
$data_penggunaan_dana = [];

foreach($worksheet_penggunaan_dana->getRowIterator() as $row_id => $row) {
    if($worksheet_penggunaan_dana->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['jenis_penggunaan'] = $array_penggunaan_dana[$row_id-1][2];
            $item['upps_ts2'] = $array_penggunaan_dana[$row_id-1][3];
			$item['upps_ts1'] = $array_penggunaan_dana[$row_id-1][4];
			$item['upps_ts'] = $array_penggunaan_dana[$row_id-1][5];
			$item['upps_rata'] = $array_penggunaan_dana[$row_id-1][6];
			$item['ps_ts2'] = $array_penggunaan_dana[$row_id-1][7];
			$item['ps_ts1'] = $array_penggunaan_dana[$row_id-1][8];
			$item['ps_ts'] = $array_penggunaan_dana[$row_id-1][9];
			$item['ps_rata'] = $array_penggunaan_dana[$row_id-1][10];
            $data_penggunaan_dana[] = $item;
        }
    }
}

$worksheet_penggunaan_dana2 = new \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet($spreadsheet_penggunaan_dana, 'Sheet 2');
$spreadsheet_penggunaan_dana->addSheet($worksheet_penggunaan_dana2);

$worksheet_penggunaan_dana2 = $spreadsheet_penggunaan_dana->getSheetByName('Sheet 2');
$worksheet_penggunaan_dana2->fromArray($data_penggunaan_dana, NULL, 'A1');

$writer_penggunaan_dana = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_penggunaan_dana, 'Xlsx');
$writer_penggunaan_dana->save('./formatted/sapto_penggunaan_dana (F).xlsx');

$spreadsheet_penggunaan_dana->disconnectWorksheets();
unset($spreadsheet_penggunaan_dana);

// Load Format Baru
$spreadsheet_penggunaan_dana2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_penggunaan_dana (F).xlsx');
$worksheet_penggunaan_dana3 = $spreadsheet_penggunaan_dana2->getSheetByName('Sheet 2');

// May be change
$biaya_dosen = $worksheet_penggunaan_dana3->rangeToArray('B10:I10', NULL, TRUE, TRUE, TRUE);
$biaya_tendik = $worksheet_penggunaan_dana3->rangeToArray('B1:I1', NULL, TRUE, TRUE, TRUE);
$biaya_ops_pembelajaran = $worksheet_penggunaan_dana3->rangeToArray('B5:I5', NULL, TRUE, TRUE, TRUE);
$biaya_ops_tdk_langsung = $worksheet_penggunaan_dana3->rangeToArray('B4:I4', NULL, TRUE, TRUE, TRUE);
$biaya_ops_mhs = $worksheet_penggunaan_dana3->rangeToArray('B6:I6', NULL, TRUE, TRUE, TRUE);
$biaya_penelitian = $worksheet_penggunaan_dana3->rangeToArray('B3:I3', NULL, TRUE, TRUE, TRUE);
$biaya_pkm = $worksheet_penggunaan_dana3->rangeToArray('B2:I2', NULL, TRUE, TRUE, TRUE);
$biaya_investasi = $worksheet_penggunaan_dana3->rangeToArray('B7:I9', NULL, TRUE, TRUE, TRUE);

$spreadsheet_penggunaan_dana2->disconnectWorksheets();
unset($spreadsheet_penggunaan_dana2);


/**
 * Tabel 5.a Kurikulum, Capaian Pembelajaran, dan Rencana Pembelajaran
 */

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


$worksheet_kurikulum->setAutoFilter('B1:V'.$highestRow_kurikulum);
$autoFilter_kurikulum = $worksheet_kurikulum->getAutoFilter();
$columnFilter_kurikulum = $autoFilter_kurikulum->getColumn('U');
$columnFilter_kurikulum->setFilterType(
    \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_FILTER
);
$columnFilter_kurikulum->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        $nama_prodi2
    );

$autoFilter_kurikulum->showHideRows();


$writer_kurikulum = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_kurikulum, 'Xls');
$writer_kurikulum->save('./formatted/sapto_kurikulum_capaian_rencana (F).xls');

$spreadsheet_kurikulum->disconnectWorksheets();
unset($spreadsheet_kurikulum);

// Load Format Baru
$spreadsheet_kurikulum2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_kurikulum_capaian_rencana (F).xls');
$worksheet_kurikulum2 = $spreadsheet_kurikulum2->getActiveSheet();

// May be change
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

$spreadsheet_integrasi_penelitian = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_integrasi_kegiatan_penelitian.xlsx');

$worksheet_integrasi_penelitian = $spreadsheet_integrasi_penelitian->getActiveSheet();
$highestRow_integrasi_penelitian = $worksheet_integrasi_penelitian->getHighestRow();

$worksheet_integrasi_penelitian->setAutoFilter('B1:G'.$highestRow_integrasi_penelitian);
$autoFilter_integrasi_penelitian = $worksheet_integrasi_penelitian->getAutoFilter();
$columnFilter_integrasi_penelitian = $autoFilter_integrasi_penelitian->getColumn('G');
$columnFilter_integrasi_penelitian->setFilterType(
    \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_FILTER
);
$columnFilter_integrasi_penelitian->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        $nama_prodi2
    );

$autoFilter_integrasi_penelitian->showHideRows();

$writer_integrasi_penelitian = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_integrasi_penelitian, 'Xlsx');
$writer_integrasi_penelitian->save('./formatted/sapto_integrasi_kegiatan_penelitian (F).xlsx');

$spreadsheet_integrasi_penelitian->disconnectWorksheets();
unset($spreadsheet_integrasi_penelitian);

// Load Format Baru
$spreadsheet_integrasi_penelitian2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_integrasi_kegiatan_penelitian (F).xlsx');
$worksheet_integrasi_penelitian2 = $spreadsheet_integrasi_penelitian2->getActiveSheet();

// May be change
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

$spreadsheet_kepuasan_mahasiswa = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_kepuasan_mahasiswa.xlsx');

$worksheet_kepuasan_mahasiswa = $spreadsheet_kepuasan_mahasiswa->getActiveSheet();
$highestRow_kepuasan_mahasiswa = $worksheet_kepuasan_mahasiswa->getHighestRow();

$worksheet_kepuasan_mahasiswa->setAutoFilter('B1:I'.$highestRow_kepuasan_mahasiswa);
$autoFilter_kepuasan_mahasiswa = $worksheet_kepuasan_mahasiswa->getAutoFilter();
$columnFilter_kepuasan_mahasiswa = $autoFilter_kepuasan_mahasiswa->getColumn('H');
$columnFilter_kepuasan_mahasiswa->setFilterType(
    \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_FILTER
);
$columnFilter_kepuasan_mahasiswa->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        $nama_prodi2
    );

$autoFilter_kepuasan_mahasiswa->showHideRows();

$writer_kepuasan_mahasiswa = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_kepuasan_mahasiswa, 'Xlsx');
$writer_kepuasan_mahasiswa->save('./formatted/sapto_kepuasan_mahasiswa (F).xlsx');

$spreadsheet_kepuasan_mahasiswa->disconnectWorksheets();
unset($spreadsheet_kepuasan_mahasiswa);

// Load Format Baru
$spreadsheet_kepuasan_mahasiswa2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_kepuasan_mahasiswa (F).xlsx');
$worksheet_kepuasan_mahasiswa2 = $spreadsheet_kepuasan_mahasiswa2->getActiveSheet();

// May be change
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
 * Tabel 6.a Penelitian DTPS yang Melibatkan Mahasiswa
 */

$spreadsheet_penelitian_dtps_mhs = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_penelitian_dtps_mahasiswa.xlsx');

$worksheet_penelitian_dtps_mhs = $spreadsheet_penelitian_dtps_mhs->getActiveSheet();

$highestRow_penelitian_dtps_mhs = $worksheet_penelitian_dtps_mhs->getHighestRow();

$worksheet_penelitian_dtps_mhs->setAutoFilter('B1:G'.$highestRow_penelitian_dtps_mhs);
$autoFilter_penelitian_dtps_mhs = $worksheet_penelitian_dtps_mhs->getAutoFilter();
$columnFilter_penelitian_dtps_mhs = $autoFilter_penelitian_dtps_mhs->getColumn('G');
$columnFilter_penelitian_dtps_mhs->setFilterType(
    \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_FILTER
);
$columnFilter_penelitian_dtps_mhs->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        $nama_prodi2
    );

$autoFilter_penelitian_dtps_mhs->showHideRows();

$writer_penelitian_dtps_mhs = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_penelitian_dtps_mhs, 'Xlsx');
$writer_penelitian_dtps_mhs->save('./formatted/sapto_penelitian_dtps_mahasiswa (F).xlsx');

$spreadsheet_penelitian_dtps_mhs->disconnectWorksheets();
unset($spreadsheet_penelitian_dtps_mhs);

// Load Format Baru
$spreadsheet_penelitian_dtps_mhs2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_penelitian_dtps_mahasiswa (F).xlsx');
$worksheet_penelitian_dtps_mhs2 = $spreadsheet_penelitian_dtps_mhs2->getActiveSheet();

// May be change
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

$spreadsheet_penelitian_dtps_rujukan_tesis = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_penelitian_dtps_rujukan_tesis.xlsx');

$worksheet_penelitian_dtps_rujukan_tesis = $spreadsheet_penelitian_dtps_rujukan_tesis->getActiveSheet();

$highestRow_penelitian_dtps_rujukan_tesis = $worksheet_penelitian_dtps_rujukan_tesis->getHighestRow();

$worksheet_penelitian_dtps_rujukan_tesis->setAutoFilter('B1:G'.$highestRow_penelitian_dtps_rujukan_tesis);
$autoFilter_penelitian_dtps_rujukan_tesis = $worksheet_penelitian_dtps_rujukan_tesis->getAutoFilter();
$columnFilter_penelitian_dtps_rujukan_tesis = $autoFilter_penelitian_dtps_rujukan_tesis->getColumn('G');
$columnFilter_penelitian_dtps_rujukan_tesis->setFilterType(
    \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_FILTER
);
$columnFilter_penelitian_dtps_rujukan_tesis->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        $nama_prodi
    );

$autoFilter_penelitian_dtps_rujukan_tesis->showHideRows();

$writer_penelitian_dtps_rujukan_tesis = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_penelitian_dtps_rujukan_tesis, 'Xlsx');
$writer_penelitian_dtps_rujukan_tesis->save('./formatted/sapto_penelitian_dtps_rujukan_tesis (F).xlsx');

$spreadsheet_penelitian_dtps_rujukan_tesis->disconnectWorksheets();
unset($spreadsheet_penelitian_dtps_rujukan_tesis);

// Load Format Baru
$spreadsheet_penelitian_dtps_rujukan_tesis2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_penelitian_dtps_rujukan_tesis (F).xlsx');
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
 * Tabel 7 PkM DTPS yang Melibatkan Mahasiswa
 */

/* $spreadsheet_pkm_dtps_mhs = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_pkm_dtps_mahasiswa.xls');

$worksheet_pkm_dtps_mhs = $spreadsheet_pkm_dtps_mhs->getActiveSheet();

$highestRow_pkm_dtps_mhs = $worksheet_pkm_dtps_mhs->getHighestRow();

$worksheet_pkm_dtps_mhs->setAutoFilter('B1:G'.$highestRow_pkm_dtps_mhs);
$autoFilter_pkm_dtps_mhs = $worksheet_pkm_dtps_mhs->getAutoFilter();
$columnFilter_pkm_dtps_mhs = $autoFilter_pkm_dtps_mhs->getColumn('G');
$columnFilter_pkm_dtps_mhs->setFilterType(
    \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_FILTER
);
$columnFilter_pkm_dtps_mhs->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        $nama_prodi2
    );

$autoFilter_pkm_dtps_mhs->showHideRows();

$writer_pkm_dtps_mhs = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_pkm_dtps_mhs, 'Xls');
$writer_pkm_dtps_mhs->save('./formatted/sapto_pkm_dtps_mahasiswa (F).xls');

$spreadsheet_pkm_dtps_mhs->disconnectWorksheets();
unset($spreadsheet_pkm_dtps_mhs);

// Load Format Baru
$spreadsheet_pkm_dtps_mhs2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_pkm_dtps_mahasiswa (F).xls');
$worksheet_pkm_dtps_mhs2 = $spreadsheet_pkm_dtps_mhs2->getActiveSheet();

// May be change
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
} */


/**
 * Tabel 8a IPK Lulusan
 */

$spreadsheet_ipk_lulusan = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_ipk_lulusan.xlsx');

$worksheet_ipk_lulusan = $spreadsheet_ipk_lulusan->getActiveSheet();

$highestRow_ipk_lulusan = $worksheet_ipk_lulusan->getHighestRow();

$worksheet_ipk_lulusan->setAutoFilter('B1:I'.$highestRow_ipk_lulusan);
$autoFilter_ipk_lulusan = $worksheet_ipk_lulusan->getAutoFilter();
$columnFilter_ipk_lulusan = $autoFilter_ipk_lulusan->getColumn('B');
$columnFilter_ipk_lulusan->setFilterType(
    \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_FILTER
);
$columnFilter_ipk_lulusan->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        $nama_prodi
    );

$autoFilter_ipk_lulusan->showHideRows();

$writer_ipk_lulusan = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_ipk_lulusan, 'Xlsx');
$writer_ipk_lulusan->save('./formatted/sapto_ipk_lulusan (F).xlsx');

$spreadsheet_ipk_lulusan->disconnectWorksheets();
unset($spreadsheet_ipk_lulusan);

// Load Format Baru
$spreadsheet_ipk_lulusan2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_ipk_lulusan (F).xlsx');
$worksheet_ipk_lulusan2 = $spreadsheet_ipk_lulusan2->getActiveSheet();

// May be change
$array_ipk_lulusan = $worksheet_ipk_lulusan2->toArray();
$data_ipk_lulusan = [];

foreach($worksheet_ipk_lulusan2->getRowIterator() as $row_id => $row) {
    if($worksheet_ipk_lulusan2->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['tahun_akademik'] = $array_ipk_lulusan[$row_id-1][3];
            $item['jml_lulusan'] = $array_ipk_lulusan[$row_id-1][5];
			$item['ipk_min'] = $array_ipk_lulusan[$row_id-1][6];
			$item['ipk_rata'] = $array_ipk_lulusan[$row_id-1][7];
			$item['ipk_maks'] = $array_ipk_lulusan[$row_id-1][8];
            $data_ipk_lulusan[] = $item;
        }
    }
}

$spreadsheet_ipk_lulusan2->disconnectWorksheets();
unset($spreadsheet_ipk_lulusan2);


/**
 * Tabel 8.b.1 Prestasi Akademik Mahasiswa
 */

$spreadsheet_prestasi_akademik = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_prestasi_mhs.xlsx');

$worksheet_prestasi_akademik = $spreadsheet_prestasi_akademik->getActiveSheet();

$worksheet_prestasi_akademik->insertNewColumnBefore('H', 3);

$highestRow_prestasi_akademik = $worksheet_prestasi_akademik->getHighestRow();

for($row = 2;$row <= $highestRow_prestasi_akademik; $row++) {
	$worksheet_prestasi_akademik->setCellValue('H'.$row, '=IF(E'.$row.'=1;"V";"")');
	$worksheet_prestasi_akademik->getCell('H'.$row)->getStyle()->setQuotePrefix(true);
	
	$worksheet_prestasi_akademik->setCellValue('I'.$row, '=IF(F'.$row.'=1;"V";"")');
	$worksheet_prestasi_akademik->getCell('I'.$row)->getStyle()->setQuotePrefix(true);
	
	$worksheet_prestasi_akademik->setCellValue('J'.$row, '=IF(G'.$row.'=1;"V";"")');
	$worksheet_prestasi_akademik->getCell('J'.$row)->getStyle()->setQuotePrefix(true);
}

$worksheet_prestasi_akademik->setAutoFilter('B1:L'.$highestRow_prestasi_akademik);
$autoFilter_prestasi_akademik = $worksheet_prestasi_akademik->getAutoFilter();
$columnFilter_prestasi_akademik = $autoFilter_prestasi_akademik->getColumn('B');
$columnFilter_prestasi_akademik2 = $autoFilter_prestasi_akademik->getColumn('L');
$columnFilter_prestasi_akademik->setFilterType(
    \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_FILTER
);
$columnFilter_prestasi_akademik->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        $nama_prodi
    );
$columnFilter_prestasi_akademik2->setFilterType(
    \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_FILTER
);
$columnFilter_prestasi_akademik2->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        1
    );

$autoFilter_prestasi_akademik->showHideRows();

$writer_prestasi_akademik = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_prestasi_akademik, 'Xls');
$writer_prestasi_akademik->save('./formatted/sapto_prestasi_mhs_1 (F).xls');

$spreadsheet_prestasi_akademik->disconnectWorksheets();
unset($spreadsheet_prestasi_akademik);

// Load Format Baru
$spreadsheet_prestasi_akademik2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_prestasi_mhs_1 (F).xls');
$worksheet_prestasi_akademik2 = $spreadsheet_prestasi_akademik2->getActiveSheet();

// May be change
$array_prestasi_akademik = $worksheet_prestasi_akademik2->toArray();
$data_prestasi_akademik = [];

foreach($worksheet_prestasi_akademik2->getRowIterator() as $row_id => $row) {
    if($worksheet_prestasi_akademik2->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['nama_kegiatan'] = $array_prestasi_akademik[$row_id-1][2];
            $item['tahun'] = $array_prestasi_akademik[$row_id-1][3];
			$item['lokal'] = $array_prestasi_akademik[$row_id-1][7];
			$item['nasional'] = $array_prestasi_akademik[$row_id-1][8];
			$item['internasional'] = $array_prestasi_akademik[$row_id-1][9];
			$item['capaian'] = $array_prestasi_akademik[$row_id-1][10];
            $data_prestasi_akademik[] = $item;
        }
    }
}

$spreadsheet_prestasi_akademik2->disconnectWorksheets();
unset($spreadsheet_ipk_prestasi_akademik2);


/**
 * Tabel 8.b.2 Prestasi Non Akademik Mahasiswa
 */

$spreadsheet_prestasi_non_akademik = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_prestasi_mhs.xlsx');

$worksheet_prestasi_non_akademik = $spreadsheet_prestasi_non_akademik->getActiveSheet();

$worksheet_prestasi_non_akademik->insertNewColumnBefore('H', 3);

$highestRow_prestasi_non_akademik = $worksheet_prestasi_non_akademik->getHighestRow();

for($row = 2;$row <= $highestRow_prestasi_non_akademik; $row++) {
	$worksheet_prestasi_non_akademik->setCellValue('H'.$row, '=IF(E'.$row.'=1;"V";"")');
	$worksheet_prestasi_non_akademik->getCell('H'.$row)->getStyle()->setQuotePrefix(true);
	
	$worksheet_prestasi_non_akademik->setCellValue('I'.$row, '=IF(F'.$row.'=1;"V";"")');
	$worksheet_prestasi_non_akademik->getCell('I'.$row)->getStyle()->setQuotePrefix(true);
	
	$worksheet_prestasi_non_akademik->setCellValue('J'.$row, '=IF(G'.$row.'=1;"V";"")');
	$worksheet_prestasi_non_akademik->getCell('J'.$row)->getStyle()->setQuotePrefix(true);
}

$worksheet_prestasi_non_akademik->setAutoFilter('B1:L'.$highestRow_prestasi_non_akademik);
$autoFilter_prestasi_non_akademik = $worksheet_prestasi_non_akademik->getAutoFilter();
$columnFilter_prestasi_non_akademik = $autoFilter_prestasi_non_akademik->getColumn('B');
$columnFilter_prestasi_non_akademik2 = $autoFilter_prestasi_non_akademik->getColumn('L');
$columnFilter_prestasi_non_akademik->setFilterType(
    \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_FILTER
);
$columnFilter_prestasi_non_akademik->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        $nama_prodi
    );
$columnFilter_prestasi_non_akademik2->setFilterType(
    \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_FILTER
);
$columnFilter_prestasi_non_akademik2->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        0
    );

$autoFilter_prestasi_non_akademik->showHideRows();

$writer_prestasi_non_akademik = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_prestasi_non_akademik, 'Xls');
$writer_prestasi_non_akademik->save('./formatted/sapto_prestasi_mhs_2 (F).xls');

$spreadsheet_prestasi_non_akademik->disconnectWorksheets();
unset($spreadsheet_prestasi_non_akademik);

// Load Format Baru
$spreadsheet_prestasi_non_akademik2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_prestasi_mhs_2 (F).xls');
$worksheet_prestasi_non_akademik2 = $spreadsheet_prestasi_non_akademik2->getActiveSheet();

// May be change
$array_prestasi_non_akademik = $worksheet_prestasi_non_akademik2->toArray();
$data_prestasi_non_akademik = [];

foreach($worksheet_prestasi_non_akademik2->getRowIterator() as $row_id => $row) {
    if($worksheet_prestasi_non_akademik2->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['nama_kegiatan'] = $array_prestasi_non_akademik[$row_id-1][2];
            $item['tahun'] = $array_prestasi_non_akademik[$row_id-1][3];
			$item['lokal'] = $array_prestasi_non_akademik[$row_id-1][7];
			$item['nasional'] = $array_prestasi_non_akademik[$row_id-1][8];
			$item['internasional'] = $array_prestasi_non_akademik[$row_id-1][9];
			$item['capaian'] = $array_prestasi_non_akademik[$row_id-1][10];
            $data_prestasi_non_akademik[] = $item;
        }
    }
}

$spreadsheet_prestasi_non_akademik2->disconnectWorksheets();
unset($spreadsheet_ipk_prestasi_non_akademik2);


/**
 * Tabel 8.c Masa Studi Lulusan
 */

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

$worksheet_masa_studi->setAutoFilter('B1:O'.$highestRow_masa_studi);
$autoFilter_masa_studi = $worksheet_masa_studi->getAutoFilter();
$columnFilter_masa_studi = $autoFilter_masa_studi->getColumn('B');
$columnFilter_masa_studi->setFilterType(
    \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_FILTER
);
$columnFilter_masa_studi->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        $nama_prodi
    );

$autoFilter_masa_studi->showHideRows();

$writer_masa_studi = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_masa_studi, 'Xls');
$writer_masa_studi->save('./formatted/sapto_masa_studi_lulusan (F).xls');

$spreadsheet_masa_studi->disconnectWorksheets();
unset($spreadsheet_masa_studi);

// Load Format Baru
$spreadsheet_masa_studi2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_masa_studi_lulusan (F).xls');
$worksheet_masa_studi2 = $spreadsheet_masa_studi2->getActiveSheet();

// May be change
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

	$ts6_lulusan = 0; $ts5_lulusan = 0; $ts4_lulusan = 0; $ts3_lulusan = 0; $ts2_lulusan = 0; $ts1_lulusan = 0;
	
	$row_jumlah += 2;
	$baris_awal = $row_jumlah;
	$ts_lulusan = $worksheet_masa_studi3->getCell('I'.($row_jumlah))->getValue(); 
	
	for($row = $row_jumlah;$row <= ($highestRow_masa_studi3+1); $row++) {
		if($worksheet_masa_studi3->getCell('A'.$row)->getValue() == $worksheet_masa_studi3->getCell('A'.($row+1))->getValue()) {
			$ts4_lulusan += $worksheet_masa_studi3->getCell('E'.($row+1))->getValue();
			$ts3_lulusan += $worksheet_masa_studi3->getCell('F'.($row+1))->getValue();
			$ts2_lulusan += $worksheet_masa_studi3->getCell('G'.($row+1))->getValue();
			$ts1_lulusan += $worksheet_masa_studi3->getCell('H'.($row+1))->getValue();
			
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
 * Tabel 8.d.1 Waktu Tunggu Lulusan
 */

// Diploma
/* $spreadsheet_waktu_tunggu_diploma = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_waktu_tunggu_lulusan_diploma.xlsx');

$worksheet_waktu_tunggu_diploma = $spreadsheet_waktu_tunggu_diploma->getActiveSheet();

$highestRow_waktu_tunggu_diploma = $worksheet_waktu_tunggu_diploma->getHighestRow();

$worksheet_waktu_tunggu_diploma->setAutoFilter('B1:J'.$highestRow_waktu_tunggu_diploma);
$autoFilter_waktu_tunggu_diploma = $worksheet_waktu_tunggu_diploma->getAutoFilter();
$columnFilter_waktu_tunggu_diploma = $autoFilter_waktu_tunggu_diploma->getColumn('I');
$columnFilter_waktu_tunggu_diploma2 = $autoFilter_waktu_tunggu_diploma->getColumn('B');
$columnFilter_waktu_tunggu_diploma->setFilterType(
    \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_FILTER
);
$columnFilter_waktu_tunggu_diploma->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        $nama_prodi2
    );
$columnFilter_waktu_tunggu_diploma2->setFilterType(
    \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_FILTER
);
$columnFilter_waktu_tunggu_diploma2->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        'TS-2'
    );
$columnFilter_waktu_tunggu_diploma2->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        'TS-3'
    );
$columnFilter_waktu_tunggu_diploma2->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        'TS-4'
    );

$autoFilter_waktu_tunggu_diploma->showHideRows();

$array_waktu_tunggu_diploma = $worksheet_waktu_tunggu_diploma->toArray();
$data_waktu_tunggu_diploma = [];

foreach($worksheet_waktu_tunggu_diploma->getRowIterator() as $row_id => $row) {
    if($worksheet_waktu_tunggu_diploma->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['tahun_lulus'] = $array_waktu_tunggu_diploma[$row_id-1][1];
            $item['jml_lulusan'] = $array_waktu_tunggu_diploma[$row_id-1][2];
			$item['jml_terlacak'] = $array_waktu_tunggu_diploma[$row_id-1][3];
			$item['jml_dipesan'] = $array_waktu_tunggu_diploma[$row_id-1][4];
			$item['kurang_3'] = $array_waktu_tunggu_diploma[$row_id-1][5];
			$item['sampai_6'] = $array_waktu_tunggu_diploma[$row_id-1][6];
			$item['lebih_6'] = $array_waktu_tunggu_diploma[$row_id-1][7];
            $data_waktu_tunggu_diploma[] = $item;
        }
    }
}

$worksheet_waktu_tunggu_diploma2 = new \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet($spreadsheet_waktu_tunggu_diploma, 'Sheet 2');
$spreadsheet_waktu_tunggu_diploma->addSheet($worksheet_waktu_tunggu_diploma2);

$worksheet_waktu_tunggu_diploma2 = $spreadsheet_waktu_tunggu_diploma->getSheetByName('Sheet 2');
$worksheet_waktu_tunggu_diploma2->fromArray($data_waktu_tunggu_diploma, NULL, 'A1');

$writer_waktu_tunggu_diploma = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_waktu_tunggu_diploma, 'Xlsx');
$writer_waktu_tunggu_diploma->save('./formatted/sapto_waktu_tunggu_lulusan_diploma (F).xlsx');

$spreadsheet_waktu_tunggu_diploma->disconnectWorksheets();
unset($spreadsheet_waktu_tunggu_diploma);

// Load Format Baru
$spreadsheet_waktu_tunggu_diploma2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_waktu_tunggu_lulusan_diploma (F).xlsx');
$worksheet_waktu_tunggu_diploma3 = $spreadsheet_waktu_tunggu_diploma2->getSheetByName('Sheet 2');

// May be change
$waktu_tunggu_diploma_ts2 = $worksheet_waktu_tunggu_diploma3->rangeToArray('B1:G1', NULL, TRUE, TRUE, TRUE);
$waktu_tunggu_diploma_ts3 = $worksheet_waktu_tunggu_diploma3->rangeToArray('B2:G2', NULL, TRUE, TRUE, TRUE);
$waktu_tunggu_diploma_ts4 = $worksheet_waktu_tunggu_diploma3->rangeToArray('B3:G3', NULL, TRUE, TRUE, TRUE);

$spreadsheet_waktu_tunggu_diploma2->disconnectWorksheets();
unset($spreadsheet_waktu_tunggu_diploma2); */


// Sarjana
$spreadsheet_waktu_tunggu_sarjana = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_waktu_tunggu_lulusan_sarjana.xlsx');

$worksheet_waktu_tunggu_sarjana = $spreadsheet_waktu_tunggu_sarjana->getActiveSheet();

$highestRow_waktu_tunggu_sarjana = $worksheet_waktu_tunggu_sarjana->getHighestRow();

$worksheet_waktu_tunggu_sarjana->setAutoFilter('B1:I'.$highestRow_waktu_tunggu_sarjana);
$autoFilter_waktu_tunggu_sarjana = $worksheet_waktu_tunggu_sarjana->getAutoFilter();
$columnFilter_waktu_tunggu_sarjana = $autoFilter_waktu_tunggu_sarjana->getColumn('H');
$columnFilter_waktu_tunggu_sarjana2 = $autoFilter_waktu_tunggu_sarjana->getColumn('B');
$columnFilter_waktu_tunggu_sarjana->setFilterType(
    \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_FILTER
);
$columnFilter_waktu_tunggu_sarjana->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        $nama_prodi2
    );
$columnFilter_waktu_tunggu_sarjana2->setFilterType(
    \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_FILTER
);
$columnFilter_waktu_tunggu_sarjana2->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        'TS-2'
    );
$columnFilter_waktu_tunggu_sarjana2->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        'TS-3'
    );
$columnFilter_waktu_tunggu_sarjana2->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        'TS-4'
    );

$autoFilter_waktu_tunggu_sarjana->showHideRows();

$array_waktu_tunggu_sarjana = $worksheet_waktu_tunggu_sarjana->toArray();
$data_waktu_tunggu_sarjana = [];

foreach($worksheet_waktu_tunggu_sarjana->getRowIterator() as $row_id => $row) {
    if($worksheet_waktu_tunggu_sarjana->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['tahun_lulus'] = $array_waktu_tunggu_sarjana[$row_id-1][1];
            $item['jml_lulusan'] = $array_waktu_tunggu_sarjana[$row_id-1][2];
			$item['jml_terlacak'] = $array_waktu_tunggu_sarjana[$row_id-1][3];
			$item['kurang_6'] = $array_waktu_tunggu_sarjana[$row_id-1][4];
			$item['sampai_18'] = $array_waktu_tunggu_sarjana[$row_id-1][5];
			$item['lebih_18'] = $array_waktu_tunggu_sarjana[$row_id-1][6];
            $data_waktu_tunggu_sarjana[] = $item;
        }
    }
}

$worksheet_waktu_tunggu_sarjana2 = new \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet($spreadsheet_waktu_tunggu_sarjana, 'Sheet 2');
$spreadsheet_waktu_tunggu_sarjana->addSheet($worksheet_waktu_tunggu_sarjana2);

$worksheet_waktu_tunggu_sarjana2 = $spreadsheet_waktu_tunggu_sarjana->getSheetByName('Sheet 2');
$worksheet_waktu_tunggu_sarjana2->fromArray($data_waktu_tunggu_sarjana, NULL, 'A1');

$writer_waktu_tunggu_sarjana = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_waktu_tunggu_sarjana, 'Xlsx');
$writer_waktu_tunggu_sarjana->save('./formatted/sapto_waktu_tunggu_lulusan_sarjana (F).xlsx');

$spreadsheet_waktu_tunggu_sarjana->disconnectWorksheets();
unset($spreadsheet_waktu_tunggu_sarjana);

// Load Format Baru
$spreadsheet_waktu_tunggu_sarjana2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_waktu_tunggu_lulusan_sarjana (F).xlsx');
$worksheet_waktu_tunggu_sarjana3 = $spreadsheet_waktu_tunggu_sarjana2->getSheetByName('Sheet 2');

// May be change
$waktu_tunggu_sarjana_ts2 = $worksheet_waktu_tunggu_sarjana3->rangeToArray('B1:F1', NULL, TRUE, TRUE, TRUE);
$waktu_tunggu_sarjana_ts3 = $worksheet_waktu_tunggu_sarjana3->rangeToArray('B2:F2', NULL, TRUE, TRUE, TRUE);
$waktu_tunggu_sarjana_ts4 = $worksheet_waktu_tunggu_sarjana3->rangeToArray('B3:F3', NULL, TRUE, TRUE, TRUE);

$spreadsheet_waktu_tunggu_sarjana2->disconnectWorksheets();
unset($spreadsheet_waktu_tunggu_sarjana2);


// Sarjana Terapan
/* $spreadsheet_waktu_tunggu_terapan = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_waktu_tunggu_lulusan_terapan.xlsx');

$worksheet_waktu_tunggu_terapan = $spreadsheet_waktu_tunggu_terapan->getActiveSheet();

$highestRow_waktu_tunggu_terapan = $worksheet_waktu_tunggu_terapan->getHighestRow();

$worksheet_waktu_tunggu_terapan->setAutoFilter('B1:I'.$highestRow_waktu_tunggu_terapan);
$autoFilter_waktu_tunggu_terapan = $worksheet_waktu_tunggu_terapan->getAutoFilter();
$columnFilter_waktu_tunggu_terapan = $autoFilter_waktu_tunggu_terapan->getColumn('H');
$columnFilter_waktu_tunggu_terapan2 = $autoFilter_waktu_tunggu_terapan->getColumn('B');
$columnFilter_waktu_tunggu_terapan->setFilterType(
    \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_FILTER
);
$columnFilter_waktu_tunggu_terapan->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        $nama_prodi2
    );
$columnFilter_waktu_tunggu_terapan2->setFilterType(
    \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_FILTER
);
$columnFilter_waktu_tunggu_terapan2->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        'TS-4'
    );
$columnFilter_waktu_tunggu_terapan2->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        'TS-3'
    );
$columnFilter_waktu_tunggu_terapan2->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        'TS-2'
    );

$autoFilter_waktu_tunggu_terapan->showHideRows();

$array_waktu_tunggu_terapan = $worksheet_waktu_tunggu_terapan->toArray();
$data_waktu_tunggu_terapan = [];

foreach($worksheet_waktu_tunggu_terapan->getRowIterator() as $row_id => $row) {
    if($worksheet_waktu_tunggu_terapan->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['tahun_lulus'] = $array_waktu_tunggu_terapan[$row_id-1][1];
            $item['jml_lulusan'] = $array_waktu_tunggu_terapan[$row_id-1][2];
			$item['jml_terlacak'] = $array_waktu_tunggu_terapan[$row_id-1][3];
			$item['kurang_3'] = $array_waktu_tunggu_terapan[$row_id-1][4];
			$item['sampai_6'] = $array_waktu_tunggu_terapan[$row_id-1][5];
			$item['lebih_6'] = $array_waktu_tunggu_terapan[$row_id-1][6];
            $data_waktu_tunggu_terapan[] = $item;
        }
    }
}

$worksheet_waktu_tunggu_terapan2 = new \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet($spreadsheet_waktu_tunggu_terapan, 'Sheet 2');
$spreadsheet_waktu_tunggu_terapan->addSheet($worksheet_waktu_tunggu_terapan2);

$worksheet_waktu_tunggu_terapan2 = $spreadsheet_waktu_tunggu_terapan->getSheetByName('Sheet 2');
$worksheet_waktu_tunggu_terapan2->fromArray($data_waktu_tunggu_terapan, NULL, 'A1');

$writer_waktu_tunggu_terapan = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_waktu_tunggu_terapan, 'Xlsx');
$writer_waktu_tunggu_terapan->save('./formatted/sapto_waktu_tunggu_lulusan_terapan (F).xlsx');

$spreadsheet_waktu_tunggu_terapan->disconnectWorksheets();
unset($spreadsheet_waktu_tunggu_terapan);

// Load Format Baru
$spreadsheet_waktu_tunggu_terapan2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_waktu_tunggu_lulusan_terapan (F).xlsx');
$worksheet_waktu_tunggu_terapan3 = $spreadsheet_waktu_tunggu_terapan2->getSheetByName('Sheet 2');

// May be change
$waktu_tunggu_terapan_ts2 = $worksheet_waktu_tunggu_terapan3->rangeToArray('B1:F1', NULL, TRUE, TRUE, TRUE);
$waktu_tunggu_terapan_ts3 = $worksheet_waktu_tunggu_terapan3->rangeToArray('B2:F2', NULL, TRUE, TRUE, TRUE);
$waktu_tunggu_terapan_ts4 = $worksheet_waktu_tunggu_terapan3->rangeToArray('B3:F3', NULL, TRUE, TRUE, TRUE);

$spreadsheet_waktu_tunggu_terapan2->disconnectWorksheets();
unset($spreadsheet_waktu_tunggu_terapan2); */


/**
 * Tabel 8.d.2 Kesesuaian Bidang Kerja Lulusan
 */

$spreadsheet_sesuai_kerja = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_kesesuaian_kerja_lulusan.xlsx');

$worksheet_sesuai_kerja = $spreadsheet_sesuai_kerja->getActiveSheet();

$highestRow_sesuai_kerja = $worksheet_sesuai_kerja->getHighestRow();

$worksheet_sesuai_kerja->setAutoFilter('B1:I'.$highestRow_sesuai_kerja);
$autoFilter_sesuai_kerja = $worksheet_sesuai_kerja->getAutoFilter();
$columnFilter_sesuai_kerja = $autoFilter_sesuai_kerja->getColumn('H');
$columnFilter_sesuai_kerja->setFilterType(
    \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_FILTER
);
$columnFilter_sesuai_kerja->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        $nama_prodi2
    );

$autoFilter_sesuai_kerja->showHideRows();

$writer_sesuai_kerja = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_sesuai_kerja, 'Xlsx');
$writer_sesuai_kerja->save('./formatted/sapto_kesesuaian_kerja_lulusan (F).xlsx');

$spreadsheet_sesuai_kerja->disconnectWorksheets();
unset($spreadsheet_sesuai_kerja);

// Load Format Baru
$spreadsheet_sesuai_kerja2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_kesesuaian_kerja_lulusan (F).xlsx');
$worksheet_sesuai_kerja2 = $spreadsheet_sesuai_kerja2->getActiveSheet();

// May be change
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

$spreadsheet_tempat_kerja = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_tempat_kerja_lulusan.xlsx');

$worksheet_tempat_kerja = $spreadsheet_tempat_kerja->getActiveSheet();

$highestRow_tempat_kerja = $worksheet_tempat_kerja->getHighestRow();

$worksheet_tempat_kerja->setAutoFilter('B1:I'.$highestRow_tempat_kerja);
$autoFilter_tempat_kerja = $worksheet_tempat_kerja->getAutoFilter();
$columnFilter_tempat_kerja = $autoFilter_tempat_kerja->getColumn('H');
$columnFilter_tempat_kerja->setFilterType(
    \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_FILTER
);
$columnFilter_tempat_kerja->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        $nama_prodi2
    );

$autoFilter_tempat_kerja->showHideRows();

$writer_tempat_kerja = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_tempat_kerja, 'Xlsx');
$writer_tempat_kerja->save('./formatted/sapto_tempat_kerja_lulusan (F).xlsx');

$spreadsheet_tempat_kerja->disconnectWorksheets();
unset($spreadsheet_tempat_kerja);

// Load Format Baru
$spreadsheet_tempat_kerja2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_tempat_kerja_lulusan (F).xlsx');
$worksheet_tempat_kerja2 = $spreadsheet_tempat_kerja2->getActiveSheet();

// May be change
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
 * Tabel 8.e.2 Kepuasan Pengguna Lulusan
 */

$spreadsheet_kepuasan_lulusan = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_kepuasan_pengguna_lulusan.xlsx');

$worksheet_kepuasan_lulusan = $spreadsheet_kepuasan_lulusan->getActiveSheet();

$highestRow_kepuasan_lulusan = $worksheet_kepuasan_lulusan->getHighestRow();

$worksheet_kepuasan_lulusan->setAutoFilter('B1:I'.$highestRow_kepuasan_lulusan);
$autoFilter_kepuasan_lulusan = $worksheet_kepuasan_lulusan->getAutoFilter();
$columnFilter_kepuasan_lulusan = $autoFilter_kepuasan_lulusan->getColumn('H');
$columnFilter_kepuasan_lulusan->setFilterType(
    \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_FILTER
);
$columnFilter_kepuasan_lulusan->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        $nama_prodi2
    );

$autoFilter_kepuasan_lulusan->showHideRows();

$writer_kepuasan_lulusan = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_kepuasan_lulusan, 'Xlsx');
$writer_kepuasan_lulusan->save('./formatted/sapto_kepuasan_pengguna_lulusan (F).xlsx');

$spreadsheet_kepuasan_lulusan->disconnectWorksheets();
unset($spreadsheet_kepuasan_lulusan);

// Load Format Baru
$spreadsheet_kepuasan_lulusan2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_kepuasan_pengguna_lulusan (F).xlsx');
$worksheet_kepuasan_lulusan2 = $spreadsheet_kepuasan_lulusan2->getActiveSheet();

// May be change
$array_kepuasan_lulusan = $worksheet_kepuasan_lulusan2->toArray();
$data_kepuasan_lulusan = [];

foreach($worksheet_kepuasan_lulusan2->getRowIterator() as $row_id => $row) {
    if($worksheet_kepuasan_lulusan2->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['jenis_kemampuan'] = $array_kepuasan_lulusan[$row_id-1][1];
            $item['puas_sangat_baik'] = $array_kepuasan_lulusan[$row_id-1][2];
			$item['puas_baik'] = $array_kepuasan_lulusan[$row_id-1][3];
			$item['puas_cukup'] = $array_kepuasan_lulusan[$row_id-1][4];
			$item['puas_kurang'] = $array_kepuasan_lulusan[$row_id-1][5];
			$item['rencana_lanjut'] = $array_kepuasan_lulusan[$row_id-1][6];
            $data_kepuasan_lulusan[] = $item;
        }
    }
}

$spreadsheet_kepuasan_lulusan2->disconnectWorksheets();
unset($spreadsheet_kepuasan_lulusan2);


/**
 * Tabel 8.f.1 Publikasi Ilmiah Mahasiswa
 */

$spreadsheet_publikasi_ilmiah_mhs = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_publikasi_ilmiah_mhs.xlsx');

$worksheet_publikasi_ilmiah_mhs = $spreadsheet_publikasi_ilmiah_mhs->getActiveSheet();

$highestRow_publikasi_ilmiah_mhs = $worksheet_publikasi_ilmiah_mhs->getHighestRow();

$worksheet_publikasi_ilmiah_mhs->setAutoFilter('B1:H'.$highestRow_publikasi_ilmiah_mhs);
$autoFilter_publikasi_ilmiah_mhs = $worksheet_publikasi_ilmiah_mhs->getAutoFilter();
$columnFilter_publikasi_ilmiah_mhs = $autoFilter_publikasi_ilmiah_mhs->getColumn('G');
$columnFilter_publikasi_ilmiah_mhs->setFilterType(
    \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_FILTER
);
$columnFilter_publikasi_ilmiah_mhs->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        $nama_prodi2
    );

$autoFilter_publikasi_ilmiah_mhs->showHideRows();

$writer_publikasi_ilmiah_mhs = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_publikasi_ilmiah_mhs, 'Xlsx');
$writer_publikasi_ilmiah_mhs->save('./formatted/sapto_publikasi_ilmiah_mhs (F).xlsx');

$spreadsheet_publikasi_ilmiah_mhs->disconnectWorksheets();
unset($spreadsheet_publikasi_ilmiah_mhs);

// Load Format Baru
$spreadsheet_publikasi_ilmiah_mhs2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_publikasi_ilmiah_mhs (F).xlsx');
$worksheet_publikasi_ilmiah_mhs2 = $spreadsheet_publikasi_ilmiah_mhs2->getActiveSheet();

// May be change
$array_publikasi_ilmiah_mhs = $worksheet_publikasi_ilmiah_mhs2->toArray();
$data_publikasi_ilmiah_mhs = [];

foreach($worksheet_publikasi_ilmiah_mhs2->getRowIterator() as $row_id => $row) {
    if($worksheet_publikasi_ilmiah_mhs2->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['jenis_publikasi'] = $array_publikasi_ilmiah_mhs[$row_id-1][1];
            $item['ts2_judul'] = $array_publikasi_ilmiah_mhs[$row_id-1][2];
			$item['ts1_judul'] = $array_publikasi_ilmiah_mhs[$row_id-1][3];
			$item['ts_judul'] = $array_publikasi_ilmiah_mhs[$row_id-1][4];
			$item['jumlah'] = $array_publikasi_ilmiah_mhs[$row_id-1][5];
            $data_publikasi_ilmiah_mhs[] = $item;
        }
    }
}

$spreadsheet_publikasi_ilmiah_mhs2->disconnectWorksheets();
unset($spreadsheet_publikasi_ilmiah_mhs2);


/**
 * Tabel 8.f.1 Pagelaran/Pameran/Presentasi/Publikasi Ilmiah Mahasiswa
 */

$spreadsheet_pagelaran_mhs = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_pagelaran_presentasi_publikasi_mhs.xlsx');

$worksheet_pagelaran_mhs = $spreadsheet_pagelaran_mhs->getActiveSheet();

$highestRow_pagelaran_mhs = $worksheet_pagelaran_mhs->getHighestRow();

$worksheet_pagelaran_mhs->setAutoFilter('B1:H'.$highestRow_pagelaran_mhs);
$autoFilter_pagelaran_mhs = $worksheet_pagelaran_mhs->getAutoFilter();
$columnFilter_pagelaran_mhs = $autoFilter_pagelaran_mhs->getColumn('G');
$columnFilter_pagelaran_mhs->setFilterType(
    \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_FILTER
);
$columnFilter_pagelaran_mhs->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        $nama_prodi2
    );

$autoFilter_pagelaran_mhs->showHideRows();

$writer_pagelaran_mhs = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_pagelaran_mhs, 'Xlsx');
$writer_pagelaran_mhs->save('./formatted/sapto_pagelaran_presentasi_publikasi_mhs (F).xlsx');

$spreadsheet_pagelaran_mhs->disconnectWorksheets();
unset($spreadsheet_pagelaran_mhs);

// Load Format Baru
$spreadsheet_pagelaran_mhs2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_pagelaran_presentasi_publikasi_mhs (F).xlsx');
$worksheet_pagelaran_mhs2 = $spreadsheet_pagelaran_mhs2->getActiveSheet();

// May be change
$array_pagelaran_mhs = $worksheet_pagelaran_mhs2->toArray();
$data_pagelaran_mhs = [];

foreach($worksheet_pagelaran_mhs2->getRowIterator() as $row_id => $row) {
    if($worksheet_pagelaran_mhs2->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['jenis_publikasi'] = $array_pagelaran_mhs[$row_id-1][1];
            $item['ts2_judul'] = $array_pagelaran_mhs[$row_id-1][2];
			$item['ts1_judul'] = $array_pagelaran_mhs[$row_id-1][3];
			$item['ts_judul'] = $array_pagelaran_mhs[$row_id-1][4];
			$item['jumlah'] = $array_pagelaran_mhs[$row_id-1][5];
            $data_pagelaran_mhs[] = $item;
        }
    }
}

$spreadsheet_pagelaran_mhs2->disconnectWorksheets();
unset($spreadsheet_pagelaran_mhs2);


/**
 * Tabel 8.f.2 Karya Ilmiah Mahasiswa yang Disitasi
 */

$spreadsheet_karya_mhs_disitasi = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_karya_mhs_disitasi.xlsx');

$worksheet_karya_mhs_disitasi = $spreadsheet_karya_mhs_disitasi->getActiveSheet();

$highestRow_karya_mhs_disitasi = $worksheet_karya_mhs_disitasi->getHighestRow();

$worksheet_karya_mhs_disitasi->setAutoFilter('B1:F'.$highestRow_karya_mhs_disitasi);
$autoFilter_karya_mhs_disitasi = $worksheet_karya_mhs_disitasi->getAutoFilter();
$columnFilter_karya_mhs_disitasi = $autoFilter_karya_mhs_disitasi->getColumn('E');
$columnFilter_karya_mhs_disitasi->setFilterType(
    \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_FILTER
);
$columnFilter_karya_mhs_disitasi->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        $nama_prodi2
    );

$autoFilter_karya_mhs_disitasi->showHideRows();

$writer_karya_mhs_disitasi = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_karya_mhs_disitasi, 'Xlsx');
$writer_karya_mhs_disitasi->save('./formatted/sapto_karya_mhs_disitasi (F).xlsx');

$spreadsheet_karya_mhs_disitasi->disconnectWorksheets();
unset($spreadsheet_karya_mhs_disitasi);

// Load Format Baru
$spreadsheet_karya_mhs_disitasi2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_karya_mhs_disitasi (F).xlsx');
$worksheet_karya_mhs_disitasi2 = $spreadsheet_karya_mhs_disitasi2->getActiveSheet();

// May be change
$array_karya_mhs_disitasi = $worksheet_karya_mhs_disitasi2->toArray();
$data_karya_mhs_disitasi = [];

foreach($worksheet_karya_mhs_disitasi2->getRowIterator() as $row_id => $row) {
    if($worksheet_karya_mhs_disitasi2->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['nama_mhs'] = $array_karya_mhs_disitasi[$row_id-1][1];
            $item['judul_artikel'] = $array_karya_mhs_disitasi[$row_id-1][2];
			$item['jml_sitasi'] = $array_karya_mhs_disitasi[$row_id-1][3];
            $data_karya_mhs_disitasi[] = $item;
        }
    }
}

$spreadsheet_karya_mhs_disitasi2->disconnectWorksheets();
unset($spreadsheet_karya_mhs_disitasi2);


/**
 * Tabel 8.f.3 Produk/Jasa Mahasiswa yang Diadopsi oleh Industri/Masyarakat
 */

/* $spreadsheet_produk_mhs_diadopsi = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_produk_jasa_mhs_masyarakat.xlsx');

$worksheet_produk_mhs_diadopsi = $spreadsheet_produk_mhs_diadopsi->getActiveSheet();

$highestRow_produk_mhs_diadopsi = $worksheet_produk_mhs_diadopsi->getHighestRow();

$worksheet_produk_mhs_diadopsi->setAutoFilter('B1:G'.$highestRow_karya_mhs_disitasi);
$autoFilter_produk_mhs_diadopsi = $worksheet_produk_mhs_diadopsi->getAutoFilter();
$columnFilter_produk_mhs_diadopsi = $autoFilter_produk_mhs_diadopsi->getColumn('F');
$columnFilter_produk_mhs_diadopsi->setFilterType(
    \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_FILTER
);
$columnFilter_produk_mhs_diadopsi->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        $nama_prodi2
    );

$autoFilter_produk_mhs_diadopsi->showHideRows();

$writer_produk_mhs_diadopsi = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_produk_mhs_diadopsi, 'Xlsx');
$writer_produk_mhs_diadopsi->save('./formatted/sapto_produk_jasa_mhs_masyarakat (F).xlsx');

$spreadsheet_produk_mhs_diadopsi->disconnectWorksheets();
unset($spreadsheet_produk_mhs_diadopsi);

// Load Format Baru
$spreadsheet_produk_mhs_diadopsi2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_produk_jasa_mhs_masyarakat (F).xlsx');
$worksheet_produk_mhs_diadopsi2 = $spreadsheet_produk_mhs_diadopsi2->getActiveSheet();

// May be change
$array_produk_mhs_diadopsi = $worksheet_produk_mhs_diadopsi2->toArray();
$data_produk_mhs_diadopsi = [];

foreach($worksheet_produk_mhs_diadopsi2->getRowIterator() as $row_id => $row) {
    if($worksheet_produk_mhs_diadopsi2->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['nama_mhs'] = $array_produk_mhs_diadopsi[$row_id-1][1];
            $item['nama_produk'] = $array_produk_mhs_diadopsi[$row_id-1][2];
			$item['desk_produk'] = $array_produk_mhs_diadopsi[$row_id-1][3];
			$item['bukti'] = $array_produk_mhs_diadopsi[$row_id-1][4];
            $data_produk_mhs_diadopsi[] = $item;
        }
    }
}

$spreadsheet_produk_mhs_diadopsi2->disconnectWorksheets();
unset($spreadsheet_produk_mhs_diadopsi2); */


/**
 * Tabel 8.f.4 Luaran Penelitian yang Dihasilkan Mahasiswa - HKI (Paten, Paten Sederhana)
 */

$spreadsheet_luaran_penelitian_mhs_1 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_luaran_penelitian_mhs.xlsx');

$worksheet_luaran_penelitian_mhs_1 = $spreadsheet_luaran_penelitian_mhs_1->getActiveSheet();

$highestRow_luaran_penelitian_mhs_1 = $worksheet_luaran_penelitian_mhs_1->getHighestRow();

$worksheet_luaran_penelitian_mhs_1->setAutoFilter('B1:E'.$highestRow_luaran_penelitian_mhs_1);
$autoFilter_luaran_penelitian_mhs_1 = $worksheet_luaran_penelitian_mhs_1->getAutoFilter();
$columnFilter_luaran_penelitian_mhs_1 = $autoFilter_luaran_penelitian_mhs_1->getColumn('E');
$columnFilter_luaran_penelitian_mhs_1->setFilterType(
    \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_FILTER
);
$columnFilter_luaran_penelitian_mhs_1->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        $nama_prodi2
    );

$autoFilter_luaran_penelitian_mhs_1->showHideRows();

$writer_luaran_penelitian_mhs_1 = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_luaran_penelitian_mhs_1, 'Xlsx');
$writer_luaran_penelitian_mhs_1->save('./formatted/sapto_luaran_penelitian_mhs (F).xlsx');

$spreadsheet_luaran_penelitian_mhs_1->disconnectWorksheets();
unset($spreadsheet_luaran_penelitian_mhs_1);

// Load Format Baru
$spreadsheet_luaran_penelitian_mhs_12 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_luaran_penelitian_mhs (F).xlsx');
$worksheet_luaran_penelitian_mhs_12 = $spreadsheet_luaran_penelitian_mhs_12->getActiveSheet();

// May be change
$array_luaran_penelitian_mhs_1 = $worksheet_luaran_penelitian_mhs_12->toArray();
$data_luaran_penelitian_mhs_1 = [];

foreach($worksheet_luaran_penelitian_mhs_12->getRowIterator() as $row_id => $row) {
    if($worksheet_luaran_penelitian_mhs_12->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['luaran_penelitian'] = $array_luaran_penelitian_mhs_1[$row_id-1][1];
            $item['tahun'] = $array_luaran_penelitian_mhs_1[$row_id-1][2];
			$item['keterangan'] = $array_luaran_penelitian_mhs_1[$row_id-1][3];
            $data_luaran_penelitian_mhs_1[] = $item;
        }
    }
}

$spreadsheet_luaran_penelitian_mhs_12->disconnectWorksheets();
unset($spreadsheet_luaran_penelitian_mhs_12);


/**
 * Tabel 8.f.4 Luaran Penelitian yang Dihasilkan Mahasiswa - HKI (Hak Cipta, Desain Produk Industri, dll.)
 */

$spreadsheet_luaran_penelitian_mhs_2 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_luaran_penelitian_mhs_2.xlsx');

$worksheet_luaran_penelitian_mhs_2 = $spreadsheet_luaran_penelitian_mhs_2->getActiveSheet();

$highestRow_luaran_penelitian_mhs_2 = $worksheet_luaran_penelitian_mhs_2->getHighestRow();

$worksheet_luaran_penelitian_mhs_2->setAutoFilter('B1:E'.$highestRow_luaran_penelitian_mhs_2);
$autoFilter_luaran_penelitian_mhs_2 = $worksheet_luaran_penelitian_mhs_2->getAutoFilter();
$columnFilter_luaran_penelitian_mhs_2 = $autoFilter_luaran_penelitian_mhs_2->getColumn('E');
$columnFilter_luaran_penelitian_mhs_2->setFilterType(
    \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_FILTER
);
$columnFilter_luaran_penelitian_mhs_2->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        $nama_prodi2
    );

$autoFilter_luaran_penelitian_mhs_2->showHideRows();

$writer_luaran_penelitian_mhs_2 = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_luaran_penelitian_mhs_2, 'Xlsx');
$writer_luaran_penelitian_mhs_2->save('./formatted/sapto_luaran_penelitian_mhs_2 (F).xlsx');

$spreadsheet_luaran_penelitian_mhs_2->disconnectWorksheets();
unset($spreadsheet_luaran_penelitian_mhs_2);

// Load Format Baru
$spreadsheet_luaran_penelitian_mhs_22 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_luaran_penelitian_mhs_2 (F).xlsx');
$worksheet_luaran_penelitian_mhs_22 = $spreadsheet_luaran_penelitian_mhs_22->getActiveSheet();

// May be change
$array_luaran_penelitian_mhs_2 = $worksheet_luaran_penelitian_mhs_22->toArray();
$data_luaran_penelitian_mhs_2 = [];

foreach($worksheet_luaran_penelitian_mhs_22->getRowIterator() as $row_id => $row) {
    if($worksheet_luaran_penelitian_mhs_22->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['luaran_penelitian'] = $array_luaran_penelitian_mhs_2[$row_id-1][1];
            $item['tahun'] = $array_luaran_penelitian_mhs_2[$row_id-1][2];
			$item['keterangan'] = $array_luaran_penelitian_mhs_2[$row_id-1][3];
            $data_luaran_penelitian_mhs_2[] = $item;
        }
    }
}

$spreadsheet_luaran_penelitian_mhs_22->disconnectWorksheets();
unset($spreadsheet_luaran_penelitian_mhs_22);


/**
 * Tabel 8.f.4 Luaran Penelitian yang Dihasilkan Mahasiswa - Teknologi Tepat Guna, Produk, Karya Seni, Rekayasa Sosial
 */

$spreadsheet_luaran_penelitian_mhs_3 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_luaran_penelitian_mhs_3.xlsx');

$worksheet_luaran_penelitian_mhs_3 = $spreadsheet_luaran_penelitian_mhs_3->getActiveSheet();

$highestRow_luaran_penelitian_mhs_3 = $worksheet_luaran_penelitian_mhs_3->getHighestRow();

$worksheet_luaran_penelitian_mhs_3->setAutoFilter('B1:E'.$highestRow_luaran_penelitian_mhs_3);
$autoFilter_luaran_penelitian_mhs_3 = $worksheet_luaran_penelitian_mhs_3->getAutoFilter();
$columnFilter_luaran_penelitian_mhs_3 = $autoFilter_luaran_penelitian_mhs_3->getColumn('E');
$columnFilter_luaran_penelitian_mhs_3->setFilterType(
    \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_FILTER
);
$columnFilter_luaran_penelitian_mhs_3->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        $nama_prodi2
    );

$autoFilter_luaran_penelitian_mhs_3->showHideRows();

$writer_luaran_penelitian_mhs_3 = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_luaran_penelitian_mhs_3, 'Xlsx');
$writer_luaran_penelitian_mhs_3->save('./formatted/sapto_luaran_penelitian_mhs_3 (F).xlsx');

$spreadsheet_luaran_penelitian_mhs_3->disconnectWorksheets();
unset($spreadsheet_luaran_penelitian_mhs_3);

// Load Format Baru
$spreadsheet_luaran_penelitian_mhs_32 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_luaran_penelitian_mhs_3 (F).xlsx');
$worksheet_luaran_penelitian_mhs_32 = $spreadsheet_luaran_penelitian_mhs_32->getActiveSheet();

// May be change
$array_luaran_penelitian_mhs_3 = $worksheet_luaran_penelitian_mhs_32->toArray();
$data_luaran_penelitian_mhs_3 = [];

foreach($worksheet_luaran_penelitian_mhs_32->getRowIterator() as $row_id => $row) {
    if($worksheet_luaran_penelitian_mhs_32->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['luaran_penelitian'] = $array_luaran_penelitian_mhs_3[$row_id-1][1];
            $item['tahun'] = $array_luaran_penelitian_mhs_3[$row_id-1][2];
			$item['keterangan'] = $array_luaran_penelitian_mhs_3[$row_id-1][3];
            $data_luaran_penelitian_mhs_3[] = $item;
        }
    }
}

$spreadsheet_luaran_penelitian_mhs_32->disconnectWorksheets();
unset($spreadsheet_luaran_penelitian_mhs_32);


/**
 * Tabel 8.f.4 Luaran Penelitian yang Dihasilkan Mahasiswa - Buku ber-ISBN, Book Chapter
 */

$spreadsheet_luaran_penelitian_mhs_4 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_luaran_penelitian_mhs_4.xlsx');

$worksheet_luaran_penelitian_mhs_4 = $spreadsheet_luaran_penelitian_mhs_4->getActiveSheet();

$highestRow_luaran_penelitian_mhs_4 = $worksheet_luaran_penelitian_mhs_4->getHighestRow();

$worksheet_luaran_penelitian_mhs_4->setAutoFilter('B1:E'.$highestRow_luaran_penelitian_mhs_4);
$autoFilter_luaran_penelitian_mhs_4 = $worksheet_luaran_penelitian_mhs_4->getAutoFilter();
$columnFilter_luaran_penelitian_mhs_4 = $autoFilter_luaran_penelitian_mhs_4->getColumn('E');
$columnFilter_luaran_penelitian_mhs_4->setFilterType(
    \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_FILTER
);
$columnFilter_luaran_penelitian_mhs_4->createRule()
    ->setRule(
        \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
        $nama_prodi2
    );

$autoFilter_luaran_penelitian_mhs_4->showHideRows();

$writer_luaran_penelitian_mhs_4 = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_luaran_penelitian_mhs_4, 'Xlsx');
$writer_luaran_penelitian_mhs_4->save('./formatted/sapto_luaran_penelitian_mhs_4 (F).xlsx');

$spreadsheet_luaran_penelitian_mhs_4->disconnectWorksheets();
unset($spreadsheet_luaran_penelitian_mhs_4);

// Load Format Baru
$spreadsheet_luaran_penelitian_mhs_42 = \PhpOffice\PhpSpreadsheet\IOFactory::load('./formatted/sapto_luaran_penelitian_mhs_4 (F).xlsx');
$worksheet_luaran_penelitian_mhs_42 = $spreadsheet_luaran_penelitian_mhs_42->getActiveSheet();

// May be change
$array_luaran_penelitian_mhs_4 = $worksheet_luaran_penelitian_mhs_42->toArray();
$data_luaran_penelitian_mhs_4 = [];

foreach($worksheet_luaran_penelitian_mhs_42->getRowIterator() as $row_id => $row) {
    if($worksheet_luaran_penelitian_mhs_42->getRowDimension($row_id)->getVisible()) {
        if($row_id > 1) { 
            $item = array();
            $item['luaran_penelitian'] = $array_luaran_penelitian_mhs_4[$row_id-1][1];
            $item['tahun'] = $array_luaran_penelitian_mhs_4[$row_id-1][2];
			$item['keterangan'] = $array_luaran_penelitian_mhs_4[$row_id-1][3];
            $data_luaran_penelitian_mhs_4[] = $item;
        }
    }
}

$spreadsheet_luaran_penelitian_mhs_42->disconnectWorksheets();
unset($spreadsheet_luaran_penelitian_mhs_42);




/**
 * Set ke Template SAPTO
 */

$spreadsheet_aps = \PhpOffice\PhpSpreadsheet\IOFactory::load('./raw/sapto_aps9.xlsx');


// Kerjasama Tridharma Pendidikan
$worksheet_aps = $spreadsheet_aps->getSheetByName('1-1');
$worksheet_aps->fromArray($data_tridharma_pendidikan, NULL, 'B12');

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
$worksheet_aps2->fromArray($data_tridharma_penelitian, NULL, 'B12');

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
$worksheet_aps3->fromArray($data_tridharma_pkm, NULL, 'B12');

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

// Seleksi Mahasiswa Baru
$worksheet_aps4 = $spreadsheet_aps->getSheetByName('2a');
$worksheet_aps4->fromArray($ts4_seleksi_mhs, NULL, 'B6');
$worksheet_aps4->fromArray($ts3_seleksi_mhs, NULL, 'B7');
$worksheet_aps4->fromArray($ts2_seleksi_mhs, NULL, 'B8');
$worksheet_aps4->fromArray($ts1_seleksi_mhs, NULL, 'B9');
$worksheet_aps4->fromArray($ts_seleksi_mhs, NULL, 'B10');

$worksheet_aps4->setCellValue('C11', '=SUM(C6:C10)');
$worksheet_aps4->setCellValue('D11', '=SUM(D6:D10)');
$worksheet_aps4->setCellValue('E11', '=SUM(E6:E10)');
$worksheet_aps4->setCellValue('F11', '=SUM(F6:F10)');
$worksheet_aps4->setCellValue('G11', '=SUM(G6:H10)');


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
$worksheet_aps9->fromArray($data_penelitian_dtps, NULL, 'B6');

$highestRow_aps9 = $worksheet_aps9->getHighestRow();
$jumlahRow_aps9 = $highestRow_aps9 + 1;

$worksheet_aps9->setCellValue('A'.$jumlahRow_aps9, 'Jumlah');
$worksheet_aps9->setCellValue('C'.$jumlahRow_aps9, '=SUM(C6:C'.$highestRow_aps9.')');
$worksheet_aps9->setCellValue('D'.$jumlahRow_aps9, '=SUM(D6:D'.$highestRow_aps9.')');
$worksheet_aps9->setCellValue('E'.$jumlahRow_aps9, '=SUM(E6:E'.$highestRow_aps9.')');
$worksheet_aps9->setCellValue('F'.$jumlahRow_aps9, '=SUM(C'.$jumlahRow_aps9.':E'.$jumlahRow_aps9.')');

$worksheet_aps9->mergeCells('A'.$jumlahRow_aps9.':B'.$jumlahRow_aps9);
$worksheet_aps9->getStyle('A6:F'.$jumlahRow_aps9)->applyFromArray($styleBorder);
$worksheet_aps9->getStyle('C6:E'.$highestRow_aps9)->applyFromArray($styleYellow);
$worksheet_aps9->getStyle('A6:A'.$jumlahRow_aps9)->applyFromArray($styleCenter);
$worksheet_aps9->getStyle('C6:F'.$jumlahRow_aps9)->applyFromArray($styleCenter);
$worksheet_aps9->getStyle('A'.$jumlahRow_aps9.':F'.$jumlahRow_aps9)->applyFromArray($styleBold);
$worksheet_aps9->getStyle('B6:B'.$highestRow_aps9)->getAlignment()->setWrapText(true);

foreach($worksheet_aps9->getRowDimensions() as $rd9) { 
    $rd9->setRowHeight(-1); 
}

for($row = 6; $row <= $highestRow_aps9; $row++) {
	$worksheet_aps9->setCellValue('A'.$row, $row-5);
}


// PkM DTPS
$worksheet_aps10 = $spreadsheet_aps->getSheetByName('3b3');
$worksheet_aps10->fromArray($data_pkm_dtps, NULL, 'B6');

$highestRow_aps10 = $worksheet_aps10->getHighestRow();
$jumlahRow_aps10 = $highestRow_aps10 + 1;

$worksheet_aps10->setCellValue('A'.$jumlahRow_aps10, 'Jumlah');
$worksheet_aps10->setCellValue('C'.$jumlahRow_aps10, '=SUM(C6:C'.$highestRow_aps10.')');
$worksheet_aps10->setCellValue('D'.$jumlahRow_aps10, '=SUM(D6:D'.$highestRow_aps10.')');
$worksheet_aps10->setCellValue('E'.$jumlahRow_aps10, '=SUM(E6:E'.$highestRow_aps10.')');
$worksheet_aps10->setCellValue('F'.$jumlahRow_aps10, '=SUM(C'.$jumlahRow_aps10.':E'.$jumlahRow_aps10.')');

$worksheet_aps10->mergeCells('A'.$jumlahRow_aps10.':B'.$jumlahRow_aps10);
$worksheet_aps10->getStyle('A6:F'.$jumlahRow_aps10)->applyFromArray($styleBorder);
$worksheet_aps10->getStyle('C6:E'.$highestRow_aps10)->applyFromArray($styleYellow);
$worksheet_aps10->getStyle('A6:A'.$jumlahRow_aps10)->applyFromArray($styleCenter);
$worksheet_aps10->getStyle('C6:F'.$jumlahRow_aps10)->applyFromArray($styleCenter);
$worksheet_aps10->getStyle('A'.$jumlahRow_aps10.':F'.$jumlahRow_aps10)->applyFromArray($styleBold);
$worksheet_aps10->getStyle('B6:B'.$highestRow_aps10)->getAlignment()->setWrapText(true);

foreach($worksheet_aps10->getRowDimensions() as $rd10) { 
    $rd10->setRowHeight(-1); 
}

for($row = 6; $row <= $highestRow_aps10; $row++) {
	$worksheet_aps10->setCellValue('A'.$row, $row-5);
}


// Publikasi Ilmiah DTPS
$worksheet_aps11 = $spreadsheet_aps->getSheetByName('3b4-1');
$worksheet_aps11->fromArray($data_publikasi_ilmiah_dtps, NULL, 'B7');

$highestRow_aps11 = $worksheet_aps11->getHighestRow();
$jumlahRow_aps11 = $highestRow_aps11 + 1;

$worksheet_aps11->setCellValue('A'.$jumlahRow_aps11, 'Jumlah');
$worksheet_aps11->setCellValue('C'.$jumlahRow_aps11, '=SUM(C7:C'.$highestRow_aps11.')');
$worksheet_aps11->setCellValue('D'.$jumlahRow_aps11, '=SUM(D7:D'.$highestRow_aps11.')');
$worksheet_aps11->setCellValue('E'.$jumlahRow_aps11, '=SUM(E7:E'.$highestRow_aps11.')');
$worksheet_aps11->setCellValue('F'.$jumlahRow_aps11, '=SUM(C'.$jumlahRow_aps11.':E'.$jumlahRow_aps11.')');

$worksheet_aps11->mergeCells('A'.$jumlahRow_aps11.':B'.$jumlahRow_aps11);
$worksheet_aps11->getStyle('A7:F'.$jumlahRow_aps11)->applyFromArray($styleBorder);
$worksheet_aps11->getStyle('C7:E'.$highestRow_aps11)->applyFromArray($styleYellow);
$worksheet_aps11->getStyle('A7:A'.$jumlahRow_aps11)->applyFromArray($styleCenter);
$worksheet_aps11->getStyle('C7:F'.$jumlahRow_aps11)->applyFromArray($styleCenter);
$worksheet_aps11->getStyle('A'.$jumlahRow_aps11.':F'.$jumlahRow_aps11)->applyFromArray($styleBold);
$worksheet_aps11->getStyle('B7:B'.$highestRow_aps11)->getAlignment()->setWrapText(true);

foreach($worksheet_aps11->getRowDimensions() as $rd11) { 
    $rd11->setRowHeight(-1); 
}

for($row = 7; $row <= $highestRow_aps11; $row++) {
	$worksheet_aps11->setCellValue('A'.$row, $row-6);
}


// Luaran Penelitian/PkM Lainnya oleh DTPS - HKI (Paten, Paten Sederhana)
$worksheet_aps12 = $spreadsheet_aps->getSheetByName('3b5-1');
$worksheet_aps12->fromArray($data_luaran_penelitian_dtps_1, NULL, 'B7');

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
$worksheet_aps13->fromArray($data_luaran_penelitian_dtps_2, NULL, 'B7');

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
$worksheet_aps14->fromArray($data_luaran_penelitian_dtps_3, NULL, 'B7');

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
$worksheet_aps15->fromArray($data_luaran_penelitian_dtps_4, NULL, 'B7');

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
$worksheet_aps16->fromArray($data_karya_dtps_disitasi, NULL, 'B6');

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

$highestRow_aps19 = $worksheet_aps19->getHighestRow();

$worksheet_aps19->setCellValue('C11', '=SUM(C6:C10)');
$worksheet_aps19->setCellValue('D11', '=SUM(D6:D10)');
$worksheet_aps19->setCellValue('E11', '=SUM(E6:E10)');
$worksheet_aps19->setCellValue('F11', '=SUM(F6:F10)');

foreach($worksheet_aps19->getRowDimensions() as $rd19) { 
    $rd19->setRowHeight(-1); 
}


// Penggunaan Dana
$worksheet_aps20 = $spreadsheet_aps->getSheetByName('4');
$worksheet_aps20->fromArray($biaya_dosen, NULL, 'C7');
$worksheet_aps20->fromArray($biaya_tendik, NULL, 'C8');
$worksheet_aps20->fromArray($biaya_ops_pembelajaran, NULL, 'C9');
$worksheet_aps20->fromArray($biaya_ops_tdk_langsung, NULL, 'C10');
$worksheet_aps20->fromArray($biaya_ops_mhs, NULL, 'C11');
$worksheet_aps20->fromArray($biaya_penelitian, NULL, 'C13');
$worksheet_aps20->fromArray($biaya_pkm, NULL, 'C14');
$worksheet_aps20->fromArray($biaya_investasi, NULL, 'C16');

$worksheet_aps20->setCellValue('C12', '=SUM(C7:C11)');
$worksheet_aps20->setCellValue('D12', '=SUM(D7:D11)');
$worksheet_aps20->setCellValue('E12', '=SUM(E7:E11)');
$worksheet_aps20->setCellValue('G12', '=SUM(G7:G11)');
$worksheet_aps20->setCellValue('H12', '=SUM(H7:H11)');
$worksheet_aps20->setCellValue('I12', '=SUM(I7:I11)');

$worksheet_aps20->setCellValue('C15', '=SUM(C13:C14)');
$worksheet_aps20->setCellValue('D15', '=SUM(D13:D14)');
$worksheet_aps20->setCellValue('E15', '=SUM(E13:E14)');
$worksheet_aps20->setCellValue('G15', '=SUM(G13:G14)');
$worksheet_aps20->setCellValue('H15', '=SUM(H13:H14)');
$worksheet_aps20->setCellValue('I15', '=SUM(I13:I14)');

$worksheet_aps20->setCellValue('C19', '=SUM(C16:C18)');
$worksheet_aps20->setCellValue('D19', '=SUM(D16:D18)');
$worksheet_aps20->setCellValue('E19', '=SUM(E16:E18)');
$worksheet_aps20->setCellValue('G19', '=SUM(G16:G18)');
$worksheet_aps20->setCellValue('H19', '=SUM(H16:H18)');
$worksheet_aps20->setCellValue('I19', '=SUM(I16:I18)');

for($row = 7; $row <= 19; $row++) {
	$worksheet_aps20->setCellValue('F'.$row, '=AVERAGE(C'.$row.':E'.$row.')');
	$worksheet_aps20->setCellValue('J'.$row, '=AVERAGE(G'.$row.':I'.$row.')');
	
	if($worksheet_aps20->getCell('C'.$row)->getValue() == "" && $worksheet_aps20->getCell('D'.$row)->getValue() == "" && 
	$worksheet_aps20->getCell('E'.$row)->getValue() == "") {
		$worksheet_aps20->setCellValue('F'.$row, 0);
	}
	
	if($worksheet_aps20->getCell('G'.$row)->getValue() == "" && $worksheet_aps20->getCell('H'.$row)->getValue() == "" && 
	$worksheet_aps20->getCell('I'.$row)->getValue() == "") {
		$worksheet_aps20->setCellValue('J'.$row, 0);
	}
}


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


// PkM DTPS yang Melibatkan Mahasiswa
/* $worksheet_aps23 = $spreadsheet_aps->getSheetByName('7');
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
} */


// IPK Lulusan
$worksheet_aps24 = $spreadsheet_aps->getSheetByName('8a');
$worksheet_aps24->fromArray($data_ipk_lulusan, NULL, 'B6');

$highestRow_aps24 = $worksheet_aps24->getHighestRow();

$worksheet_aps24->getStyle('A6:F'.$highestRow_aps24)->applyFromArray($styleBorder);
$worksheet_aps24->getStyle('C6:F'.$highestRow_aps24)->applyFromArray($styleYellow);
$worksheet_aps24->getStyle('A6:F'.$highestRow_aps24)->applyFromArray($styleCenter);
$worksheet_aps24->getStyle('D6:F'.$highestRow_aps24)->getNumberFormat()->setFormatCode('0.00'); 

foreach($worksheet_aps24->getRowDimensions() as $rd24) { 
    $rd24->setRowHeight(-1); 
}

for($row = 6; $row <= $highestRow_aps24; $row++) {
	$worksheet_aps24->setCellValue('A'.$row, $row-5);
}


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
$jumlahRow_aps27 = $highestRow_aps27 + 1;

$worksheet_aps27->setCellValue('A'.$jumlahRow_aps27, 'Jumlah');
$worksheet_aps27->setCellValue('B'.$jumlahRow_aps27, '=SUM(B7:B'.$highestRow_aps27.')');
$worksheet_aps27->setCellValue('C'.$jumlahRow_aps27, '=SUM(C7:C'.$highestRow_aps27.')');
$worksheet_aps27->setCellValue('D'.$jumlahRow_aps27, '=SUM(D7:D'.$highestRow_aps27.')');
$worksheet_aps27->setCellValue('E'.$jumlahRow_aps27, '=SUM(E7:E'.$highestRow_aps27.')');
$worksheet_aps27->setCellValue('F'.$jumlahRow_aps27, '=SUM(F7:F'.$highestRow_aps27.')');

$worksheet_aps27->getStyle('A7:F'.$jumlahRow_aps27)->applyFromArray($styleBorder);
$worksheet_aps27->getStyle('B7:F'.$highestRow_aps27)->applyFromArray($styleYellow);
$worksheet_aps27->getStyle('A7:F'.$highestRow_aps27)->applyFromArray($styleCenter); 
$worksheet_aps27->getStyle('B'.$jumlahRow_aps27.':F'.$jumlahRow_aps27)->applyFromArray($styleCenter); 
$worksheet_aps27->getStyle('A'.$jumlahRow_aps27.':F'.$jumlahRow_aps27)->applyFromArray($styleBold);

foreach($worksheet_aps27->getRowDimensions() as $rd27) { 
    $rd27->setRowHeight(-1); 
}


// Tempat Kerja Lulusan
$worksheet_aps28 = $spreadsheet_aps->getSheetByName('8e1');
$worksheet_aps28->fromArray($data_tempat_kerja, NULL, 'A7');

$highestRow_aps28 = $worksheet_aps28->getHighestRow();
$jumlahRow_aps28 = $highestRow_aps28 + 1;

$worksheet_aps28->setCellValue('A'.$jumlahRow_aps28, 'Jumlah');
$worksheet_aps28->setCellValue('B'.$jumlahRow_aps28, '=SUM(B7:B'.$highestRow_aps28.')');
$worksheet_aps28->setCellValue('C'.$jumlahRow_aps28, '=SUM(C7:C'.$highestRow_aps28.')');
$worksheet_aps28->setCellValue('D'.$jumlahRow_aps28, '=SUM(D7:D'.$highestRow_aps28.')');
$worksheet_aps28->setCellValue('E'.$jumlahRow_aps28, '=SUM(E7:E'.$highestRow_aps28.')');
$worksheet_aps28->setCellValue('F'.$jumlahRow_aps28, '=SUM(F7:F'.$highestRow_aps28.')');

$worksheet_aps28->getStyle('A7:F'.$jumlahRow_aps28)->applyFromArray($styleBorder);
$worksheet_aps28->getStyle('B7:F'.$highestRow_aps28)->applyFromArray($styleYellow);
$worksheet_aps28->getStyle('A7:F'.$highestRow_aps28)->applyFromArray($styleCenter); 
$worksheet_aps28->getStyle('B'.$jumlahRow_aps28.':F'.$jumlahRow_aps28)->applyFromArray($styleCenter); 
$worksheet_aps28->getStyle('A'.$jumlahRow_aps28.':F'.$jumlahRow_aps28)->applyFromArray($styleBold);

foreach($worksheet_aps28->getRowDimensions() as $rd28) { 
    $rd28->setRowHeight(-1); 
}


// Kepuasan Pengguna Lulusan
$worksheet_aps29 = $spreadsheet_aps->getSheetByName('8e2');
$worksheet_aps29->fromArray($data_kepuasan_lulusan, NULL, 'B7');

$highestRow_aps29 = $worksheet_aps29->getHighestRow();

$worksheet_aps29->setCellValue('C14', '=SUM(C7:C13)');
$worksheet_aps29->setCellValue('D14', '=SUM(D7:D13)');
$worksheet_aps29->setCellValue('E14', '=SUM(E7:E13)');
$worksheet_aps29->setCellValue('F14', '=SUM(F7:F13)');

foreach($worksheet_aps29->getRowDimensions() as $rd29) { 
    $rd29->setRowHeight(-1); 
}


// Publikasi Ilmiah Mahasiswa
$worksheet_aps30 = $spreadsheet_aps->getSheetByName('8f1-1');
$worksheet_aps30->fromArray($data_publikasi_ilmiah_mhs, NULL, 'B7');

$highestRow_aps30 = $worksheet_aps30->getHighestRow();
$jumlahRow_aps30 = $highestRow_aps30 + 1;

$worksheet_aps30->setCellValue('A'.$jumlahRow_aps30, 'Jumlah');
$worksheet_aps30->setCellValue('C'.$jumlahRow_aps30, '=SUM(C7:C'.$highestRow_aps30.')');
$worksheet_aps30->setCellValue('D'.$jumlahRow_aps30, '=SUM(D7:D'.$highestRow_aps30.')');
$worksheet_aps30->setCellValue('E'.$jumlahRow_aps30, '=SUM(E7:E'.$highestRow_aps30.')');
$worksheet_aps30->setCellValue('F'.$jumlahRow_aps30, '=SUM(C'.$jumlahRow_aps30.':E'.$jumlahRow_aps30.')');

$worksheet_aps30->mergeCells('A'.$jumlahRow_aps30.':B'.$jumlahRow_aps30);
$worksheet_aps30->getStyle('A7:F'.$jumlahRow_aps30)->applyFromArray($styleBorder);
$worksheet_aps30->getStyle('C7:E'.$highestRow_aps30)->applyFromArray($styleYellow);
$worksheet_aps30->getStyle('A7:A'.$jumlahRow_aps30)->applyFromArray($styleCenter);
$worksheet_aps30->getStyle('C7:F'.$jumlahRow_aps30)->applyFromArray($styleCenter);
$worksheet_aps30->getStyle('A'.$jumlahRow_aps30.':F'.$jumlahRow_aps30)->applyFromArray($styleBold);
$worksheet_aps30->getStyle('B7:B'.$highestRow_aps30)->getAlignment()->setWrapText(true);

foreach($worksheet_aps30->getRowDimensions() as $rd30) { 
    $rd30->setRowHeight(-1); 
}

for($row = 7; $row <= $highestRow_aps30; $row++) {
	$worksheet_aps30->setCellValue('A'.$row, $row-6);
}


// Pagelaran/Pameran/Presentasi/Publikasi Ilmiah Mahasiswa
$worksheet_aps31 = $spreadsheet_aps->getSheetByName('8f1-2');
$worksheet_aps31->fromArray($data_pagelaran_mhs, NULL, 'B7');

$highestRow_aps31 = $worksheet_aps31->getHighestRow();
$jumlahRow_aps31 = $highestRow_aps31 + 1;

$worksheet_aps31->setCellValue('A'.$jumlahRow_aps31, 'Jumlah');
$worksheet_aps31->setCellValue('C'.$jumlahRow_aps31, '=SUM(C7:C'.$highestRow_aps31.')');
$worksheet_aps31->setCellValue('D'.$jumlahRow_aps31, '=SUM(D7:D'.$highestRow_aps31.')');
$worksheet_aps31->setCellValue('E'.$jumlahRow_aps31, '=SUM(E7:E'.$highestRow_aps31.')');
$worksheet_aps31->setCellValue('F'.$jumlahRow_aps31, '=SUM(C'.$jumlahRow_aps31.':E'.$jumlahRow_aps31.')');

$worksheet_aps31->mergeCells('A'.$jumlahRow_aps31.':B'.$jumlahRow_aps31);
$worksheet_aps31->getStyle('A7:F'.$jumlahRow_aps31)->applyFromArray($styleBorder);
$worksheet_aps31->getStyle('C7:E'.$highestRow_aps31)->applyFromArray($styleYellow);
$worksheet_aps31->getStyle('A7:A'.$jumlahRow_aps31)->applyFromArray($styleCenter);
$worksheet_aps31->getStyle('C7:F'.$jumlahRow_aps31)->applyFromArray($styleCenter);
$worksheet_aps31->getStyle('A'.$jumlahRow_aps31.':F'.$jumlahRow_aps31)->applyFromArray($styleBold);
$worksheet_aps31->getStyle('B7:B'.$highestRow_aps31)->getAlignment()->setWrapText(true);

foreach($worksheet_aps31->getRowDimensions() as $rd31) { 
    $rd31->setRowHeight(-1); 
}

for($row = 7; $row <= $highestRow_aps31; $row++) {
	$worksheet_aps31->setCellValue('A'.$row, $row-6);
}


// Karya Ilmiah Mahasiswa yang Disitasi
$worksheet_aps32 = $spreadsheet_aps->getSheetByName('8f2');
$worksheet_aps32->fromArray($data_karya_mhs_disitasi, NULL, 'B6');

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


// Luaran Penelitian/PkM Lainnya oleh Mahasiswa - HKI (Paten, Paten Sederhana)
$worksheet_aps33 = $spreadsheet_aps->getSheetByName('8f4-1');
$worksheet_aps33->fromArray($data_luaran_penelitian_mhs_1, NULL, 'B8');

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
$worksheet_aps34->fromArray($data_luaran_penelitian_mhs_2, NULL, 'B8');

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
$worksheet_aps35->fromArray($data_luaran_penelitian_mhs_3, NULL, 'B8');

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
$worksheet_aps36->fromArray($data_luaran_penelitian_mhs_4, NULL, 'B8');

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


// Masa Studi Lulusan
$worksheet_aps37 = $spreadsheet_aps->getSheetByName('8c');

if(substr($nama_prodi, 0, 3) == "S-1" || substr($nama_prodi, 0, 3) == "D-4") {
	// $worksheet_aps37->fromArray($masa_studi_angkatan1, NULL, 'B18');
	$worksheet_aps37->fromArray($masa_studi_angkatan1, NULL, 'B17');
	$worksheet_aps37->fromArray($masa_studi_angkatan2, NULL, 'B16');
	$worksheet_aps37->fromArray($masa_studi_angkatan3, NULL, 'B15');
} elseif(substr($nama_prodi, 0, 3) == "D-3") {
	$worksheet_aps37->fromArray($masa_studi_angkatan1, NULL, 'B9');
	$worksheet_aps37->fromArray($masa_studi_angkatan2, NULL, 'B8');
	$worksheet_aps37->fromArray($masa_studi_angkatan3, NULL, 'B7');
} elseif(substr($nama_prodi, 0, 3) == "S-2") {
	$worksheet_aps37->fromArray($masa_studi_angkatan1, NULL, 'B26');
	$worksheet_aps37->fromArray($masa_studi_angkatan2, NULL, 'B25');
	$worksheet_aps37->fromArray($masa_studi_angkatan3, NULL, 'B24');
} elseif(substr($nama_prodi, 0, 3) == "S-3") {
	$worksheet_aps37->fromArray($masa_studi_angkatan1, NULL, 'B36');
	$worksheet_aps37->fromArray($masa_studi_angkatan2, NULL, 'B35');
	$worksheet_aps37->fromArray($masa_studi_angkatan3, NULL, 'B34');
	$worksheet_aps37->fromArray($masa_studi_angkatan3, NULL, 'B33');
	$worksheet_aps37->fromArray($masa_studi_angkatan3, NULL, 'B32');
}

$worksheet_aps37->getStyle('I7:I9')->getNumberFormat()->setFormatCode('0.00'); 
$worksheet_aps37->getStyle('K15:K18')->getNumberFormat()->setFormatCode('0.00'); 
$worksheet_aps37->getStyle('H24:H26')->getNumberFormat()->setFormatCode('0.00'); 
$worksheet_aps37->getStyle('K32:K36')->getNumberFormat()->setFormatCode('0.00'); 


// Waktu Tunggu Lulusan
$worksheet_aps38 = $spreadsheet_aps->getSheetByName('8d1');

if(substr($nama_prodi2, 0, 3) == "D-3") {
	$worksheet_aps38->fromArray($waktu_tunggu_diploma_ts4, NULL, 'B7');
	$worksheet_aps38->fromArray($waktu_tunggu_diploma_ts3, NULL, 'B8');
	$worksheet_aps38->fromArray($waktu_tunggu_diploma_ts2, NULL, 'B9');
	
	$worksheet_aps38->setCellValue('B10', '=SUM(B7:B9)');
	$worksheet_aps38->setCellValue('C10', '=SUM(C7:C9)');
	$worksheet_aps38->setCellValue('D10', '=SUM(D7:D9)');
	$worksheet_aps38->setCellValue('E10', '=SUM(E7:E9)');
	$worksheet_aps38->setCellValue('F10', '=SUM(F7:F9)');
	$worksheet_aps38->setCellValue('G10', '=SUM(G7:G9)');
	
} elseif(substr($nama_prodi2, 0, 3) == "S-1") {
	$worksheet_aps38->fromArray($waktu_tunggu_sarjana_ts4, NULL, 'B16');
	$worksheet_aps38->fromArray($waktu_tunggu_sarjana_ts3, NULL, 'B17');
	$worksheet_aps38->fromArray($waktu_tunggu_sarjana_ts2, NULL, 'B18');
	
	$worksheet_aps38->setCellValue('B19', '=SUM(B16:B18)');
	$worksheet_aps38->setCellValue('C19', '=SUM(C16:C18)');
	$worksheet_aps38->setCellValue('D19', '=SUM(D16:D18)');
	$worksheet_aps38->setCellValue('E19', '=SUM(E16:E18)');
	$worksheet_aps38->setCellValue('F19', '=SUM(F16:F18)');
	
} elseif(substr($nama_prodi2, 0, 3) == "D-4") {
	$worksheet_aps38->fromArray($waktu_tunggu_terapan_ts4, NULL, 'B25');
	$worksheet_aps38->fromArray($waktu_tunggu_terapan_ts3, NULL, 'B26');
	$worksheet_aps38->fromArray($waktu_tunggu_terapan_ts2, NULL, 'B27');
	
	$worksheet_aps38->setCellValue('B28', '=SUM(B25:B27)');
	$worksheet_aps38->setCellValue('C28', '=SUM(C25:C27)');
	$worksheet_aps38->setCellValue('D28', '=SUM(D25:D27)');
	$worksheet_aps38->setCellValue('E28', '=SUM(E25:E27)');
	$worksheet_aps38->setCellValue('F28', '=SUM(F25:F27)');
}


// Dosen Pembimbing Utama TA
$worksheet_aps39 = $spreadsheet_aps->getSheetByName('3a2');

for($group = 1; $group <= 88; $group++) {
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


// Mahasiswa Asing
$worksheet_aps40 = $spreadsheet_aps->getSheetByName('2b');

for($group = 1; $group <= 3; $group++) {
	$worksheet_aps40->fromArray(${"mhs_asing".$group}, NULL, 'B'.($group+6));
}
	
$highestRow_aps40 = $worksheet_aps40->getHighestRow();

$worksheet_aps40->getStyle('A7:K'.$highestRow_aps40)->applyFromArray($styleBorder);
$worksheet_aps40->getStyle('B7:K'.$highestRow_aps40)->applyFromArray($styleYellow);
$worksheet_aps40->getStyle('A7:A'.$highestRow_aps40)->applyFromArray($styleCenter);
$worksheet_aps40->getStyle('C7:K'.$highestRow_aps40)->applyFromArray($styleCenter);
$worksheet_aps40->getStyle('A7:A'.$highestRow_aps40)->getAlignment()->setWrapText(true);

foreach($worksheet_aps40->getRowDimensions() as $rd40) { 
    $rd40->setRowHeight(-1); 
}

for($row = 7; $row <= $highestRow_aps40; $row++) {
	$worksheet_aps40->setCellValue('A'.$row, $row-6);
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


// Pagelaran/Pameran/Presentasi/Publikasi Ilmiah DTPS
/* $worksheet_aps42 = $spreadsheet_aps->getSheetByName('3b4-2');
$worksheet_aps42->fromArray($data_pagelaran_dtps, NULL, 'B7');

$highestRow_aps42 = $worksheet_aps42->getHighestRow();
$jumlahRow_aps42 = $highestRow_aps42 + 1;

$worksheet_aps42->setCellValue('A'.$jumlahRow_aps42, 'Jumlah');
$worksheet_aps42->setCellValue('C'.$jumlahRow_aps42, '=SUM(C7:C'.$highestRow_aps42.')');
$worksheet_aps42->setCellValue('D'.$jumlahRow_aps42, '=SUM(D7:D'.$highestRow_aps42.')');
$worksheet_aps42->setCellValue('E'.$jumlahRow_aps42, '=SUM(E7:E'.$highestRow_aps42.')');
$worksheet_aps42->setCellValue('F'.$jumlahRow_aps42, '=SUM(C'.$jumlahRow_aps42.':E'.$jumlahRow_aps42.')');

$worksheet_aps42->mergeCells('A'.$jumlahRow_aps42.':B'.$jumlahRow_aps42);
$worksheet_aps42->getStyle('A7:F'.$jumlahRow_aps42)->applyFromArray($styleBorder);
$worksheet_aps42->getStyle('C7:E'.$highestRow_aps42)->applyFromArray($styleYellow);
$worksheet_aps42->getStyle('A7:A'.$jumlahRow_aps42)->applyFromArray($styleCenter);
$worksheet_aps42->getStyle('C7:F'.$jumlahRow_aps42)->applyFromArray($styleCenter);
$worksheet_aps42->getStyle('A'.$jumlahRow_aps42.':F'.$jumlahRow_aps42)->applyFromArray($styleBold);
$worksheet_aps42->getStyle('B7:B'.$highestRow_aps42)->getAlignment()->setWrapText(true);

foreach($worksheet_aps42->getRowDimensions() as $rd42) { 
    $rd42->setRowHeight(-1); 
}

for($row = 7; $row <= $highestRow_aps42; $row++) {
	$worksheet_aps42->setCellValue('A'.$row, $row-6);
} */


// Produk/Jasa DTPS yang Diadopsi oleh Industri/Masyarakat
/* $worksheet_aps43 = $spreadsheet_aps->getSheetByName('3b7');
$worksheet_aps43->fromArray($data_produk_dtps_diadopsi, NULL, 'B6');

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
} */


// Produk/Jasa Mahasiswa yang Diadopsi oleh Industri/Masyarakat
/* $worksheet_aps44 = $spreadsheet_aps->getSheetByName('8f3');
$worksheet_aps44->fromArray($data_produk_mhs_diadopsi, NULL, 'B6');

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
} */


$writer_aps = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet_aps, 'Xlsx');
$writer_aps->save('./result/sapto_aps9 (F).xlsx');

$spreadsheet_aps->disconnectWorksheets();
unset($spreadsheet_aps);


?>