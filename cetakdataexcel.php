<?php
include('koneksi.php');
require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\xlsx;

$spreadsheet=new Spreadsheet();
$sheet=$spreadsheet->getActiveSheet();
$sheet->setCellValue('A1', 'No');
$sheet->setCellValue('B1', 'ID');
$sheet->setCellValue('C1', 'Jenis Pendaftaran');
$sheet->setCellValue('D1', 'Tanggal Masuk Sekolah');
$sheet->setCellValue('E1', 'NIS');
$sheet->setCellValue('F1', 'Nomer Peserta Ujian');
$sheet->setCellValue('G1', 'Pernah PAUD');
$sheet->setCellValue('H1', 'Pernah TK');
$sheet->setCellValue('I1', 'SKHUN');
$sheet->setCellValue('J1', 'Ijazah');
$sheet->setCellValue('K1', 'Hobi');
$sheet->setCellValue('L1', 'Cita-cita');
$sheet->setCellValue('M1', 'Nama');
$sheet->setCellValue('N1', 'Jenis Kelamin');
$sheet->setCellValue('O1', 'NISN');
$sheet->setCellValue('P1', 'NIK');
$sheet->setCellValue('Q1', 'Tempat Lahir');
$sheet->setCellValue('R1', 'Tanggal Lahir');
$sheet->setCellValue('S1', 'Agama');
$sheet->setCellValue('T1', 'Berkebutuhan Khusus');
$sheet->setCellValue('U1', 'Alamat');
$sheet->setCellValue('V1', 'RT');
$sheet->setCellValue('W1', 'RW');
$sheet->setCellValue('X1', 'Dusun');
$sheet->setCellValue('Y1', 'Desa');
$sheet->setCellValue('Z1', 'Kecamatan');
$sheet->setCellValue('AA1', 'Kode Pos');
$sheet->setCellValue('AB1', 'Tempat Tinggal');
$sheet->setCellValue('AC1', 'Transportasi');
$sheet->setCellValue('AD1', 'HP');
$sheet->setCellValue('AE1', 'Telp');
$sheet->setCellValue('AF1', 'Email');
$sheet->setCellValue('AG1', 'Penerima KPS');
$sheet->setCellValue('AH1', 'No. KPS');
$sheet->setCellValue('AI1', 'Kewarganegaraan');
$sheet->setCellValue('AJ1', 'Nama Negara');

$query=mysqli_query($con, "select * from data_diri");
$i=2;
$no=1;
while ($row=mysqli_fetch_array($query)){
	$sheet->setCellValue('A'.$i, $no++);
	$sheet->setCellValue('B'.$i, $row['id_data']);
	$sheet->setCellValue('C'.$i, $row['jenis_pendaftaran']);
	$sheet->setCellValue('D'.$i, $row['tanggal_masuk_sekolah']);
	$sheet->setCellValue('E'.$i, $row['nis']);
	$sheet->setCellValue('F'.$i, $row['nomer_peserta_ujian']);
	$sheet->setCellValue('G'.$i, $row['pernah_paud']);
	$sheet->setCellValue('H'.$i, $row['pernah_tk']);
	$sheet->setCellValue('I'.$i, $row['skhun']);
	$sheet->setCellValue('J'.$i, $row['ijazah']);
	$sheet->setCellValue('K'.$i, $row['hobi']);
	$sheet->setCellValue('L'.$i, $row['citacita']);
	$sheet->setCellValue('M'.$i, $row['nama']);
	$sheet->setCellValue('N'.$i, $row['jenis_kelamin']);
	$sheet->setCellValue('O'.$i, $row['nisn']);
	$sheet->setCellValue('P'.$i, $row['nik']);
	$sheet->setCellValue('Q'.$i, $row['tempat_lahir']);
	$sheet->setCellValue('R'.$i, $row['tanggal_lahir']);
	$sheet->setCellValue('S'.$i, $row['agama']);
	$sheet->setCellValue('T'.$i, $row['berkebutuhan_khusus']);
	$sheet->setCellValue('U'.$i, $row['alamat']);
	$sheet->setCellValue('V'.$i, $row['rt']);
	$sheet->setCellValue('W'.$i, $row['rw']);
	$sheet->setCellValue('X'.$i, $row['dusun']);
	$sheet->setCellValue('Y'.$i, $row['desa']);
	$sheet->setCellValue('Z'.$i, $row['kecamatan']);
	$sheet->setCellValue('AA'.$i, $row['kode_pos']);
	$sheet->setCellValue('AB'.$i, $row['tempat_tinggal']);
	$sheet->setCellValue('AC'.$i, $row['transportasi']);
	$sheet->setCellValue('AD'.$i, $row['hp']);
	$sheet->setCellValue('AE'.$i, $row['telp']);
	$sheet->setCellValue('AF'.$i, $row['email']);
	$sheet->setCellValue('AG'.$i, $row['penerima_kps']);
	$sheet->setCellValue('AH'.$i, $row['no_kps']);
	$sheet->setCellValue('AI'.$i, $row['kewarganegaraan']);
	$sheet->setCellValue('AJ'.$i, $row['nama_negara']);
	$i++;
}

$styleArray = [
	'borders' => [
		'allBorders' => [
			'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN, 
		],
	],
];
$i=$i-1;
$sheet->getStyle('A1:AJ'.$i)->applyFromArray($styleArray);
$writer = new Xlsx($spreadsheet);
$writer->save('Report Data Pendaftar.xlsx');
?>
<!DOCTYPE html>
<html>
<head>
	<title></title>
</head>
<body>
<br><br><br><center><h1>Berhasil Simpan Data</h1></center>
</body>
</html>
}
