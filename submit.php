<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx as ReaderXlsx; // Import the Xlsx reader

if ($_SERVER["REQUEST_METHOD"] == "POST") {
  $name = $_POST['name'] ?? '';
  $email = $_POST['email'] ?? '';

  // Check if both name and email are provided
  if (!empty($name) && !empty($email)) {
    // Set the path where you want to save the file
    $filePath = 'D:\filesss.xlsx';

    // Create a new spreadsheet object
    $spreadsheet = new Spreadsheet();

    // Check if the file exists
    if (file_exists($filePath)) {
      // Load the existing file
      try {
        // Use the Xlsx reader explicitly to ensure it's using the correct reader
        $reader = new ReaderXlsx();
        $spreadsheet = $reader->load($filePath);
      } catch (\PhpOffice\PhpSpreadsheet\Reader\Exception $e) {
        die('Error loading file: ' . $e->getMessage());
      }
    } else {
      // Set default styling if creating a new file
      $sheet = $spreadsheet->getActiveSheet();
      $sheet->setTitle('Sheet1');
      $sheet->setCellValue('A1', 'SN');
      $sheet->setCellValue('B1', 'Name');
      $sheet->setCellValue('C1', 'Email');
    }

    // Add the new data
    $sheet = $spreadsheet->getActiveSheet();
    $row = $sheet->getHighestRow() + 1;

    // Generate serial number dynamically
    $sn = $row - 1;
    // putting the data in cell of excel 
    $sheet->setCellValue('A' . $row, $sn);
    $sheet->setCellValue('B' . $row, $name);
    $sheet->setCellValue('C' . $row, $email);

    // Save the file to the specified location with output flushing
    try {
      $writer = new Xlsx($spreadsheet);
      $writer->save($filePath);
      // Flush output after save
      ob_flush();
      flush();
      echo "Data saved successfully!";
    } catch (\PhpOffice\PhpSpreadsheet\Writer\Exception $e) {
      die('Error saving file: ' . $e->getMessage());
    }
  } else {
    echo "Name and email are required!";
  }
} else {
  echo "Invalid request method!";
}
?>
