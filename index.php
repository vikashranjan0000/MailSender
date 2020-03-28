<?php
 require __DIR__ . '/vendor/autoload.php';


require_once 'mailServer.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

 $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
 $data = $reader->load("New-Session-Data.xlsx");
 //echo $spreadsheet;


$worksheet = $data->getActiveSheet();
echo $worksheet->getCell('A1');
echo $worksheet->getCell('A2');

//$data = new Spreadsheet_Excel_Reader("New-Session-Data.xlsx");
 
echo "Total Sheets in this xls file: ".count($data->sheets)."<br /><br />";
 
$html="<table border='1'>";
for($i=0;$i<count($data->sheets);$i++) // Loop to get all sheets in a file.
{ 
 if(count($data->sheets[$i][cells])>0) // checking sheet not empty
 {
 echo "Sheet $i:<br /><br />Total rows in sheet $i  ".count($data->sheets[$i][cells])."<br />";
 for($j=1;$j<=count($data->sheets[$i][cells]);$j++) // loop used to get each row of the sheet
 { 
 $html.="<tr>";
 for($k=1;$k<=count($data->sheets[$i][cells][$j]);$k++) // This loop is created to get data in a table format.
 {
 $html.="<td>";
 $html.=$data->sheets[$i][cells][$j][$k];
 $html.="</td>";
 }
 $eid = mysqli_real_escape_string($connection,$data->sheets[$i][cells][$j][1]);
 $name = mysqli_real_escape_string($connection,$data->sheets[$i][cells][$j][2]);
 $email = mysqli_real_escape_string($connection,$data->sheets[$i][cells][$j][3]);
 $dob = mysqli_real_escape_string($connection,$data->sheets[$i][cells][$j][4]);
 $query = "insert into excel(eid,name,email,dob) values('".$eid."','".$name."','".$email."','".$dob."')";
 
 mysqli_query($connection,$query);
 $html.="</tr>";
 }
 }
 
}
 
$html.="</table>";
echo $html;
echo "<br />Data Inserted in dababase";
?>