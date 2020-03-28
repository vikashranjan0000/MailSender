<?php
 	require __DIR__ . '/vendor/autoload.php';
	require_once 'mailServer.php';
	require_once 'DBUtils.php';

	use PhpOffice\PhpSpreadsheet\Spreadsheet;
	use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

	function openSpreadSheet($fileName){
		$reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
		$spreadSheet = $reader->load($fileName);
		processWorksheet($spreadSheet, $fileName);
	}
	function modifyDuplicate($worksheet, $rowNumber, $spreadsheet, $fileName){
	    $highestColumn = $worksheet->getHighestColumn();
		$highestColumn++;
	    for ($col = 'A'; $col != $highestColumn; $col++) {
			$worksheet->getStyle($col.$rowNumber)->getFont()->setBold(true);
			$worksheet->getStyle($col.$rowNumber)->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)->getStartColor()->setARGB('64FF0000');
	    }
	    $worksheet->getComment('A'.$rowNumber)
	    ->getText()->createTextRun('Duplicate record');		
	}



	function getEventListSQL($eventsheet){
		$highestRow = $eventsheet->getHighestRow(); 
		$highestColumn = $eventsheet->getHighestColumn(); 
		$highestColumnIndex = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($highestColumn); 

		$sql = "";
		for ($row = 2; $row <= $highestRow; ++$row) {
			$Sr_No = $eventsheet->getCellByColumnAndRow(1, $row)->getValue();
			$E_Month = $eventsheet->getCellByColumnAndRow(2, $row)->getValue();
			$E_Date = $eventsheet->getCellByColumnAndRow(3, $row)->getValue();
			$E_Name = $eventsheet->getCellByColumnAndRow(4, $row)->getValue();
			$E_Year = $eventsheet->getCellByColumnAndRow(5, $row)->getValue();
			$sql .= "INSERT INTO event_list (`Sr_No`, `E_Month`, `E_Date`, `E_Name`, `E_Year`) VALUES (".$Sr_No.", '".$E_Month."', '".$E_Date."', '".$E_Name."', '".$E_Year."');";
		}
		return $sql;
	}

	function getTeacherSQL($eventsheet){
		$highestRow = $eventsheet->getHighestRow(); 
		$highestColumn = $eventsheet->getHighestColumn(); 
		$highestColumnIndex = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($highestColumn); 

		$sql = "";
		for ($row = 2; $row <= $highestRow; ++$row) {
			$Name = $eventsheet->getCellByColumnAndRow(1, $row)->getValue();
			$DOB = $eventsheet->getCellByColumnAndRow(2, $row)->getValue();
			$isBus = $eventsheet->getCellByColumnAndRow(3, $row)->getValue();
			$BusBoardingMonth = $eventsheet->getCellByColumnAndRow(4, $row)->getValue();
			$BusNumber = $eventsheet->getCellByColumnAndRow(5, $row)->getValue();
			$BusStopage = $eventsheet->getCellByColumnAndRow(6, $row)->getValue();
			$Gender = $eventsheet->getCellByColumnAndRow(7, $row)->getValue();
			$Mobile = $eventsheet->getCellByColumnAndRow(8, $row)->getValue();
			$Phone = $eventsheet->getCellByColumnAndRow(9, $row)->getValue();
			$EmailId = $eventsheet->getCellByColumnAndRow(10, $row)->getValue();
			$ClassSection = $eventsheet->getCellByColumnAndRow(11, $row)->getValue();
			$sql .= "INSERT INTO `teacher` (`Name`, `DOB`, `isBus`, `BusBoardingMonth`, `BusNumber`, `BusStopage`, `Gender`, `Mobile`, `Phone`, `EmailId`, `ClassSection`) VALUES ('".$Name."', '".$DOB."', '".$isBus."', '".$BusBoardingMonth."', '".$BusNumber."', '".$BusStopage."', '".$Gender."', '".$Mobile."', '".$Phone."', '".$EmailId."','".$ClassSection."');";
		}
		return $sql;
	}

	function processWorksheet($spreadSheet, $fileName){
		$sheetName = $spreadSheet->getSheetNames();

        $eventsheet = $spreadSheet->getSheetByName("Event List");
        if($eventsheet){
        	$eventSQL = getEventListSQL($eventsheet);
        	$connection = getDBConn();
			insertMultiEntry($connection, $eventSQL);
			closeConn($connection);
        }
        $eventsheet = $spreadSheet->getSheetByName("Teacher");
        if($eventsheet){
        	$teacherSQL = getTeacherSQL($eventsheet);
        	$connection = getDBConn();
			insertMultiEntry($connection, $teacherSQL);
			closeConn($connection);
        }
        $eventsheet = $spreadSheet->getSheetByName("Student");
        if($eventsheet){
        	$studSQL = getStudentSQL($eventsheet, $spreadSheet, $fileName);
        	$connection = getDBConn();
			insertMultiEntry($connection, $studSQL);
			closeConn($connection);
        }
        $writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadSheet, 'Xlsx');
		$writer->save('Report.xlsx');
	}



	function getStudentSQL($eventsheet, $spreadSheet, $fileName){
		$admssionNumber = array();
		$counter = 0;
		$highestRow = $eventsheet->getHighestRow(); 
		$highestColumn = $eventsheet->getHighestColumn(); 
		$highestColumnIndex = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($highestColumn); 

		$sql = "";
		for ($row = 2; $row <= $highestRow; ++$row) {
			$AdmissionNo = $eventsheet->getCellByColumnAndRow(1, $row)->getValue();
			if(!$AdmissionNo){
				continue;
			}else{
				if(in_array($AdmissionNo,$admssionNumber)){
					modifyDuplicate($eventsheet, $row,$spreadSheet, $fileName);
					continue;
				}
			}
			$admssionNumber[$counter++] = $AdmissionNo;
			$RollNumber = $eventsheet->getCellByColumnAndRow(2, $row)->getValue();
			$Name = $eventsheet->getCellByColumnAndRow(3, $row)->getValue();
			$DOB = $eventsheet->getCellByColumnAndRow(4, $row)->getValue();
			$FatherName = $eventsheet->getCellByColumnAndRow(5, $row)->getValue();
			$FatherQualification = $eventsheet->getCellByColumnAndRow(6, $row)->getValue();
			$FatherOccupation = $eventsheet->getCellByColumnAndRow(7, $row)->getValue();
			$MotherName = $eventsheet->getCellByColumnAndRow(8, $row)->getValue();
			$Address = $eventsheet->getCellByColumnAndRow(9, $row)->getValue();
			$District = $eventsheet->getCellByColumnAndRow(10, $row)->getValue();
			$BloodGroup = $eventsheet->getCellByColumnAndRow(11, $row)->getValue();
			$Height = $eventsheet->getCellByColumnAndRow(12, $row)->getValue();
			$Weight = $eventsheet->getCellByColumnAndRow(13, $row)->getValue();
			$Category = $eventsheet->getCellByColumnAndRow(14, $row)->getValue();
			$Religion = $eventsheet->getCellByColumnAndRow(15, $row)->getValue();
			$isBus = $eventsheet->getCellByColumnAndRow(16, $row)->getValue();
			$BusBoardingMonth = $eventsheet->getCellByColumnAndRow(17, $row)->getValue();
			$BusNumber = $eventsheet->getCellByColumnAndRow(18, $row)->getValue();
			$BusStopage = $eventsheet->getCellByColumnAndRow(19, $row)->getValue();
			$Gender = $eventsheet->getCellByColumnAndRow(20, $row)->getValue();
			$Mobile = $eventsheet->getCellByColumnAndRow(21, $row)->getValue();
			$Phone = $eventsheet->getCellByColumnAndRow(22, $row)->getValue();
			$EmailId = $eventsheet->getCellByColumnAndRow(23, $row)->getValue();
			$Class = $eventsheet->getCellByColumnAndRow(24, $row)->getValue();
			$ClassOfFirstAdmission = $eventsheet->getCellByColumnAndRow(25, $row)->getValue();
			$AdmissionDate = $eventsheet->getCellByColumnAndRow(26, $row)->getValue();
			$sql .= "INSERT INTO `student`(`AdmissionNo`, `RollNumber`, `Name`, `DOB`, `FatherName`, `FatherQualification`, `FatherOccupation`, `MotherName`, `Address`, `District`, `BloodGroup`, `Height`, `Weight`, `Category`, `Religion`, `isBus`, `BusBoardingMonth`, `BusNumber`, `BusStopage`, `Gender`, `Mobile`, `Phone`, `EmailId`, `Class`, `ClassOfFirstAdmission`, `AdmissionDate`) VALUES ('".$AdmissionNo."', '".$RollNumber."', '".$Name."', '".$DOB."', '".$FatherName."', '".$FatherQualification."', '".$FatherOccupation."', '".$MotherName."', '".$Address."', '".$District."','".$BloodGroup."', '".$Height."', '".$Weight."', '".$Category."', '".$Religion."', '".$isBus."', '".$BusBoardingMonth."', '".$BusNumber."', '".$BusStopage."', '".$Gender."', '".$Mobile."', '".$Phone."', '".$EmailId."', '".$Class."', '".$ClassOfFirstAdmission."', '".$AdmissionDate."') ";

			$sql .=" ON DUPLICATE KEY UPDATE ";


			if($RollNumber){
				$sql .="RollNumber= '".$RollNumber."'";
			}
			if($Name){
				$sql .=",Name= '".$Name."'";
			}
			if($DOB){
				$sql .=",DOB= '".$DOB."'";
			}
			if($FatherName){
				$sql .=",FatherName= '".$FatherName."'";
			}
			if($FatherQualification){
				$sql .=",FatherQualification= '".$FatherQualification."'";
			}
			if($FatherOccupation){
				$sql .=",FatherOccupation= '".$FatherOccupation."'";
			}
			if($MotherName){
				$sql .=",MotherName= '".$MotherName."'";
			}
			if($Address){
				$sql .=",Address= '".$Address."'";
			}
			if($District){
				$sql .=",District= '".$District."'";
			}
			if($BloodGroup){
				$sql .=",BloodGroup='".$BloodGroup."'";
			}
			if($Height){
				$sql .=",Height= '".$Height."'";
			}
			if($Weight){
				$sql .=",Weight= '".$Weight."'";
			}
			if($Category){
				$sql .=",Category= '".$Category."'";
			}
			if($Religion){
				$sql .=",Religion= '".$Religion."'";
			}
			if($isBus){
				$sql .=",isBus= '".$isBus."'";
			}
			if($BusBoardingMonth){
				$sql .=",BusBoardingMonth= '".$BusBoardingMonth."'";
			}
			if($BusNumber){
				$sql .=",BusNumber= '".$BusNumber."'";
			}
			if($BusStopage){
				$sql .=",BusStopage= '".$BusStopage."'";
			}
			if($Gender){
				$sql .=",Gender= '".$Gender."'";
			}
			if($Mobile){
				$sql .=",Mobile= '".$Mobile."'";
			}
			if($Phone){
				$sql .=",Phone= '".$Phone."'";
			}
			if($EmailId){
				$sql .=",EmailId= '".$EmailId."'";
			}
			if($Class){
				$sql .=",Class= '".$Class."'";
			}
			if($ClassOfFirstAdmission){
				$sql .=",ClassOfFirstAdmission= '".$ClassOfFirstAdmission."'";
			}
			if($AdmissionDate){
				$sql .=",AdmissionDate= '".$AdmissionDate."';";
			}
		}
		return $sql;
	}

	function getSpreadSheet($fileName){
		$reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
		$spreadsheet = $reader->load($fileName);
		$worksheet = $data->getActiveSheet();
		return $spreadsheet;
	}
	
 ?>