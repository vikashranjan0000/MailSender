<?php
	require __DIR__ . '/vendor/autoload.php';


	function getDBConn(){
		$servername = "localhost";
		$username = "root";
		$password = "pwdpwd";
		$dbname = "xtudy_mail_report";
		try{
			$conn = new mysqli($servername, $username, $password, $dbname);			
			if ($conn->connect_error) {
			    die("Connection failed: " . $conn->connect_error);
			    return false;
			}else {
				return $conn;
			}
		}catch(Exception $e){
			return false;
		}
	}


	function insertMultiEntry($conn, $sql){
		try{
			if ($conn->multi_query($sql) === TRUE) {
		    	echo "New records created successfully";
			} else {
			    echo "Error: " . $sql . "<br>" . $conn->error;
			}	
		}catch(Exception $e){
			
		}
	}

	function closeConn($conn){
		$conn->close();
	}

	
?>