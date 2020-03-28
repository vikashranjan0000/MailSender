<?php
	require __DIR__ . '/vendor/autoload.php';
	require_once 'mailServer.php';
	$servername = "xtudy.in";
	$portNo = 993;
	$username = "avinash.kumar@xtudy.in";
	$password = "Xtudy@123";
	$dbname ="avinashk";
	$folder="INBOX";
	$ssl=false;
	$conn = userMail_login($servername, $portNo, $username, $password, $folder, $ssl);
	$check = processInboxMail($conn, 1);
 ?>