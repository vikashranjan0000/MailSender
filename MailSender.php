<?php
use PHPMailer\PHPMailer\PHPMailer;
require 'vendor/autoload.php';
$mail = new PHPMailer;
$mail->isSMTP();
$mail->SMTPDebug = 2;
$mail->Host = 'mail.xtudy.in';
$mail->Port = 587;
$mail->protocol = 'mail';
$mail->SMTPAuth = true;
$mail->SMTPSecure = false;
$mail->SMTPOptions = array(
    'ssl' => array(
        'verify_peer' => false,
        'verify_peer_name' => false,
        'allow_self_signed' => true
    )
);
//$mail->SMTPSecure = 'tls'; 
//$mail->SMTPAuth = true;
$mail->Username = 'avinash.kumar@xtudy.in';
$mail->Password = 'Xtudy@123';
$mail->setFrom('avinash.kumar@xtudy.in', 'Avinash Kumar');
$mail->addReplyTo('avinash.kumar@xtudy.in', 'Avinash Kumar');
$mail->addAddress('arun.kumar@xtudy.in', 'Arun Kumar');
$mail->Subject = 'PHPMailer SMTP message';
//$mail->msgHTML(file_get_contents('message.html'), __DIR__);
//$mail->msgHTML("this is hello world mail!!");
$mail->Body ="Dear receiver, this is test mail..";
$mail->AltBody = 'This is a plain text message body';
//$mail->addAttachment('Report.xlsx');
if (!$mail->send()) {
    echo 'Mailer Error: ' . $mail->ErrorInfo;
} else {
    echo 'Message sent!';
}
?>

