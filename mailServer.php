<?php
    require_once 'spreadSheetParser.php';

    function userMail_login($host,$port,$user,$pass,$folder="INBOX",$ssl=false)
    {
        $ssl=($ssl==false)?"/novalidate-cert":"";
        return (imap_open("{"."$host:$port/imap/ssl$ssl"."}$folder",$user,$pass));
    } 

    function processInboxMail($connection, $startDate="", $endIndex="")
    {   
        try {            
            $response = imap_search($connection,'SINCE "18 April 2019"'); 
            checkMailStructure($connection, $response);
        }catch(Exception $e) {
            imap_close($connection);
        }
    }

    function checkMailStructure($connection, $response){
        try{
            foreach ($response as $msg) {
                $email_number = $msg;//$msg->msgno;
                $mailInfo = imap_fetchstructure($connection, $email_number);
                if(isset($mailInfo->parts) && count($mailInfo->parts))
                {
                    processAttachmentData($connection, $email_number, $mailInfo);
                }else{
                    $retMailData = true;
                }
            }
            imap_close($connection);
        }catch(Exception $e) {
            imap_close($connection);
        }
    }

    function endsWith($haystack, $needle)
    {
        $length = strlen($needle);
        if ($length == 0) {
            return true;
        }
        return (substr($haystack, -$length) === $needle);
    }

    function processAttachmentData($connection, $message_number, $structure){
        $attachments = array();
        for($i = 0; $i < count($structure->parts); $i++) {
            $attachments[$i] = array(
                'is_attachment' => false,
                'filename' => '',
                'name' => '',
                'attachment' => ''
            );
            
            if($structure->parts[$i]->ifdparameters) {
                foreach($structure->parts[$i]->dparameters as $object) {
                    if(strtolower($object->attribute) == 'filename') {
                        $attachments[$i]['is_attachment'] = true;
                        $attachments[$i]['filename'] = $object->value;
                    }
                }
            }
            
            if($structure->parts[$i]->ifparameters) {
                foreach($structure->parts[$i]->parameters as $object) {
                    if(strtolower($object->attribute) == 'name') {
                        $attachments[$i]['is_attachment'] = true;
                        $attachments[$i]['name'] = $object->value;
                    }
                }
            }
            
            if($attachments[$i]['is_attachment'] && endsWith($attachments[$i]['name'],".xlsx")) {
                echo $attachments[$i]['name'];
                $attachments[$i]['attachment'] = imap_fetchbody($connection, $message_number, $i+1);
                if($structure->parts[$i]->encoding == 3) { // 3 = BASE64
                    $attachments[$i]['attachment'] = base64_decode($attachments[$i]['attachment']);
                }
                elseif($structure->parts[$i]->encoding == 4) { // 4 = QUOTED-PRINTABLE
                    $attachments[$i]['attachment'] = quoted_printable_decode($attachments[$i]['attachment']);
                }
                file_put_contents( $attachments[$i]['name'], $attachments[$i]['attachment']);
                openSpreadSheet( $attachments[$i]['name']);
            }
        }
    }

    function getInboxDetails($connection,$messageSize="")
    {
        if ($messageSize)
        {
            //$range=$messageSize;
            $MC = imap_check($connection);
            $range = "1:20";
        } else {
            $MC = imap_check($connection);
            $range = "1:20";
        }
        $response = imap_fetch_overview($connection,$range);
        
         foreach ($response as $msg) {
            $fifthEle = $msg->msgno;
            $result[$msg->msgno]=(array)$msg;
            echo $msg->msgno;

        }

        return setFlaggedOnMail($connection, $fifthEle);
    }

    function setFlaggedOnMail($connection, $sequence){
        $flag = "\\FLAGGED";
        $status = imap_setflag_full ( $connection, $sequence , $flag );
        echo gettype($status) . "\n";
        echo $status . "\n";
        getAttachmentMail($connection);

    }

    function getAttachmentMail($connection){
        $unreadEmails = imap_mail_move($connection,'14','INBOX.Test');
        $email_number = $unreadEmails[0];
    }

?>