<?php


$localhost = "localhost";
$user = "root";
$pass = '';
$db = "excel";
$conn = new MYSQli($localhost,$user,$pass,$db);

if($conn->connect_error){
    die("Connection failed".$conn->connect_error);
}


?>