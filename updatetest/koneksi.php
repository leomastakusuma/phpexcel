<?php
$dbhost = "localhost";
$dbuser = "root";
$dbpass = "root";
$dbname = "garudacos";
$dbc = @mysql_connect( $dbhost , $dbuser , $dbpass ) or die( mysql_error() );
mysql_select_db( $dbname );
?>