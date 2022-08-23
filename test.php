<?php
use Phppot\DataSource;

require_once 'DataSource.php';
$db = new DataSource();
$conn = $db->getConnection();

/* check if server is alive */
if ($conn->ping()) {
    printf ("The connection is ok!\n");
} else {
    printf ("Error: %s\n", $mysqli->error);
}

?>