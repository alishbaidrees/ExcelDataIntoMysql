<?php
include("config.php");
require("excelimp/vendor/autoload.php");
use PhpOffice\PhpSpreadsheet\IOFactory;


$sql = "SELECT * FROM sheet";
$result = $conn->query($sql);

?>

<!DOCTYPE html>
<html>
<head>
    <title>Excel Import</title>
</head>
<body>
    <form method="post" enctype="multipart/form-data">
        <input type="file" name="excel">
        <input type="submit" name="btn" value="IMPORT">
    </form>

    <table border="1" cellspacing="0" cellpadding="5px">
        <tr>
            <th>#id</th>
            <th>Name</th>
            <th>Age</th>
            <th>City</th>
        </tr>

        <?php
        if ($result->num_rows > 0) {
            while ($row = $result->fetch_assoc()) {
                ?>
                <tr>
                    <td><?php echo $row['id']; ?></td>
                    <td><?php echo $row['name']; ?></td>
                    <td><?php echo $row['age']; ?></td>
                    <td><?php echo $row['city']; ?></td>
                </tr>
                <?php
            }
        }
        ?>

    </table>

    <?php
    if (isset($_FILES['excel'])) {
        $filename = $_FILES['excel']['name'];
        $tmp = $_FILES['excel']['tmp_name'];
        $fileext = strtolower(pathinfo($filename, PATHINFO_EXTENSION));

        if ($fileext == 'xlsx' || $fileext == 'xls' || $fileext == 'csv') {
            $targetdir = "uploads/" . $filename;

            if (move_uploaded_file($tmp, $targetdir)) {
                try {
                    $spreadsheet = IOFactory::load($targetdir);
                    $data = $spreadsheet->getActiveSheet()->toArray();

                    foreach ($data as $key => $row) {
                        $name = $row[0];
                        $age = $row[1];
                        $city = $row[2];

                        // Use prepared statement to prevent SQL injection
                        $sql2 = "INSERT INTO sheet (name, age, city) VALUES (?, ?, ?)";
                        $stmt = $conn->prepare($sql2);
                        $stmt->bind_param("sis", $name, $age, $city);

                        if ($stmt->execute()) {
                            echo "Mission successful";
                        } else {
                            echo "Failed to insert data";
                        }

                        $stmt->close();
                    }
                } catch (Exception $e) {
                    echo "Error: " . $e->getMessage();
                }
            } else {
                echo "Error uploading the file.";
            }
        } else {
            echo "Invalid file format. Allowed formats: xlsx, xls, csv.";
        }
    }
    ?>
</body>
</html>
