<?php

header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Type: application/x-www-form-urlencoded');
header('Content-Transfer-Encoding: Binary');
header('Cache-Control: max-age=0');
header('Cache-Control: cache, must-revalidate'); // HTTP/1.1
header('Pragma: public'); // HTTP/1.0

require __DIR__ . '/vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet; //classe responsável pela manipulação da planilha

use PhpOffice\PhpSpreadsheet\IOFactory;

if ($_POST) {

    $file = $_FILES["xlsxfile"];
    $fileName = $file["name"];
    $fileTmpName = $file["tmp_name"];

    if (!empty($fileName)) {
        $allowed_extension = array('xls', 'xlsx');
        $file_array = explode(".", $_FILES['xlsxfile']['name']);
        $file_extension = end($file_array);
        if (in_array($file_extension, $allowed_extension)) {
            $map = array(
                'á' => 'a',
                'à' => 'a',
                'ã' => 'a',
                'â' => 'a',
                'é' => 'e',
                'ê' => 'e',
                'í' => 'i',
                'ó' => 'o',
                'ô' => 'o',
                'õ' => 'o',
                'ú' => 'u',
                'ü' => 'u',
                'ç' => 'c',
                'Á' => 'A',
                'À' => 'A',
                'Ã' => 'A',
                'Â' => 'A',
                'É' => 'E',
                'Ê' => 'E',
                'Í' => 'I',
                'Ó' => 'O',
                'Ô' => 'O',
                'Õ' => 'O',
                'Ú' => 'U',
                'Ü' => 'U',
                'Ç' => 'C'
            );
            // $from = "áàãâéêíóôõúüçÁÀÃÂÉÊÍÓÔÕÚÜÇ";
            // $to = "aaaaeeiooouucAAAAEEIOOOUUC";
            $from = array_keys($map);
            $to = array_values($map);

            $courseNickname = str_replace($from, $to, trim(filter_input(INPUT_POST, "courseNickname")));
            $courseNickname = (empty($courseNickname)) ? "" : " " . $courseNickname;
            $outputFile = trim(filter_input(INPUT_POST, "outputFile"));
            $outputFile = (empty($outputFile)) ? "Lista para o Google Contatos.csv" : $outputFile . ".csv";

            header("Content-Disposition: attachment; filename=\"" . $outputFile . "\"; filename*=UTF-8''\"" . $outputFile . "\"");

            $file_type = IOFactory::identify($fileTmpName);
            $reader = IOFactory::createReader($file_type);
            $spreadsheetIn = $reader->load($fileTmpName);
            $data = $spreadsheetIn->getActiveSheet()->toArray();
            $spreadsheetOut = new Spreadsheet(); //instanciando uma nova planilha
            $sheetOut = $spreadsheetOut->getActiveSheet(); //retornando a aba ativa
            $sheetOut->setTitle("Contatos$courseNickname");
            $sheetOut->setCellValue(
                "A1",
                "Name,Given Name,Additional Name,Family Name,Yomi Name,Given Name Yomi,Additional Name Yomi,Family Name Yomi,Name Prefix,Name Suffix,Initials,Nickname,Short Name,Maiden Name,Birthday,Gender,Location,Billing Information,Directory Server,Mileage,Occupation,Hobby,Sensitivity,Priority,Subject,Notes,Language,Photo,Group Membership,Phone 1 - Type,Phone 1 - Value,Phone 2 - Type,Phone 2 - Value,Organization 1 - Type,Organization 1 - Name,Organization 1 - Yomi Name,Organization 1 - Title,Organization 1 - Department,Organization 1 - Symbol,Organization 1 - Location,Organization 1 - Job Description"
            ); //Definindo a célula A1
            $j = 2;
            foreach ($data as $row) {
                $name = explode(" ", mb_convert_encoding(trim($row[0]), "UTF-8"));
                if (
                    strtolower($name[1]) == "da" || strtolower($name[1]) == "de" ||
                    strtolower($name[1]) == "di" || strtolower($name[1]) == "do" ||
                    strtolower($name[1]) == "du" || strtolower($name[1]) == "das" ||
                    strtolower($name[1]) == "des" || strtolower($name[1]) == "dis" ||
                    strtolower($name[1]) == "dos" || strtolower($name[1]) == "dus"
                ) {
                    $lastName = trim($name[1]) . " " . mb_substr(trim($name[2]), 0, 1);
                } else {
                    $lastName = mb_substr(trim($name[1]), 0, 1);
                }
                $name = str_replace($from, $to, trim($name[0])) . " " . $lastName . $courseNickname;
                $phone1 = str_replace(["(", ")", " ", "-"], "", trim($row[1]));
                $phone2 = (!empty(trim($row[2]))) ? ",Home," . str_replace(["(", ")", " ", "-"], "", trim($row[2])) . ",,,,,,,," : ",,,,,,,,,,";
                $content = $name . ",,,,,,,,,,,,,,,,,,,,,,,,,,,,,Mobile," . $phone1 . $phone2;
                $sheetOut->setCellValue("A$j", $content);
                $j++;
            }

            // $writer = IOFactory::createWriter($spreadsheetOut, 'Csv');
            $writer = IOFactory::createWriter($spreadsheetOut, 'Csv')->setDelimiter(',')->setEnclosure('"')->setSheetIndex(0);
            $writer->setExcelCompatibility(true);
            // $writer->save('php://output');

            $writer->save($outputFile);

            readfile($outputFile);

            unlink($outputFile);

            exit;
            // $message = 1;

            // echo "
            // <script>
            //     document.getElementById('process').removeAttribute('disabled');
            //     document.getElementById('process').value = 'Processar';
            //     if (document.getElementById('message').classList.contains('green')) {
            //         document.getElementById('message').classList.remove('green');
            //     } else {
            //         document.getElementById('message').classList.remove('red');
            //     }
            //     document.getElementById('message').classList.add('green');
            //     document.getElementById('message').innerHTML = 'Processamento realizado com sucesso!';
            //     document.getElementById('uploadForm').reset();
            // </script>";
        } else {
            // $message = 2;
            // echo "
            // <script>
            //     if (document.getElementById('message').classList.contains('green')) {
            //         document.getElementById('message').classList.remove('green');
            //     } else {
            //         document.getElementById('message').classList.remove('red');
            //     }
            //     document.getElementById('message').classList.add('red');
            //     document.getElementById('message').innerHTML = 'Somente arquivos .xls ou .xlsx são permitidos.';
            // </script>";
        }
    } else {
        // $message = 3;
        // echo "
        // <script>
        //     if (document.getElementById('message').classList.contains('green')) {
        //         document.getElementById('message').classList.remove('green');
        //     } else {
        //         document.getElementById('message').classList.remove('red');
        //     }
        //     document.getElementById('message').classList.add('red');
        //     document.getElementById('message').innerHTML = 'Nenhum arquivo foi selecionado. Por favor, selecione um arquivo.';
        // </script>";
    }
}

// echo $message;
