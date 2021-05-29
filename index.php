<!DOCTYPE html>
<html lang="pt-BR">

<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>XLSX to CSV</title>
    <style>
        .red {
            color: red;
            font-weight: bold;
        }

        .green {
            color: green;
            font-weight: bold;
        }
    </style>
</head>

<body>
    <h1>XLSX to CSV</h1>
    <form id="uploadForm" method="post" enctype="multipart/form-data" action="process.php">
        <p>
            <label for="courseNickname">Apelido do curso: </label>
            <input type="text" name="courseNickname" id="courseNickname" placeholder="Ex.: NP4">
        </p>
        <p>
            <label for="xlsxfile">Carregar arquivo: </label>
            <input type="file" name="xlsxfile" id="xlsxfile">
        </p>
        <p>
            <label for="outputFile">Nome do arquivo de saída: </label>
            <input type="text" name="outputFile" id="outputFile" placeholder="Ex.: Turma de NP4">
        </p>
        <p>
            <input type="submit" name="process" id="process" value="Processar">
        </p>
        <p>
            <span id="message"></span>
        </p>
    </form>
    <script>
        window.onload = document.getElementById("courseNickname").focus();
    </script>
</body>

</html>