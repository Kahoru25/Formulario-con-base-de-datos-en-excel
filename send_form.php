<?php

require 'vendor/autoload.php'; 

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

if ($_SERVER["REQUEST_METHOD"] == "POST") {
    $nombre = $_POST["nombre"];
    $correo = $_POST["correo"];
    $dni = $_POST["dni"];
    $celular = $_POST["celular"];
    $residencia = $_POST["residencia"];
    $correspondencia = $_POST["correspondencia"];
    $genero = $_POST["genero"];
    $eps = $_POST["eps"];
    $rh = $_POST["rh"];
    $numEmergencia = $_POST["numEmergencia"];
    $nacimiento = $_POST["nacimiento"];
    $edad = $_POST["edad"];
    $profesion = $_POST["profesion"];
    $mascotas = $_POST["mascotas"];
    $transporte = $_POST["transporte"];
    $genCamisa = $_POST["genCamisa"];
    $tallas = $_POST["tallas"];
    $nombrecamisa = $_POST["nombrecamisa"];
    $nombredeseado = $_POST["nombredeseado"];
    $grupo = $_POST["grupo"];
    $terminos1 = $_POST["terminos1"];
    $autoriza = $_POST["autoriza"];
    $terminos2 = $_POST["terminos2"];

    // ----------------------------------------------
// Crea Spreadsheet
    $rutaArchivo = 'download/neusa.xlsx'; // ruta
    if (file_exists($rutaArchivo)) {
        // Carga el archivo existente
        $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($rutaArchivo);
        $sheet = $spreadsheet->getActiveSheet();
    } else {
        // Crea un nuevo archivo con encabezados de columna
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->setCellValue('A1', 'Nombre');
        $sheet->setCellValue('B1', 'Correo');
        $sheet->setCellValue('C1', 'Documento');
        $sheet->setCellValue('D1', 'Telefono');
        $sheet->setCellValue('E1', 'Ciudad de residencia');
        $sheet->setCellValue('F1', 'Dirección');
        $sheet->setCellValue('G1', 'Genero');
        $sheet->setCellValue('H1', 'EPS');
        $sheet->setCellValue('I1', 'Tipo de sangre');
        $sheet->setCellValue('J1', 'Numero de emergencia');
        $sheet->setCellValue('K1', 'Fecha de nacimiento');
        $sheet->setCellValue('L1', 'Edad');
        $sheet->setCellValue('M1', 'Profesion');
        $sheet->setCellValue('N1', 'Mascotas');
        $sheet->setCellValue('O1', 'Transporte');
        $sheet->setCellValue('P1', 'Tipo de Camisa');
        $sheet->setCellValue('Q1', 'Talla');
        $sheet->setCellValue('R1', 'Nombre de la camisa');
        $sheet->setCellValue('S1', 'Nombre deseado');
        $sheet->setCellValue('T1', 'Grupo o club');
        $sheet->setCellValue('U1', 'Estado de salud');
        $sheet->setCellValue('V1', 'Autoriza patrocinadores');
        $sheet->setCellValue('W1', 'Tratamiento de datos');
    }

    // Obtener la última fila ocupada
    $lastRow = $sheet->getHighestRow() + 1;

    // Agregar los datos del usuario 
    $sheet->setCellValue('A' . $lastRow, $nombre);
    $sheet->setCellValue('B' . $lastRow, $correo);
    $sheet->setCellValue('C'. $lastRow, $dni);
    $sheet->setCellValue('D'. $lastRow, $celular);
    $sheet->setCellValue('E'. $lastRow, $residencia);
    $sheet->setCellValue('F'. $lastRow, $correspondencia);
    $sheet->setCellValue('G'. $lastRow, $genero);
    $sheet->setCellValue('H'. $lastRow, $eps);
    $sheet->setCellValue('I'. $lastRow, $rh);
    $sheet->setCellValue('J'. $lastRow, $numEmergencia);
    $sheet->setCellValue('K'. $lastRow, $nacimiento);
    $sheet->setCellValue('L'. $lastRow, $edad);
    $sheet->setCellValue('M'. $lastRow, $profesion);
    $sheet->setCellValue('N'. $lastRow, $mascotas);
    $sheet->setCellValue('O'. $lastRow, $transporte);
    $sheet->setCellValue('P'. $lastRow, $genCamisa);
    $sheet->setCellValue('Q'. $lastRow, $tallas);
    $sheet->setCellValue('R'. $lastRow, $nombrecamisa);
    $sheet->setCellValue('S'. $lastRow, $nombredeseado);
    $sheet->setCellValue('T'. $lastRow, $grupo);
    $sheet->setCellValue('U'. $lastRow, $terminos1);
    $sheet->setCellValue('V'. $lastRow, $autoriza);
    $sheet->setCellValue('W'. $lastRow, $terminos2);
    

    // Guardar el archivo 
    $writer = new Xlsx($spreadsheet);
    $writer->save($rutaArchivo);

    $url_excel = "./excel/archivo.xlsx";
    $img_png = "https://logodownload.org/wp-content/uploads/2020/04/excel-logo-0.png";
    // ---------------------------------------------------------------------------------
    //Enviar correo
    $mensaje_html = "
    <html>
    <head>
        <style>
            /* Estilos CSS para el correo electrónico */
            body {
                font-family: Arial, sans-serif;
            }
            .container {
                max-width: 600px;
                margin: 0 auto;
                padding: 20px;
                background-color: #f9f9f9;
                border-radius: 10px;
            }
            h2 {
                color: #333;
            h3 {
                color: #B44141;
            }
            .info {
                margin-bottom: 20px;
            }
            .info p {
                margin: 10px 0;
            }
        </style>
    </head>
    <body>
        <div class='container'>
            <h2>Nuevo envío de formulario</h2>
            <div class='info'> 
            <h3>Information Personal</h3>
                <p><strong>Nombre:</strong> $nombre</p>
                <p><strong>Género:</strong> $genero</p>
                <p><strong>Correo:</strong> $correo</p>
                <p><strong>DNI:</strong> $dni</p>
                <p><strong>Celular:</strong> $celular</p>
                <p><strong>Residencia:</strong> $residencia</p>
                <p><strong>Correspondencia:</strong> $correspondencia</p>
                <h3>Información Adicional</h3>
                <p><strong>EPS:</strong> $eps</p>
                <p><strong>RH:</strong> $rh</p>
                <p><strong>Número de emergencia:</strong> $numEmergencia</p>
                <h3>Otros datos</h3>
                <p><strong>Fecha de nacimiento:</strong> $nacimiento</p>
                <p><strong>Edad:</strong> $edad</p>
                <p><strong>Profesión:</strong> $profesion</p>
                <p><strong>Mascotas:</strong> $mascotas</p>
                <p><strong>Transporte:</strong> $transporte</p>
                <h3>Camisas, tallas & grupos</h3>
                <p><strong>Tipo de camisa:</strong> $genCamisa</p>
                <p><strong>Talla:</strong> $tallas</p>
                <p><strong>Nombre en la camisa?:</strong> $nombrecamisa</p>
                <p><strong>Nombre deseado:</strong> $nombredeseado</p>
                <p><strong>Grupo o club:</strong> $grupo</p>
                <h3>Autorizacion de datos y otros</h3>
                <p><strong>Estado de salud</strong> $terminos1</p>
                <p><strong>Autoriza compartir correo y numero con patrocinadores:</strong> $autoriza</p>
                <p><strong>Terminos y Condiciones:</strong> $terminos2</p>
                <br>
                <br>
                <br>
                <br>
                <h2>Puedes consultar el excel en este enlace</h2>
                <p></p>
                <a href=`$url_excel`><h3>Descargar excel</h3></a>
                
            </div>
        </div>
    </body>
    </html>
    ";
    
    

    $destinatario = "ejemplo@correo.com"; //correo destino.
    $asunto = "Nuevo envio de formulario de $nombre";
    $cabeceras = "MIME-Version: 1.0" . "\r\n";
    $cabeceras .= "Content-type:text/html;charset=UTF-8" . "\r\n";
    $cabeceras .= "From: $correo";

    // Enviar el correo electrónico
    mail($destinatario, $asunto, $mensaje_html, $cabeceras);

    // Redirigir 
    header("Location: sucess.html");
    exit();
}
?>