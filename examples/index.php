<?php
    require_once '..\docx2md.php';

    $converter = new Docx2md\Docx2md;

    $files = glob('docx/*.docx', GLOB_BRACE);
    foreach ($files as $file) {
        echo $converter->parseFile($file) . '<br><br>';
    }