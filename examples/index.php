<?php
    require_once '..\docx2md.php';

    $converter = new Docx2md\Docx2md;

    $files = glob('docx/*.docx', GLOB_BRACE);

    foreach ($files as $file) {
        $result = $converter->parseFile($file);

        print_r($result->metadata);

        echo '<br><br>';
        echo nl2br($result->markdown, false);
        echo '<br><br>';
    }
