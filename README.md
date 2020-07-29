# docx2md

## Convert a `.docx` file to markdown using the given command line arguments

- This is a simple converter that will read a `.docx` file and convert it into markdown format.
  - It works on the command line or included as a standalone script.

### On the command line

- Run `php docx2md.php` to see a list of available options and how to use the converter

### As a standalone script

```php
<?php

    require_once 'docx2md.php';

    $converter = new Docx2md\Docx2md();
    $converter = $converter->parseFile('word.docx');
    $markdown  = $converter->markdown;

    echo $markdown;
```

## Background

- This was put together based on code snippets from a closed-source project written in 2011.
- It would be trivial to extend this script to work on a webserver baring in mind the management of file uploads may need to be considered.
- Since most of the conversion work is left to XSL transformations, this utility could be very easily ported to another language other than PHP.
  - PHP was chosen because existing code snippets were already written in the language.
- I look forward to any pull requests or bug reports to help make this converter as reliable as possible.