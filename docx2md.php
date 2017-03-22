#!/usr/bin/php
<?php
if (PHP_SAPI !== 'cli') {
	die('This script is meant to be run in command-line mode only.');
}

const VERSION  = '1.0';
const ENCODING = 'UTF-8';

/**
 * Convert a .docx file to markdown using the given command-line arguments
 *
 * @param  array   $args
 * @param  boolean $isTestMode
 * @return string
 */
function docx2md(array $args, $isTestMode = false) {
	$debugDumpWord         = false;
	$debugDumpIntermediary = false;

	// Set command-line to use utf-8
	shell_exec('chcp 65001');

	// Check command line options
	$longOptionsArray = array('image', 'test');

	$shortOptionsArray = array_map(function ($value) {
		return substr($value, 0, 1);
	}, $longOptionsArray);
	$shortOptions = implode($shortOptionsArray);

	$options = getopt($shortOptions, $longOptionsArray);

	foreach ($longOptionsArray as $index => $longOption) {
		${'option' . ucfirst($longOption)} = false;
		${$longOption . 'Options'} = array($shortOptionsArray[$index], $longOptionsArray[$index], "-{$shortOptionsArray[$index]}", "--{$longOptionsArray[$index]}");

		if (array_key_exists(${$longOption . 'Options'}[0], $options) || array_key_exists(${$longOption . 'Options'}[1], $options) || array_intersect($args, ${$longOption . 'Options'})) {
			${'option' . ucfirst($longOption)} = true;
		}
	}

	// Remove first argument: the script name
	$args = array_slice($args, 1);
	// Remove all options from the list of arguments
	$args = array_filter($args, function ($value) {
		return (substr($value, 0, 1) === '-') === false;
	});
	// Re-index the array
	$args = array_values($args);


	if (count($args) <= 0) {
		// If option is set and not already in test mode
		// run tests and *stop*
		if ($optionTest && !$isTestMode) {
			return runTests($args);
		}

		$output  = 'Converts Micro$oft Word .docx files to Markdown format.' . PHP_EOL;
		$output .= 'docx2md is written by Mathieu Bouchard (@matb33).' . PHP_EOL;
		$output .= PHP_EOL;
		$output .= 'Usage: php ./docx2md.php [options] [source.docx|path/to/dir] [destination.md|path/to/dir]' . PHP_EOL;
		$output .= PHP_EOL;
		$output .= 'Options:';
		$output .= PHP_EOL;
		$output .= '  -i, --image Parse images';
		$output .= PHP_EOL;
		$output .= '  -t, --test  Run tests';
		$output .= PHP_EOL;
		$output .= PHP_EOL;
		$output .= 'If no destination file is specified, output will be written to the console excluding any images.';
		$output .= PHP_EOL;
		die($output);
	} else if (!$debugDumpWord && !$debugDumpIntermediary) {
		// If option is set and not already in test mode
		// run tests and *continue on* with converting
		if ($optionTest && !$isTestMode) {
			runTests($args);
		}
	}

	//==========================================================================
	// Extract command-line parameters
	//==========================================================================

	$docxFilename = null;
	$mdFilename   = null;

	if (array_key_exists(0, $args) && $args[0] !== '') {
		$docxFilename = $args[0];
	}

	if (array_key_exists(1, $args) && $args[1] !== '') {
		$mdFilename = $args[1];
	}

	if (!file_exists($docxFilename)) {
		die("Input .docx file/directory does not exist: \"{$docxFilename}\"");
	} else {
		$docxFilename = realpath($docxFilename);
	}

	$isMultipleFiles = false;
	if (is_dir($docxFilename)) {
		$isMultipleFiles = true;
		$sourceFiles = glob("{$docxFilename}\\*.docx");
		$destination = realpath($mdFilename) . '\\';
	} else {
		$sourceFiles = glob($docxFilename);
		$destination = realpath(dirname($docxFilename)) . '\\';
	}

	foreach ($sourceFiles as $index => $docxFilename) {
		if (!$isTestMode && $mdFilename !== null) {
			if ($isMultipleFiles) {
				$mdFilename = basename($docxFilename, 'docx') . 'md';
			} else if(file_exists($mdFilename)) {
				// Generate a random extension so as not to overwrite destination filename
				$mdFilename = $mdFilename . '.' . substr(md5(uniqid(rand(), true)), 0, 5);
			}
		}

		//==========================================================================
		// Step 1: Extract Word doc to a temporary location and delete relevant images
		//==========================================================================

		$documentFolder = sys_get_temp_dir() . '/' . md5($docxFilename);

		if (file_exists($documentFolder)) {
			rrmdir($documentFolder);
			mkdir($documentFolder);
		}

		if ($optionImage) {
			if ($isTestMode) {
				$imageFolder = 'images';
			} else {
				$imageFolder = $destination . 'images';
				if (file_exists($imageFolder) && is_dir($imageFolder)) {
					// Clean-up existing images only associated with the defined markdown file
					$images = glob("{$imageFolder}/" . basename($mdFilename, '.md') . '.*.{bmp,gif,jpg,jpeg,png}', GLOB_BRACE);
					foreach ($images as $image) {
						if (is_file($image)) {
							unlink($image);
						}
					}
				} else {
					mkdir($imageFolder, 0777, true);
				}
			}
		}

		$zip = new ZipArchive;
		$res = $zip->open($docxFilename);

		if ($res === true) {
			if (!$isTestMode && $optionImage) {
				extractFolder($zip, 'word/media', $documentFolder, $imageFolder, $mdFilename);
			} else {
				extractFolder($zip, 'word/media', $documentFolder);
			}
			$zip->extractTo($documentFolder, array('word/document.xml', 'word/_rels/document.xml.rels', 'docProps/core.xml'));
			$zip->close();
		} else {
			die("The .docx file appears to be corrupt (i.e. it can't be opened using Zip). Please try re-saving your document and re-uploading, or ensuring that you are providing a valid .docx file.");
		}

		//==========================================================================
		// Step 2: Read the main document.xml and also bring in the rels document
		//==========================================================================

		$wordDocument = new DOMDocument(VERSION, ENCODING);
		$wordDocument->load($documentFolder . '/word/document.xml');

		$wordDocumentRels = new DOMDocument(VERSION, ENCODING);
		$wordDocumentRels->load($documentFolder . '/word/_rels/document.xml.rels');
		$wordDocument->documentElement->appendChild($wordDocument->importNode($wordDocumentRels->documentElement, true));

		$xml = $wordDocument->saveXML();

		// libxml < 2.7 fix
		$xml = str_replace('r:id=', 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id=', $xml);
		$xml = str_replace('r:embed=', 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed=', $xml);

		$mainDocument = new DOMDocument(VERSION, ENCODING);
		$mainDocument->loadXML($xml);

		if ($debugDumpWord) {
			$mainDocument->preserveWhiteSpace = false;
			$mainDocument->formatOutput = true;
			die($mainDocument->saveXML());
		}

		//==========================================================================
		// Step 3: Convert the bulk of the docx XML to an intermediary format
		//==========================================================================

		$xslDocument = new DOMDocument(VERSION, ENCODING);
		$xslDocument->loadXML(DOCX_TO_INTERMEDIARY_TRANSFORM);

		$processor = new XSLTProcessor();
		$processor->importStyleSheet($xslDocument);
		$intermediaryDocument = $processor->transformToDoc($mainDocument);

		//==========================================================================
		// Step 4: Use string functions to trim away unwanted whitespace in
		// specific places. Use DOMXPath to iterate through specific tags and
		// clean the data
		//==========================================================================

		$xml = $intermediaryDocument->saveXML();

		$displayTags = array('i:para', 'i:heading', 'i:listitem');

		foreach ($displayTags as $tag) {
			// Remove any number of spaces that follow the opening tag
			$xml = preg_replace("/(<{$tag}[^>]*>)[ ]*/", ' \\1', $xml);

			// Remove multiple spaces before closing tags
			$xml = preg_replace("/[ ]*<\/{$tag}>/", "</{$tag}>", $xml);
		}

		$formattingTags = array('i:bold', 'i:italic', 'i:strikethrough', 'i:line');

		foreach ($formattingTags as $tag) {
			// Convert parallel repeated tags to single instance
			// e.g. `<i:x>foo</i:x><i:x>bar</i:x>` to `<i:x>foo bar</i:x>`
			$xml = preg_replace("/(<\/{$tag}>)[ ]*<{$tag}>/", ' ', $xml);

			// Remove any number of spaces that follow the opening tag
			$xml = preg_replace("/(<{$tag}[^>]*>)[ ]*/", ' \\1', $xml);

			// Remove multiple spaces before closing tags
			$xml = preg_replace("/[ ]*<\/{$tag}>/", "</{$tag}>", $xml);
		}

		// Remove white spaces between tags
		$xml = preg_replace('/(\>)\s*(\<)/m', '$1$2', $xml);

		$intermediaryDocument->loadXML($xml);

		// Remove empty tags
		$xpath = new DOMXPath($intermediaryDocument);
		while (($nodes = $xpath->query('//*[not(*) and not(\'i:image\') and not(text()[normalize-space()])]')) && ($nodes->length)) {
			foreach ($nodes as $node) {
				$node->parentNode->removeChild($node);
			}
		}

		$allTags = array_merge($displayTags, $formattingTags);

		foreach ($allTags as $tag) {
			foreach ($xpath->query("//{$tag}/text()") as $textNode) {
				$output = $textNode->nodeValue;

				// Cleanse data
				$output = cleanData($output);

				// Replace multiple spaces with a single space
				$output = preg_replace('!\s+!', ' ', $output);

				// Remove spaces preceding punctuation
				$output = preg_replace('/\s*([\.,\?\!])/', '\\1', $output);

				// Escape existing chars used in markdown as formatting
				$output = addcslashes($output, '*_~`');

				// Assign result
				$textNode->nodeValue = $output;
			}
		}

		if ($debugDumpIntermediary) {
			$intermediaryDocument->preserveWhiteSpace = false;
			$intermediaryDocument->formatOutput = true;
			die($intermediaryDocument->saveXML());
		}

		//==========================================================================
		// Step 5: Convert from the intermediary XML format to Markdown
		//==========================================================================

		$xslDocument = new DOMDocument(VERSION, ENCODING);
		if ($optionImage) {
			// Replace image placeholder with image template
			$imageFilename = ($mdFilename) ? basename($mdFilename, '.md') . '.' : null;
			$imageTemplate = sprintf(IMAGE_TEMPLATE, $imageFolder, $imageFilename);
			$xslDocument->loadXML(sprintf(INTERMEDIARY_TO_MARKDOWN_TRANSFORM, $imageTemplate));
		} else {
			// Replace image placeholder with a blank string
			$xslDocument->loadXML(sprintf(INTERMEDIARY_TO_MARKDOWN_TRANSFORM, ''));
		}

		$processor = new XSLTProcessor();
		$processor->importStyleSheet($xslDocument);
		$markdown = $processor->transformToXml($intermediaryDocument);
		$markdown = rtrim(join(PHP_EOL, array_map('rtrim', explode("\n", $markdown))));

		//==========================================================================
		// Step 6: If the Markdown output file was specified, write it. Otherwise
		// just write to STDOUT (echo)
		//==========================================================================

		if (!$isTestMode) {
			if (!$isMultipleFiles || $index === 0) {
				$formatter       = "%s \033[32m" . html_entity_decode('&radic;') . " \033[0m  ";
				$completeMessage = 'Performing conversion... finished';

				if ($optionImage) {
					$completeMessage .= ' with images included';
				}

				echo sprintf($formatter, $completeMessage);
			}

			if ($mdFilename !== null) {
				file_put_contents("{$destination}" . $mdFilename, $markdown);
				echo PHP_EOL;

				if ($isMultipleFiles) {
					$index++;
					echo " {$index}.";
				}

				echo ' Created: "' . basename($mdFilename) . '"';
			} else {
				echo PHP_EOL . PHP_EOL;
				echo 'Markdown:' . PHP_EOL;
				echo str_repeat('-', 9) . PHP_EOL;
				echo $markdown;
			}
		}

		//==========================================================================
		// Step 7: Clean-up
		//==========================================================================

		if (file_exists($documentFolder)) {
			rrmdir($documentFolder);
		}
	}

	return $markdown;
}

//==============================================================================
// Helper functions
//==============================================================================

/**
 * Extract the files from a given zipped folder.
 * Optionally this includes all images.
 *
 * @param  string $zip
 * @param  string $folderName
 * @param  string $destination
 * @param  string $imageFolder
 * @param  string $mdFilename
 * @return void
 */
function extractFolder($zip, $folderName, $destination, $imageFolder = null, $mdFilename = null) {
	for ($i = 0; $i < $zip->numFiles; $i++) {
		$fileName = $zip->getNameIndex($i);

		if (strpos($fileName, $folderName) !== false) {
			if (!is_null($imageFolder) && !is_null($mdFilename)) {
				// Save matching images to disk
				if (preg_match('([^\s]+(\.(?i)(bmp|gif|jpe?g|png))$)', $fileName)) {
					file_put_contents("{$imageFolder}/" . basename($mdFilename, '.md') . '.' . basename($fileName), $zip->getFromIndex($i));
				}
			}

			$zip->extractTo($destination, $fileName);
		}
	}
}

/**
 * Recursively remove directories.
 *
 * @param  string $directory
 * @return void
 */
function rrmdir($directory) {
	foreach (glob($directory . '/*') as $file) {
		if (is_dir($file)) {
			rrmdir($file);
		} else {
			unlink($file);
		}
	}

	rmdir($directory);
}

/**
 * Replace curly quotes and other special characters
 * with their standard equivalents.
 *
 * @param  string $data
 * @return string
 */
function cleanData($data){
	$replacementChars = array("'", "'", '"', '"', '-', '--', '...', ' ');

	// Replace UTF-8 characters
	$cleanedData = str_replace(array("\xe2\x80\x98", "\xe2\x80\x99", "\xe2\x80\x9c", "\xe2\x80\x9d", "\xe2\x80\x93", "\xe2\x80\x94", "\xe2\x80\xa6", "\xc2\xa0"), $replacementChars, $data);

	// Replace Windows-1252 equivalents
	// Replacement of horizontal ellipsis - chr(133) - as it interferes with 'A with circle' char
	$cleanedData = str_replace(array(chr(145), chr(146), chr(147), chr(148), chr(150), chr(151), chr(133), chr(194)), $replacementChars, $cleanedData);

	return $cleanedData;
}

/**
 * Test markdown converter
 *
 * @param $args
 * @return void
 */
function runTests($args){
	$src       = 'examples';
	$formatter = " %s. %s\033[0m: %s" . PHP_EOL;
	$formatter = " %s. %s\033[0m: %s" . PHP_EOL;
	$output    = "\033[0m";

	echo 'Running tests...';

	$files = glob("{$src}/docx/*.docx");
	$size  = sizeof($files);
	$charCount = 0;

	foreach ($files as $n => $file1) {
		$n++;
		$file2 = basename($file1, '.docx') . '.md';

		$markdown = docx2md(array('', '-i', $file1, $file2), true);
		$md       = "{$src}/md/{$file2}";

		$fileHash1 = sha1(preg_replace('/\v+/', PHP_EOL . PHP_EOL, $markdown));
		$fileHash2 = sha1(preg_replace('/\v+/', PHP_EOL . PHP_EOL, file_get_contents($md)));

		// Padding required on the last line to prevent
		// miscellaneous characters printing to the console
		if ($n === $size) {
			$size++;
			$file1 .= str_repeat(' ', ($size * 2));
		}

		if ($fileHash1 === $fileHash2) {
			$sprintf = sprintf($formatter, $n, "\033[32mPassed " . html_entity_decode('&radic;'), $file1);
		} else {
			$sprintf = sprintf($formatter, $n, "\033[31mFailed " . html_entity_decode('&times;'), $file1);
		}

		$charCount = strlen(rtrim($sprintf));
		$output   .= $sprintf;
	}

	echo ' finished' . " \033[32m" . html_entity_decode('&radic;') . " \033[0m" . PHP_EOL . rtrim($output, PHP_EOL);

	if ($args) {
		// If performing conversion after running tests, print a separator
		echo PHP_EOL . str_repeat('_', $charCount) . PHP_EOL . PHP_EOL;
	}
}

//==============================================================================
// XSL Stylesheets
//==============================================================================

define('DOCX_TO_INTERMEDIARY_TRANSFORM', <<<'XML'
<?xml version="1.0"?>
<xsl:stylesheet version="1.0"
	xmlns:i="urn:docx2md:intermediary"
	xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
	xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
	xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
	xmlns:rels="http://schemas.openxmlformats.org/package/2006/relationships"
	xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
	xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
	xmlns:xsl="http://www.w3.org/1999/XSL/Transform">

	<xsl:template match="/w:document">
		<i:document>
			<xsl:apply-templates />
		</i:document>
	</xsl:template>

	<xsl:template match="w:body">
		<i:body>
			<xsl:apply-templates />
		</i:body>
	</xsl:template>

	<xsl:template match="rels:Relationships" />

	<!-- Heading styles -->
	<xsl:template match="w:p[w:pPr/w:pStyle/@w:val[starts-with(., 'Heading')]]">
		<xsl:variable name="style" select="w:pPr/w:pStyle/@w:val[starts-with(., 'Heading')]" />
		<xsl:variable name="level" select="substring($style, 8, 1)" />
		<xsl:variable name="type" select="translate(substring($style, 9), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz')" />
		<xsl:if test="count(w:r)">
			<i:heading>
				<xsl:attribute name="level">
					<xsl:value-of select="$level" />
				</xsl:attribute>
				<xsl:if test="$type != ''">
					<xsl:attribute name="type">
						<xsl:value-of select="$type" />
					</xsl:attribute>
				</xsl:if>
				<xsl:apply-templates />
			</i:heading>
		</xsl:if>
	</xsl:template>

	<!-- Regular paragraph style -->
	<xsl:template match="w:p">
		<xsl:if test="count(w:r)">
			<i:para>
				<xsl:apply-templates />
			</i:para>
		</xsl:if>
		<!-- Horizontal line -->
		<xsl:if test="count(w:pPr/w:pBdr)">
			<i:line>---</i:line>
		</xsl:if>
	</xsl:template>

	<!-- Table -->
	<xsl:template match="w:tbl">
		<i:table>
			<xsl:apply-templates />
		</i:table>
	</xsl:template>
	<!-- Table: row -->
	<xsl:template match="w:tbl/w:tr">
		<xsl:if test="count(w:tc) and position() &lt; 4">
			<i:header>
				<xsl:apply-templates />
			</i:header>
		</xsl:if>
		<xsl:if test="count(w:tc) and position() &gt; 3">
			<i:row>
				<xsl:apply-templates />
			</i:row>
		</xsl:if>
	</xsl:template>
	<!-- Table: cell -->
	<xsl:template match="w:tbl/w:tr/w:tc">
		<xsl:if test="count(w:p/w:r/w:t)">
			<i:cell>
				<xsl:apply-templates />
			</i:cell>
		</xsl:if>

		<!-- Table: blank cells -->
		<xsl:if test="count(w:p/w:r/w:t) &lt; 1">
			<i:cell>-</i:cell>
		</xsl:if>
	</xsl:template>

	<!-- List items -->
	<xsl:template match="w:p[w:pPr/w:numPr]">
		<xsl:if test="count(w:r)">
			<i:listitem level="{w:pPr/w:numPr/w:ilvl/@w:val}" type="{w:pPr/w:numPr/w:numId/@w:val}">
				<xsl:apply-templates />
			</i:listitem>
		</xsl:if>
	</xsl:template>
	<xsl:template match="w:p[w:pPr/w:pStyle/@w:val = 'ListBullet']">
		<xsl:if test="count(w:r)">
			<i:listitem level="0" type="1">
				<xsl:apply-templates />
			</i:listitem>
		</xsl:if>
	</xsl:template>
	<xsl:template match="w:p[w:pPr/w:pStyle/@w:val = 'ListNumber']">
		<xsl:if test="count(w:r)">
			<i:listitem level="0" type="2">
				<xsl:apply-templates />
			</i:listitem>
		</xsl:if>
	</xsl:template>

	<!-- Text content -->
	<xsl:template match="w:r">
		<xsl:apply-templates />
	</xsl:template>
	<xsl:template match="w:t">
		<!-- Normal -->
		<xsl:value-of select="." />
	</xsl:template>
	<xsl:template match="w:r[w:rPr/w:b and not(w:rPr/w:i)]/w:t">
		<!-- Bold -->
		<i:bold>
			<xsl:value-of select="." />
		</i:bold>
	</xsl:template>
	<xsl:template match="w:r[w:rPr/w:i and not(w:rPr/w:b)]/w:t">
		<!-- Italic -->
		<i:italic>
			<xsl:value-of select="." />
		</i:italic>
	</xsl:template>
	<xsl:template match="w:r[w:rPr/w:b and w:rPr/w:i]/w:t">
		<!-- Bold + Italic -->
		<i:bold>
			<i:italic>
				<xsl:value-of select="." />
			</i:italic>
		</i:bold>
	</xsl:template>
	<xsl:template match="w:r[w:rPr/w:strike]/w:t">
		<!-- Strikethrough -->
		<i:strikethrough>
			<xsl:value-of select="." />
		</i:strikethrough>
	</xsl:template>
	<xsl:template match="w:br">
		<i:linebreak />
	</xsl:template>

	<!-- Hyperlinks -->
	<xsl:template match="w:p[w:hyperlink]">
		<xsl:variable name="id" select="w:hyperlink/@r:id" />
		<xsl:if test="count(w:hyperlink/w:r)">
			<i:link>
				<xsl:attribute name="href">
					<xsl:value-of select="/w:document/rels:Relationships/rels:Relationship[@Id=$id]/@Target" />
				</xsl:attribute>
				<xsl:if test="/w:document/rels:Relationships/rels:Relationship[@Id=$id]/@TargetMode">
					<xsl:attribute name="target">
						<xsl:value-of select="/w:document/rels:Relationships/rels:Relationship[@Id=$id]/@TargetMode" />
					</xsl:attribute>
				</xsl:if>
				<xsl:apply-templates />
			</i:link>
		</xsl:if>
	</xsl:template>

	<!-- Images -->
	<xsl:template match="w:drawing">
		<xsl:apply-templates select=".//a:blip" />
	</xsl:template>
	<xsl:template match="a:blip">
		<xsl:variable name="id" select="@r:embed" />
		<i:image>
			<xsl:attribute name="src">
				<xsl:value-of select="/w:document/data/@word-folder" />
				<xsl:value-of select="/w:document/rels:Relationships/rels:Relationship[@Id=$id]/@Target" />
			</xsl:attribute>
			<xsl:attribute name="width">
				<xsl:value-of select="round(ancestor::w:drawing[1]//wp:extent/@cx div 9525)" />
			</xsl:attribute>
			<xsl:attribute name="height">
				<xsl:value-of select="round(ancestor::w:drawing[1]//wp:extent/@cy div 9525)" />
			</xsl:attribute>
		</i:image>
	</xsl:template>

	<!-- Edit: Inserted text -->
	<xsl:template match="w:ins">
		<xsl:apply-templates />
	</xsl:template>

	<!-- Edit: Deleted text -->
	<xsl:template match="w:del" />
</xsl:stylesheet>
XML
);

define('INTERMEDIARY_TO_MARKDOWN_TRANSFORM', <<<'XML'
<?xml version="1.0"?>
<xsl:stylesheet version="1.0"
	xmlns:i="urn:docx2md:intermediary"
	xmlns:str="http://exslt.org/strings" extension-element-prefixes="str"
	xmlns:xsl="http://www.w3.org/1999/XSL/Transform">

	<xsl:output
		media-type="text/plain"
		method="text"
		omit-xml-declaration="yes"
	/>

	<xsl:template match="@*|node()">
		<xsl:copy>
			<xsl:apply-templates select="@*|node()"/>
		</xsl:copy>
	</xsl:template>

	<xsl:template match="i:document">
		<xsl:apply-templates />
		<xsl:text>&#xa;</xsl:text>
		<xsl:for-each select="//i:link">
			<xsl:text>&#32;&#32;[</xsl:text>
			<xsl:value-of select="position()" />
			<xsl:text>]:&#32;</xsl:text>
			<xsl:value-of select="@href" />
			<xsl:text>&#xa;</xsl:text>
		</xsl:for-each>
	</xsl:template>

	<xsl:template match="i:body">
		<xsl:apply-templates />
	</xsl:template>

	<xsl:template match="i:heading">
		<xsl:value-of select="substring('######', 1, @level)" />
		<xsl:text>&#32;</xsl:text>
		<xsl:apply-templates />
		<xsl:text>&#xa;&#xa;</xsl:text>
	</xsl:template>

	<xsl:template match="i:link">
		<xsl:text>[</xsl:text>
		<xsl:value-of select="." />
		<xsl:text>][</xsl:text>
		<xsl:value-of select="count(preceding::i:link) + 1" />
		<xsl:text>]</xsl:text>
	</xsl:template>

	<xsl:template match="i:italic">
		<xsl:text>_</xsl:text>
		<xsl:apply-templates />
		<xsl:text>_</xsl:text>
	</xsl:template>

	<xsl:template match="i:bold">
		<xsl:text>**</xsl:text>
		<xsl:apply-templates />
		<xsl:text>**</xsl:text>
	</xsl:template>

	<xsl:template match="i:strikethrough">
		<xsl:text>~~</xsl:text>
		<xsl:apply-templates />
		<xsl:text>~~</xsl:text>
	</xsl:template>

	<xsl:template match="i:para">
		<xsl:if test="./* or text() != ''">
			<xsl:apply-templates />
			<xsl:if test="not(parent::i:cell)">
				<xsl:text>&#xa;&#xa;</xsl:text>
			</xsl:if>
		</xsl:if>
	</xsl:template>

	<xsl:template match="i:line">
		<xsl:text>---&#xa;&#xa;</xsl:text>
	</xsl:template>

	<xsl:template match="i:linebreak">
		<xsl:text>&#xa;</xsl:text>
	</xsl:template>

	<xsl:template match="i:table">
		<xsl:apply-templates />
		<xsl:text>&#xa;&#xa;</xsl:text>
	</xsl:template>
	<xsl:template match="i:header">
		<xsl:apply-templates />
		<xsl:variable name="count" select="count(../i:row/i:cell)" />
		<xsl:if test="$count &gt; 0">
			<xsl:text>&#xa;| </xsl:text>
			<xsl:call-template name="string-repeat">
				<xsl:with-param name="string" select="'--- | '" />
				<xsl:with-param name="times" select="count(i:cell)" />
			</xsl:call-template>
		</xsl:if>
	</xsl:template>
	<xsl:template match="i:row">
		<xsl:text>&#xa;</xsl:text>
		<xsl:apply-templates />
	</xsl:template>
	<xsl:template match="i:cell">
		<xsl:variable name="count" select="count(../../i:row/i:cell)" />
		<xsl:if test="$count = 0">
			<xsl:apply-templates />
			<xsl:text>&#xa;&#xa;</xsl:text>
		</xsl:if>
		<xsl:if test="$count &gt; 0">
			<xsl:if test="position() = 1">
				<xsl:text>| </xsl:text>
			</xsl:if>
			<xsl:apply-templates />
			<xsl:text> | </xsl:text>
		</xsl:if>
	</xsl:template>

	<!-- Bulleted list-item -->
	<xsl:template match="i:listitem[@type!='2']">
		<xsl:value-of select="substring('		  ', 1, @level * 2)" />
		<xsl:text> - </xsl:text>
		<xsl:apply-templates />
		<xsl:text>&#xa;</xsl:text>
		<xsl:if test="local-name(following-sibling::i:*[1]) != 'listitem'"
			><xsl:text>&#xa;</xsl:text>
		</xsl:if>
	</xsl:template>

	<!-- Numbered list-item -->
	<xsl:template match="i:listitem[@type='2']">
		<xsl:variable name="level" select="@level" />
		<xsl:variable name="type" select="@type" />
		<xsl:value-of select="substring('		  ', 1, $level * 2)" />
		<xsl:text> 1. </xsl:text>
		<xsl:apply-templates />
		<xsl:text>&#xa;</xsl:text>
		<xsl:if test="local-name(following-sibling::i:*[1]) != 'listitem'">
			<xsl:text>&#xa;</xsl:text>
		</xsl:if>
	</xsl:template>

	<!-- Image Template Placeholder -->
	%s

	<!-- Escape asterix -->
	<xsl:template match="text()">
		<xsl:call-template name="string-replace-all">
			<xsl:with-param name="text" select="." />
			<xsl:with-param name="replace" select="'*'" />
			<xsl:with-param name="by" select="'\*'" />
		</xsl:call-template>
	</xsl:template>

	<!-- Superscript ® -->
	<xsl:template match="text()">
		<xsl:call-template name="string-replace-all">
			<xsl:with-param name="text" select="." />
			<xsl:with-param name="replace" select="'®'" />
			<xsl:with-param name="by" select="'&lt;sup&gt;®&lt;/sup&gt;'" />
		</xsl:call-template>
	</xsl:template>

	<!-- Helper: string replace -->
	<xsl:template name="string-replace-all">
		<xsl:param name="text" />
		<xsl:param name="replace" />
		<xsl:param name="by" />
		<xsl:choose>
			<xsl:when test="contains($text, $replace)">
				<xsl:value-of select="substring-before($text, $replace)" />
				<xsl:value-of select="$by" />

				<xsl:call-template name="string-replace-all">
					<xsl:with-param name="text" select="substring-after($text, $replace)" />
					<xsl:with-param name="replace" select="$replace" />
					<xsl:with-param name="by" select="$by" />
				</xsl:call-template>
			</xsl:when>
			<xsl:otherwise>
				<xsl:value-of select="$text" />
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>

	<!-- Helper: string repeat -->
	<xsl:template name="string-repeat">
		<xsl:param name="string" select="''" />
		<xsl:param name="times" select="1" />

		<xsl:if test="number($times) &gt; 0">
			<xsl:value-of select="$string" />
			<xsl:call-template name="string-repeat">
				<xsl:with-param name="string" select="$string" />
				<xsl:with-param name="times" select="$times - 1" />
			</xsl:call-template>
		</xsl:if>
	</xsl:template>
</xsl:stylesheet>
XML
);

define('IMAGE_TEMPLATE', '<!-- Image --><xsl:template match="i:image"><xsl:text>![Image](%s\%s</xsl:text><xsl:value-of select="str:tokenize(@src, \'/\')[last()]" /><xsl:text>)&#xa;</xsl:text></xsl:template>');

docx2md($argv);