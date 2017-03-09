#!/usr/bin/php
<?php

function docx2md($args) {
	$debugDumpWord = false;
	$debugDumpIntermediary = false;

	if (PHP_SAPI !== "cli") {
		die("This script is meant to run in command-line mode only.\n");
	}

	// Check command line options for including images: -i, --image
	$imageLongOption = array('image');
	$imageShortOption = $imageLongOption[0][0]; // First letter of *i*mage
	$options = getopt($imageShortOption, $imageLongOption);

	$includeImages = false;
	if (array_key_exists($imageShortOption, $options) || array_key_exists($imageLongOption[0], $options)) {
		$includeImages = true;
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
		echo "docx2md: Written by Mathieu Bouchard @matb33\n";
		echo "Converts MS Word .docx files to Markdown format.\n";
		echo "\n";
		echo "Usage: php ./docx2md.php [options] input.docx [output.md]\n";
		echo "\n";
		echo 'Options:';
		echo "\n";
		echo '  -i, --image Include images';
		echo "\n";
		echo "\n";
		echo 'If no output file is specified, writes to STDOUT excluding images.';
		echo "\n";
		exit();
	}

	//==========================================================================
	// Extract command-line parameters
	//==========================================================================

	$docxFilename = null;
	$mdFilename = null;

	if (array_key_exists(0, $args) && $args[0] !== "") {
		$docxFilename = $args[0];
	}

	if (array_key_exists(1, $args) && $args[1] !== "") {
		$mdFilename = $args[1];
	} else {
		// Override including images as no output filename has been provided
		$includeImages = false;
	}

	if (!file_exists($docxFilename)) {
		die("Input docx file does not exist: " . $docxFilename . "\n");
	} else {
		$docxFilename = realpath($docxFilename);
	}

	// Generate a random extension so as not to overwrite destination filename
	if ($mdFilename !== null && file_exists($mdFilename)) {
		$mdFilename = $mdFilename . "." . substr(md5(uniqid(rand(), true)), 0, 5);
	}

	//==========================================================================
	// Step 1: Extract Word doc to a temporary location and delete relevant images
	//==========================================================================

	$documentFolder = sys_get_temp_dir() . "/" . md5($docxFilename);

	if (file_exists($documentFolder)) {
		rrmdir($documentFolder);
		mkdir($documentFolder);
	}

	$imageFolder = 'images';

	if ($includeImages) {
		// Clean-up existing images only associated with the defined markdown file
		$images = glob("$imageFolder/" . basename($mdFilename, '.md') . '.*.{bmp,gif,jpg,jpeg,png}', GLOB_BRACE);
		foreach ($images as $image) {
			if (is_file($image)) {
				unlink($image);
			}
		}
	}

	$zip = new ZipArchive;
	$res = $zip->open($docxFilename);

	if ($res === true) {
		if ($includeImages) {
			extractFolder($zip, "word/media", $documentFolder, $imageFolder, $mdFilename);
		} else {
			extractFolder($zip, "word/media", $documentFolder);
		}
		$zip->extractTo($documentFolder, array("word/document.xml", "word/_rels/document.xml.rels", "docProps/core.xml"));
		$zip->close();
	} else {
		die("The docx file appears to be corrupt (i.e. it can't be opened using Zip).  Please try re-saving your document and re-uploading, or ensuring that you are providing a valid docx file.\n");
	}

	//==========================================================================
	// Step 2: Read the main document.xml and also bring in the rels document
	//==========================================================================

	$wordDocument = new DOMDocument("1.0", "UTF-8");
	$wordDocument->load($documentFolder . "/word/document.xml");

	$wordDocumentRels = new DOMDocument("1.0", "UTF-8");
	$wordDocumentRels->load($documentFolder . "/word/_rels/document.xml.rels");
	$wordDocument->documentElement->appendChild($wordDocument->importNode($wordDocumentRels->documentElement, true));

	$xml = $wordDocument->saveXML();

	// libxml < 2.7 fix
	$xml = str_replace("r:id=", "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=", $xml);
	$xml = str_replace("r:embed=", "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:embed=", $xml);

	$mainDocument = new DOMDocument("1.0", "UTF-8");
	$mainDocument->loadXML($xml);

	if ($debugDumpWord) {
		$mainDocument->preserveWhiteSpace = false;
		$mainDocument->formatOutput = true;
		echo $mainDocument->saveXML();
		exit();
	}

	//==========================================================================
	// Step 3: Convert the bulk of the docx XML to an intermediary format
	//==========================================================================

	$xslDocument = new DOMDocument("1.0", "UTF-8");
	$xslDocument->loadXML(DOCX_TO_INTERMEDIARY_TRANSFORM);

	$processor = new XSLTProcessor();
	$processor->importStyleSheet($xslDocument);
	$intermediaryDocument = $processor->transformToDoc($mainDocument);

	//==========================================================================
	// Step 4: Use string functions to trim away unwanted whitespace in very
	// specific places. We do this using string manipulation for increased
	// control over exactly how we target and trim this whitespace:
	//==========================================================================

	$xml = $intermediaryDocument->saveXML();

	$tags = array("i:para", "i:heading", "i:listitem");

	foreach ($tags as $tag) {
		// Remove any number of spaces that follow the opening tag
		$xml = preg_replace("/(<{$tag}[^>]*>)[ ]*/", "\\1", $xml);

		// Remove a single space that precedes the closing tag
		$xml = str_replace(" </{$tag}>", "</{$tag}>", $xml);
	}

	$intermediaryDocument->loadXML($xml);

	if ($debugDumpIntermediary) {
		$intermediaryDocument->preserveWhiteSpace = false;
		$intermediaryDocument->formatOutput = true;
		echo $intermediaryDocument->saveXML();
		exit();
	}

	//==========================================================================
	// Step 5: Convert from the intermediary XML format to Markdown
	//==========================================================================

	$xslDocument = new DOMDocument("1.0", "UTF-8");
	if ($includeImages) {
		// Replace image placeholder with image template
		$imageTemplate = sprintf(IMAGE_TEMPLATE, $imageFolder, basename($mdFilename, '.md'));
		$xslDocument->loadXML(sprintf(INTERMEDIARY_TO_MARKDOWN_TRANSFORM, $imageTemplate));
	} else {
		// Replace image placeholder with a blank string
		$xslDocument->loadXML(sprintf(INTERMEDIARY_TO_MARKDOWN_TRANSFORM, ''));
	}

	$processor = new XSLTProcessor();
	$processor->importStyleSheet($xslDocument);
	$markdown = $processor->transformToXml($intermediaryDocument);

	// Trim lines so we have no space in the beginning
	$lines = explode("\n", $markdown);
	for ($i=0; $i<count($lines); $i++) {
		$lines[$i] = trim($lines[$i]);
	}
	$markdown = implode("\n", $lines);

	//==========================================================================
	// Step 6: If the Markdown output file was specified, write it. Otherwise
	// just write to STDOUT (echo)
	//==========================================================================

	if ($mdFilename !== null) {
		file_put_contents($mdFilename, $markdown);
	} else {
		echo $markdown;
	}

	//==========================================================================
	// Step 7: Clean-up
	//==========================================================================

	if (file_exists($documentFolder)) {
		rrmdir($documentFolder);
	}
}

//==============================================================================
// Support functions
//==============================================================================

function extractFolder($zip, $folderName, $destination, $imageFolder = null, $mdFilename = null) {
	for ($i = 0; $i < $zip->numFiles; $i++) {
		$fileName = $zip->getNameIndex($i);

		if (strpos($fileName, $folderName) !== false) {
			if (!is_null($imageFolder) && !is_null($mdFilename)) {
				// Save matching images to disk
				if (preg_match('([^\s]+(\.(?i)(bmp|gif|jpe?g|png))$)', $fileName)) {
					file_put_contents("$imageFolder/" . basename($mdFilename, '.md') . '.' . basename($fileName), $zip->getFromIndex($i));
				}
			}

			$zip->extractTo($destination, $fileName);
		}
	}
}

function rrmdir($dir) {
	foreach(glob($dir . "/*") as $file) {
		if(is_dir($file)) {
			rrmdir($file);
		} else {
			unlink($file);
		}
	}
	rmdir($dir);
}

//==============================================================================
// XSL Stylesheets
//==============================================================================

define("DOCX_TO_INTERMEDIARY_TRANSFORM", <<<'XML'
<?xml version="1.0"?>
<xsl:stylesheet version="1.0"
	xmlns:i="urn:docx2md:intermediary"
	xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
	xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
	xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
	xmlns:rels="http://schemas.openxmlformats.org/package/2006/relationships"
	xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
	xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
	xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties">

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
	<xsl:template match="w:p[ w:pPr/w:pStyle/@w:val[ starts-with( ., 'Heading' ) ] ]">
		<xsl:variable name="style" select="w:pPr/w:pStyle/@w:val[ starts-with( ., 'Heading' ) ]" />
		<xsl:variable name="level" select="substring( $style, 8, 1 )" />
		<xsl:variable name="type" select="translate( substring( $style, 9 ), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz' )" />
		<xsl:if test="count(w:r)">
			<i:heading>
				<xsl:attribute name="level"><xsl:value-of select="$level" /></xsl:attribute>
				<xsl:if test="$type != ''"><xsl:attribute name="type"><xsl:value-of select="$type" /></xsl:attribute></xsl:if>
				<xsl:apply-templates />
			</i:heading>
		</xsl:if>
	</xsl:template>

	<!-- Regular paragraph style -->
	<xsl:template match="w:p">
		<xsl:if test="count(w:r)">
			<i:para><xsl:apply-templates /></i:para>
		</xsl:if>
	</xsl:template>

	<!-- List items -->
	<xsl:template match="w:p[ w:pPr/w:numPr ]">
		<xsl:if test="count(w:r)">
			<i:listitem level="{ w:pPr/w:numPr/w:ilvl/@w:val }" type="{ w:pPr/w:numPr/w:numId/@w:val }"><xsl:apply-templates /></i:listitem>
		</xsl:if>
	</xsl:template>
	<xsl:template match="w:p[ w:pPr/w:pStyle/@w:val = 'ListBullet']">
		<xsl:if test="count(w:r)">
			<i:listitem level="0" type="1"><xsl:apply-templates /></i:listitem>
		</xsl:if>
	</xsl:template>
	<xsl:template match="w:p[ w:pPr/w:pStyle/@w:val = 'ListNumber']">
		<xsl:if test="count(w:r)">
			<i:listitem level="0" type="2"><xsl:apply-templates /></i:listitem>
		</xsl:if>
	</xsl:template>

	<!-- Text content -->
	<xsl:template match="w:r">
		<xsl:apply-templates />
	</xsl:template>
	<xsl:template match="w:r[w:rPr/w:b and not(w:rPr/w:i)]/w:t">
		<!-- bold -->
		<i:bold><xsl:value-of select="." /></i:bold>
	</xsl:template>
	<xsl:template match="w:r[w:rPr/w:i and not(w:rPr/w:b)]/w:t">
		<!-- italic -->
		<i:italic><xsl:value-of select="." /></i:italic>
	</xsl:template>
	<xsl:template match="w:r[w:rPr/w:i and w:rPr/w:b]/w:t">
		<!-- bold + italic -->
		<i:italic><i:bold><xsl:value-of select="." /></i:bold></i:italic>
	</xsl:template>
	<xsl:template match="w:t">
		<!-- normal -->
		<xsl:value-of select="." />
	</xsl:template>
	<xsl:template match="w:br">
		<i:linebreak />
	</xsl:template>

	<!-- Complete hyperlinks -->
	<xsl:template match="w:hyperlink">
		<xsl:variable name="id" select="@r:id" />
		<xsl:if test="count(w:r)">
			<i:link>
				<xsl:attribute name="href"><xsl:value-of select="/w:document/rels:Relationships/rels:Relationship[@Id=$id]/@Target" /></xsl:attribute>
				<xsl:if test="/w:document/rels:Relationships/rels:Relationship[@Id=$id]/@TargetMode">
					<xsl:attribute name="target"><xsl:value-of select="/w:document/rels:Relationships/rels:Relationship[@Id=$id]/@TargetMode" /></xsl:attribute>
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
			<xsl:attribute name="src"><xsl:value-of select="/w:document/data/@word-folder" /><xsl:value-of select="/w:document/rels:Relationships/rels:Relationship[@Id=$id]/@Target" /></xsl:attribute>
			<xsl:attribute name="width"><xsl:value-of select="round( ancestor::w:drawing[1]//wp:extent/@cx div 9525 )" /></xsl:attribute>
			<xsl:attribute name="height"><xsl:value-of select="round( ancestor::w:drawing[1]//wp:extent/@cy div 9525 )" /></xsl:attribute>
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

define("INTERMEDIARY_TO_MARKDOWN_TRANSFORM", <<<'XML'
<?xml version="1.0"?>
<xsl:stylesheet version="1.0"
	xmlns:i="urn:docx2md:intermediary"
	xmlns:str="http://exslt.org/strings" extension-element-prefixes="str"
	xmlns:xsl="http://www.w3.org/1999/XSL/Transform">

	<xsl:output
		method="text"
		omit-xml-declaration="yes"
		media-type="text/plain"
	/>

	<xsl:template match="@*|node()"><xsl:copy><xsl:apply-templates select="@*|node()"/></xsl:copy></xsl:template>

	<xsl:template match="i:document"><xsl:apply-templates /><xsl:text>&#xa;</xsl:text><xsl:for-each select="//i:link"><xsl:text>&#32;&#32;[</xsl:text><xsl:value-of select="position()" /><xsl:text>]:&#32;</xsl:text><xsl:value-of select="@href" /><xsl:text>&#xa;</xsl:text></xsl:for-each></xsl:template>

	<xsl:template match="i:body"><xsl:apply-templates /></xsl:template>

	<xsl:template match="i:heading"><xsl:value-of select="substring('######', 1, @level)" /><xsl:text>&#32;</xsl:text><xsl:apply-templates /><xsl:text>&#xa;</xsl:text><xsl:text>&#xa;</xsl:text></xsl:template>

	<xsl:template match="i:link"><xsl:text>[</xsl:text><xsl:value-of select="." /><xsl:text>][</xsl:text><xsl:value-of select="count(preceding::i:link) + 1" /><xsl:text>]</xsl:text></xsl:template>

	<xsl:template match="i:italic"><xsl:text>*</xsl:text><xsl:apply-templates /><xsl:text>*</xsl:text></xsl:template>

	<xsl:template match="i:bold"><xsl:text>__</xsl:text><xsl:apply-templates /><xsl:text>__</xsl:text></xsl:template>

	<xsl:template match="i:para"><xsl:if test="./* or text() != ''"><xsl:apply-templates /><xsl:text>&#xa;</xsl:text><xsl:text>&#xa;</xsl:text></xsl:if></xsl:template>

	<xsl:template match="i:linebreak"><xsl:text>&#xa;</xsl:text></xsl:template>

	<!-- Bullet list-item -->
	<xsl:template match="i:listitem[@type='1']"><xsl:value-of select="substring('&#9;&#9;&#9;&#9;&#9;&#9;&#9;&#9;&#9;&#9;', 1, @level)" /><xsl:text>-&#9;</xsl:text><xsl:apply-templates /><xsl:text>&#xa;</xsl:text><xsl:if test="local-name(following-sibling::i:*[1]) != 'listitem'"><xsl:text>&#xa;</xsl:text></xsl:if></xsl:template>

	<!-- Numbered list-item -->
	<xsl:template match="i:listitem[@type='2']"><xsl:variable name="level" select="@level" /><xsl:variable name="type" select="@type" /><xsl:value-of select="substring('&#9;&#9;&#9;&#9;&#9;&#9;&#9;&#9;&#9;&#9;', 1, $level)" /><xsl:value-of select="count(preceding::i:listitem[@level=$level and @type=$type]) + 1" /><xsl:text>.&#9;</xsl:text><xsl:apply-templates /><xsl:text>&#xa;</xsl:text><xsl:if test="local-name(following-sibling::i:*[1]) != 'listitem'"><xsl:text>&#xa;</xsl:text></xsl:if></xsl:template>

	<!-- Image Template Placeholder -->
	%s

	<!-- Trim whitespace on headings, paragraphs and list-items -->
	<!--xsl:template match="i:heading/text() | i:para/text() | i:listitem/text()"><xsl:choose><xsl:when test="substring(., string-length(.), 1) = ' '"><xsl:value-of select="substring(., 1, string-length(.) - 1)" /></xsl:when><xsl:otherwise><xsl:value-of select="." /></xsl:otherwise></xsl:choose></xsl:template-->

	<!-- Escape asterix -->
	<xsl:template match="text()"><xsl:call-template name="string-replace-all">
		<xsl:with-param name="text" select="." />
		<xsl:with-param name="replace" select="'*'" />
		<xsl:with-param name="by" select="'\*'" />
	</xsl:call-template></xsl:template>

	<!-- Superscript ® -->
	<xsl:template match="text()"><xsl:call-template name="string-replace-all">
		<xsl:with-param name="text" select="." />
		<xsl:with-param name="replace" select="'®'" />
		<xsl:with-param name="by" select="'&lt;sup&gt;®&lt;/sup&gt;'" />
	</xsl:call-template></xsl:template>

	<!-- Utility string replace -->
	<xsl:template name="string-replace-all">
		<xsl:param name="text" />
		<xsl:param name="replace" />
		<xsl:param name="by" />
		<xsl:choose>
			<xsl:when test="contains($text, $replace)"><xsl:value-of select="substring-before($text, $replace)" /><xsl:value-of select="$by" /><xsl:call-template name="string-replace-all">
				<xsl:with-param name="text" select="substring-after($text, $replace)" />
				<xsl:with-param name="replace" select="$replace" />
				<xsl:with-param name="by" select="$by" />
			</xsl:call-template></xsl:when>
			<xsl:otherwise><xsl:value-of select="$text" /></xsl:otherwise>
		</xsl:choose>
	</xsl:template>

</xsl:stylesheet>
XML
);

define('IMAGE_TEMPLATE', '<!-- Image --><xsl:template match="i:image"><xsl:text>![Image](%s/%s.</xsl:text><xsl:value-of select="str:tokenize(@src, \'/\')[last()]" /><xsl:text>)</xsl:text></xsl:template>');

docx2md($argv);