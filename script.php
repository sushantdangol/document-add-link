<?php


$fileNameArr = [];

$file = $argv[1];
$directory = dirname($file);

$f = pathinfo($file, PATHINFO_FILENAME);

foreach (glob($directory . '/' . $f . '/*.*') as $filename) {
    $filename = basename($filename);
    $fileNameArr[] = $filename;
}

$zip = new ZipArchive();

if ($zip->open($file, ZipArchive::CREATE) !== TRUE) {
    echo "Cannot open $file :( "; die;
}

$xml = $zip->getFromName('word/document.xml');

$document = new DOMDocument();
$document->loadXML($xml);

$rels = $zip->getFromName('word/_rels/document.xml.rels');

$relationshipDocument = new DOMDocument();
$relationshipDocument->preserveWhiteSpace = false;
$relationshipDocument->formatOutput = true;
$relationshipDocument->loadXML($rels);

$relationships = $relationshipDocument->getElementsByTagName('Relationship');

$maxRId = 0;
foreach ($relationships as $relationship) {
    $rId = $relationship->getAttribute('Id');
    $rIdNumber = (int)str_replace('rId', '', $rId);
    if ($rIdNumber > $maxRId) {
        $maxRId = $rIdNumber;
    }
}

foreach ($fileNameArr as $fname) {
    $stringPosition = strpos($xml, $fname);
    
    if ($stringPosition !== false) {
        $nextRId = 'rId' . ($maxRId + 1);

        $newRelationship = $relationshipDocument->createElement('Relationship');
        $newRelationship->setAttribute('Id', $nextRId);
        $newRelationship->setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink');
        $newRelationship->setAttribute('Target', $directory . '/' . $f . '/' . $fname);
        $newRelationship->setAttribute('TargetMode', 'External');

        $relationshipDocument->documentElement->appendChild($newRelationship);

        $hyperlinkXml = '</w:t></w:r>' .
                        '<w:hyperlink r:id="' . $nextRId . '">' .
                        '<w:r>' .
                        '<w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr>' .
                        '<w:t>' . $fname . '</w:t>' .
                        '</w:r>' .
                        '</w:hyperlink><w:r><w:rPr></w:rPr><w:t xml:space="preserve">';
        
        $xml = str_replace($fname, $hyperlinkXml, $xml);

        $maxRId++;
    }
}

$zip->addFromString('word/_rels/document.xml.rels', $relationshipDocument->saveXML());
$zip->addFromString('word/document.xml', $xml);

$zip->close();

echo "done";

