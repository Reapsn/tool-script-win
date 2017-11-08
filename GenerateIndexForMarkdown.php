<?php
function generate_md_index($content, $depth, $prefix)
{
	$indexes = "";
	$lines = explode("\n", $content);
	
	foreach ($lines as $line) {
		$pos = strrpos(substr($line, 0, 8), "#");
 
		$tag = substr($line, 0, $pos + 1);
 
		$title = trim(substr($line, $pos + 1));
		$text = formatText($title, $prefix);
 
		switch ($tag) {
		case "#":
			$indexes .= "- " . $text;
			break;
		case "##":
			if ($depth > 1) {
				$indexes .= "    - " . $text;
			}
			break;
		case "###":
			if ($depth > 2) {
				$indexes .= "        - " . $text;
			}
			break;
		case "####":
			if ($depth > 3) {
				$indexes .= "                - " . $text;
			}
			break;
		case "#####":
			if ($depth > 4) {
				$indexes .= "                    - " . $text;
			}
			break;
		}
	}
 
	return $indexes;
}

function formatText($title, $prefix) {
	
	//$text = strtolower($text);
	$text = $title;
	$text = str_replace(Array("#", "¡ª", "+", "/", ".", "(", ")", "£¨", "£©", "£º", ":", "£¬", ",", "¡¾", "¡¿", ">", "<"), "", $text);
	$text = str_replace(" ", "-", $text);
	do {
		$text = str_replace("--", "-", $text);
	} while (strstr($text, "--"));	
	
	$prefix = ("" == $prefix) ? "" : $prefix;
	$text = urlencode($text);
	
	return "[" . $title . "](" . $prefix . "#" . $text . ")\n";
}
 
$prefix = ($argc > 3) ? $argv[3] : "";
$depth = ($argc > 2) ? (int)$argv[2] : 3;
$content = file_get_contents($argv[1]);
 
$mdindex = generate_md_index($content, $depth, $prefix);
echo $mdindex;
?>