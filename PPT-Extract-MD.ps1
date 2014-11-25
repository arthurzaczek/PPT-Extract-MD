<# 
.SYNOPSIS 
    Extracts power point slides to markdown
.DESCRIPTION 
    Supports Heading, Lists, Code and Images
.EXAMPLE
	.\PPT-Extract-MD.ps1 'MyPresentation.pptx'
.EXAMPLE
	ls *.pptx | % { .\PPT-Extract-MD.ps1 $_ }
.NOTES 
    Author     : Arthur Zaczek, arthur@dasz.at
	License    : GNU General Public License (GPL)
.LINK 
    http://dasz.at
#> 

param(
	[Parameter(HelpMessage="Path to the Powerpoint to extract")]
	[string]$file,
	
	[Parameter(HelpMessage="Extract each text frame without any bullet as source code")]
	[switch]$SourceCode,
	
	[Parameter(HelpMessage="Open the result when finished")]
	[switch]$Open
)

if(!$file) {
	get-help .\PPT-Extract-MD.ps1 -Full
	exit 1
}

# ---------------- Init variables ------------------------------
$file = resolve-path $file
$fileName = [System.IO.Path]::GetFileNameWithoutExtension($file)
$outFile =  "$fileName.md"

# ---------------- helper functions ------------------------------
function hasBullets($paragraphs) {
    foreach($p in $paragraphs) {
	   if($p.ParagraphFormat.Bullet.Visible) {
            return $true; 
       }
    }
    return $false;
}

function sanitizeName($name) {
	if(!$name) { return "" }
	return ($name -replace "\d","").Trim().Replace(" ", "-")
}

function sanitizeFileName($name) {
	if(!$name) { return "" }
	$name = $name.Replace(" ", "_")
	$name = $name.Replace("ö", "oe")
	$name = $name.Replace("Ö", "Oe")
	$name = $name.Replace("ä", "ae")
	$name = $name.Replace("Ä", "Ae")
	$name = $name.Replace("ü", "ue")
	$name = $name.Replace("Ü", "Ue")
	$name = $name.Replace("ß", "sz")
	$name = $name.Replace("'", "_")
	return $name.Trim()
}


function isSingleParagraph($paragraphs) {
    return $paragraphs.Count -le 1;
}

function out-result {
    $input | out-file $outFile -Append -Encoding "UTF8"
}

# ---------------- render functions ------------------------------
function renderHeader() {
	# init the file
	'' | out-file $outFile -Encoding "UTF8"
	# noting to do yet
}

function renderFooter() {
	# noting to do yet
}

function renderParagraphs($paragraphs) {
	foreach($p in $paragraphs) {
		if($p.Text -and $p.Text.Trim()) {
			if($p.ParagraphFormat.Bullet.Visible) {
				('* ' + $p.Text).TrimEnd() | out-result
			} else {
				$p.Text.TrimEnd() | out-result
			} 
		}
		else {
			'' | out-result
		}
	}
}

function renderSourceCode($shape) {
    "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ { .cs }" | out-result
	$shape.TextFrame2.TextRange.Text -split '[\r\n]' | ForEach {
	    # ensure correct line ending
		$_ | out-result
	}
    "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" | out-result
}

function renderTextShape($shape, $heading) {
	$style = sanitizeName $shape.Name
	
    if(($style -eq  "Titel") -or ($style -eq "Title")) {
		$text = $shape.TextFrame2.TextRange.Text
		"   $heading $text" | out-host
		'' | out-result
		"$heading $text" | out-result
	} else {
		$paragraphs = $shape.TextFrame2.TextRange.Paragraphs()
		if($SourceCode -and !(hasBullets $paragraphs) -and !(isSingleParagraph $paragraphs)) {
			renderSourceCode $shape
		} else {
			renderParagraphs $paragraphs
		}
	}
}

function renderPictureShape($shape) {
	if(!(test-path ".\images")) { mkdir ".\images" }
	$imgName = sanitizeFileName ("$fileName-" + $shape.Name + ".png");
	$imgAbsPath =  (Resolve-Path ".\images").Path + "\" + $imgName
	"  Saving image to $imgAbsPath" | out-host
	$shape.Export($imgAbsPath, 2)
	"![](images/$imgName)" | out-result
}

function renderSlide($slide) {
	$className = (sanitizeName $slide.CustomLayout.Name)
    "-> " + $slide.Name + " ($className)" | out-host
	
	switch($className) {
		"Titelfolie" { $heading = '#' }
		"Abschnittsüberschrift" { $heading = '#' }
		default { $heading = '##' }
	}
	
    foreach($shape in $slide.Shapes) {
		if($shape.Type -eq 13) { # msoPicture			
			renderPictureShape $shape
		} elseif($shape.HasTextFrame) {
			renderTextShape $shape $heading
        }
    }
}

# ---------------- Main ------------------------------
'Extracting "' + $file + '"' | out-host
'to         "' + $outFile + '"' | out-host
renderHeader

# init powerpoint
Add-type -AssemblyName office
$app = New-Object -ComObject powerpoint.application
$app.visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
$presentation = $app.Presentations.open($file)

foreach($slide in $presentation.Slides) {
	renderSlide $slide
}

renderFooter

# Quit powerpoint
$app.quit()
$app = $null
[gc]::collect()
[gc]::WaitForPendingFinalizers()

"finished...." | out-host

if($Open) {
	& .\$outFile
}
