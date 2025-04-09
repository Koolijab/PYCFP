# Created by: Koolijab
# Created at: 10.04.2025

#To use this you must have Imagemagick 7+ installed.


# ==== Request input ====
$widthInput = Read-Host "Enter card width in mm (standard: 63)"
$TargetWidthMM = if ($widthInput) { [double]$widthInput } else { 63 }

$heightInput = Read-Host "Enter card height in mm (standard: 88)"
$TargetHeightMM = if ($heightInput) { [double]$heightInput } else { 88 }

$dpiInput = Read-Host "Enter the desired DPI (standard: 450)"
$DPI = if ($dpiInput) { [int]$dpiInput } else { 450 }

$borderInput = Read-Host "Border width for printing tolerance in mm (standard: 3, no boarder: 0)"
$BorderMM = if ($borderInput) { [double]$borderInput } else { 3 }

$targetPathInput = Read-Host "Saving Directory (e.g. C:\Users\User\Cards) - folder structure will be preserved"
if (-not (Test-Path $targetPathInput)) {
    Write-Host "Directory does not exist!" -ForegroundColor Red
    exit
}
$TargetBase = $targetPathInput



# ==== Calculation ====
$pxPerMM = $DPI / 25.4
$TargetWidthPX = [math]::Round($TargetWidthMM * $pxPerMM)
$TargetHeightPX = [math]::Round($TargetHeightMM * $pxPerMM)
$BorderPX = [math]::Round($BorderMM * $pxPerMM)

Write-Host "Pixels/mm:     $pxPerMM"
Write-Host "Pixels width:  $TargetWidthPX"
Write-Host "Pixels height: $TargetHeightPX"
Write-Host "Border width:  $BorderPX"



# ==== Image processing ====
$SourceRoot = Get-Location
$images = Get-ChildItem -Recurse -Include *.jpg, *.jpeg, *.png -File

foreach ($image in $images) {
    Write-Host "`nProcessing: $($image.FullName)"

    $sourcePath = $image.Directory.FullName
    $relativePath = $sourcePath.Replace($SourceRoot.Path, "").TrimStart('\')
    $outputDir = Join-Path $TargetBase $relativePath
    $outputFile = Join-Path $outputDir $image.Name

    if (-not (Test-Path $outputDir)) {
        New-Item -ItemType Directory -Path $outputDir | Out-Null
    }

    $resizeArg = "${TargetWidthPX}x${TargetHeightPX}!"
    $baseArgs = @(
        "`"$($image.FullName)`"",
        "-resize", $resizeArg,
        "-units", "PixelsPerInch",
        "-density", $DPI
    )

    if ($BorderMM -gt 0) {
		$ExtendedWidth = $TargetWidthPX + 2 * $BorderPX
		$ExtendedHeight = $TargetHeightPX + 2 * $BorderPX
		
        $baseArgs += (
			"-set", "option:distort:viewport", "${ExtendedWidth}x${ExtendedHeight}",
			"-virtual-pixel", "edge",
			"-distort", "srt", "0,0 1 0 ${BorderPx},${BorderPx}"
		)	
    }

    $baseArgs += "`"$outputFile`""

    magick @baseArgs

    Write-Host "Saved: $outputFile"
}
