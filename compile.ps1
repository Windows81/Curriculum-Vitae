[string]$docPath = Join-Path $PSScriptRoot 'cv.odt'
[string]$pdfPath = Join-Path $PSScriptRoot 'cv.pdf'

Push-Location (Join-Path $PSScriptRoot src)
7z -w u $docPath * > $null
$word = New-Object -ComObject Word.Application
$doc = $word.Documents.Open($docPath)
$doc.SaveAs($pdfPath, 17) # 17 is an alias for PDF.
$doc.Close()
$word.Quit()
Pop-Location