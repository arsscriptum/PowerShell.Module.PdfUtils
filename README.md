# Pdf Utils PowerShell Module

A module providing somem pdf related functionalities in PowerShell

### Author 

Guillaume Plante

Protect-FileToPdfCipher $File1

    $File1 = "D:\Scripts\AutoHotkey\RemapKeys.ahk"
    $File2 = "D:\Scripts\AutoHotkey\MyShortcuts.ahk"


function Protect-FileToPdfCipher {
    [CmdletBinding(SupportsShouldProcess)]
    Param(
        [Parameter(Mandatory=$True, Position=0)]
        [ValidateScript({Test-Path $_})]
        [string]$Path
    )

    $OutFile =  "$ENV:Temp\Temp.txt"

    $Data = Protect-FileToCipherBlock -Path $Path | Out-String
    $FileName = (Get-Item $Path).Name
    $Basename = Get-Item $Path | Select -ExpandProperty Basename
    $DirectoryName = Get-Item $Path | Select -ExpandProperty DirectoryName
    $PdfFile = "{0}\{1}.pdf" -f $DirectoryName, $Basename

    Remove-Item -Path $PdfFile -Force -ErrorAction Ignore | Out-Null

    Write-PdfCryptoFile -Path $PdfFile -Title $FileName -Text $Data
    Write-Host "Wrote file: $PdfFile . To Decrypt, use Unprotect-FileToPdfCipher"
}


function Unprotect-FileToPdfCipher {
    [CmdletBinding(SupportsShouldProcess)]
    Param(
        [Parameter(Mandatory=$True, Position=0)]
        [ValidateScript({Test-Path $_})]
        [string]$Path
    )
    [iTextSharp.text.pdf.PdfReader]$reader  = [iTextSharp.text.pdf.PdfReader]::new($Path)
    $text = ""
    for ($i = 1; $i -le $reader.NumberOfPages; $i++) {
        $text = $text + [iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($reader, $i)
    }

    $Content = Unprotect-CipherBlockToFile -Content $text
    $Content
}