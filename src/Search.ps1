#╔════════════════════════════════════════════════════════════════════════════════╗
#║                                                                                ║
#║   Search.ps1                                                                   ║
#║   search utils                                                                 ║
#║                                                                                ║
#╟────────────────────────────────────────────────────────────────────────────────╢
#║   Guillaume Plante <guillaume.plante@luminator.com>                            ║
#║   Copyright (C) Luminator Technology Group.  All rights reserved.              ║
#╚════════════════════════════════════════════════════════════════════════════════╝




function Search-PdfFolder {
    <#
    .SYNOPSIS
            Search PDF files for a pattern
    .DESCRIPTION
            Search PDF files for a pattern
    #>

    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory=$true,Position=0)]
        [string]$Path,
        [Parameter(Mandatory=$true,Position=1)]
        [string]$Pattern
    )
    try{
        Register-ITextSharpLib
        
        $pdfs = (gci $Path -Filter '*.pdf' -File).Fullname
        

        foreach($pdfname in $pdfs) {

            # prepare the pdf
            $reader = New-Object iTextSharp.text.pdf.pdfreader -ArgumentList "$pdfname"
            $Basename = (Get-Item $pdfname).Basename

            $pageNum = $reader.NumberOfPages
            Write-Host "[$Basename] " -f White -NoNewLine 
            Write-Host "$pageNum pages..." -f DarkGray

            $TotalMatches = 0
         
            for($page = 1; $page -le $reader.NumberOfPages; $page++) {

                # set the page text
                $pageText = [iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($reader,$page).Split([char]0x000A)
                $show = $false
                
                ForEach($val in $pageText){
                    if($val -match $Pattern) {
                        $TotalMatches++

                        Write-Host "$val" -f DarkYellow
                        $Array = $val.Split(' ')
                        
                    }
                }             
            }

            Write-Host "Total: " -f DarkRed -NoNewLine
            Write-Host "$TotalMatches" -f DarkYellow
        
        }


    }
    catch{
        Write-Error $_
    }
}
