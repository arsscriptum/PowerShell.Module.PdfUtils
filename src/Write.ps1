#╔════════════════════════════════════════════════════════════════════════════════╗
#║                                                                                ║
#║   Write.ps1                                                                    ║
#║   Write utils                                                                  ║
#║                                                                                ║
#╟────────────────────────────────────────────────────────────────────────────────╢
#║   Guillaume Plante <guillaume.plante@luminator.com>                            ║
#║   Copyright (C) Luminator Technology Group.  All rights reserved.              ║
#╚════════════════════════════════════════════════════════════════════════════════╝

function Write-PdfCryptoFile {

    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory=$true,Position=0)]
        [string]$Path,
        [Parameter(Mandatory=$true,Position=1)]
        [string]$Title,
        [Parameter(Mandatory=$true,Position=2)]
        [string]$Text
    )
    try{
        Register-ITextSharpLib
        
        $pdf = [iTextSharp.text.Document]::new()
        $Null = New-PdfDocument -Document $pdf -File "$Path" -TopMargin 20 -BottomMargin 20 -LeftMargin 20 -RightMargin 20 -Author "$ENV:USERNAME"
        $pdf.Open()
      
        $ret = Add-TitleToPdf -Document $pdf -Text "$Title" -Color "Red" -Centered -FontName "Consolas"
        $ret = Add-TextToPdf -Document $pdf -Text "$Text" -FontName "Consolas"
    
        $pdf.Close()
    }
    catch{
        Write-Error $_
    }
}


Function New-PdfDocument{
    [CmdletBinding(SupportsShouldProcess)]
    param( 
        [Parameter(Mandatory=$true,Position=0)]
        [iTextSharp.text.Document]$Document, 
        [Parameter(Mandatory=$true,Position=1)]
        [ValidateNotNullOrEmpty()]
        [string]$File, 
        [Parameter(Mandatory=$false)]
        [int32]$TopMargin=20, 
        [Parameter(Mandatory=$false)]
        [int32]$BottomMargin=20, 
        [Parameter(Mandatory=$false)]
        [int32]$LeftMargin=20, 
        [Parameter(Mandatory=$false)]
        [int32]$RightMargin=20, 
        [Parameter(Mandatory=$false)]
        [string]$Author = "$ENV:USERNAME"
    )
    try{
        $Document.SetPageSize([iTextSharp.text.PageSize]::A4)
        $Document.SetMargins($LeftMargin, $RightMargin, $TopMargin, $BottomMargin)
        [void][iTextSharp.text.pdf.PdfWriter]::GetInstance($Document, [System.IO.File]::Create($File))
        $Document.AddAuthor($Author)
        $Document;
    }catch{
        Show-ExceptionDetails $_ -ShowStack
    }
}



        
# Add a text paragraph to the document, optionally with a font name, size and color
function Add-TextToPdf {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory=$true,Position=0)]
        [iTextSharp.text.Document]$Document, 
        [Parameter(Mandatory=$true,Position=1)]
        [ValidateNotNullOrEmpty()]
        [string]$Text, 
        [Parameter(Mandatory=$false)]
        [string]$FontName = "Arial", 
        [Parameter(Mandatory=$false)]
        [int32]$FontSize = 12, 
        [Parameter(Mandatory=$false)]
        [ValidateSet("WHITE", "LIGHT_GRAY", "GRAY", "DARK_GRAY", "BLACK", "RED", "PINK", "ORANGE", "YELLOW", "GREEN", "MAGENTA", "CYAN", "BLUE")]
        [string]$Color = "BLACK",
        [Parameter(Mandatory=$false)]
        [ValidateSet('BOLDITALIC', 'DEFAULTSIZE', 'ITALIC', 'NORMAL', 'STRIKETHRU', 'UNDEFINED', 'UNDERLINE')]
        [string]$Style = "NORMAL"
    )
    try{
       
        $p = New-Object iTextSharp.text.Paragraph
        $p.Font = [iTextSharp.text.FontFactory]::GetFont($FontName, $FontSize,[iTextSharp.text.Font]::"$Style", [iTextSharp.text.BaseColor]::"$Color")
        $p.SpacingBefore = 2
        $p.SpacingAfter = 2
        $p.Add($Text)
        $Document.Add($p)
    }catch{
        Show-ExceptionDetails $_ -ShowStack
    }
}
        
# Add a title to the document, optionally with a font name, size, color and centered
function Add-TitleToPdf {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory=$true,Position=0)]
        [iTextSharp.text.Document]$Document, 
        [Parameter(Mandatory=$true,Position=1)]
        [ValidateNotNullOrEmpty()]
        [string]$Text, 
        [Parameter(Mandatory=$false)]
        [string]$FontName = "Arial", 
        [Parameter(Mandatory=$false)]
        [int32]$FontSize = 12, 
        [ValidateSet("WHITE", "LIGHT_GRAY", "GRAY", "DARK_GRAY", "BLACK", "RED", "PINK", "ORANGE", "YELLOW", "GREEN", "MAGENTA", "CYAN", "BLUE")]
        [string]$Color = "BLACK",
        [Parameter(Mandatory=$false)]
        [ValidateSet('BOLDITALIC', 'DEFAULTSIZE', 'ITALIC', 'NORMAL', 'STRIKETHRU', 'UNDEFINED', 'UNDERLINE')]
        [string]$Style = "NORMAL",
        [Parameter(Mandatory=$false)]
        [switch]$Centered
    )    

    try{
        $p = New-Object iTextSharp.text.Paragraph
        $p.Font = [iTextSharp.text.FontFactory]::GetFont($FontName, $FontSize,[iTextSharp.text.Font]::"$Style", [iTextSharp.text.BaseColor]::"$Color")
        if ($Centered) { $p.Alignment = [iTextSharp.text.Element]::ALIGN_CENTER }
        $p.SpacingBefore = 5
        $p.SpacingAfter = 5
        $p.Add($Text)
        $Document.Add($p)
    }catch{
        Show-ExceptionDetails $_ -ShowStack
    }        
}
        
# Add an image to the document, optionally scaled
function Add-ImageToPdf {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory=$true,Position=0)]
        [iTextSharp.text.Document]$Document, 
        [Parameter(Mandatory=$true,Position=1)]
        [ValidateNotNullOrEmpty()]
        [string]$File, 
        [Parameter(Mandatory=$false)]
        [int32]$Scale = 100
    )

    try{
        [iTextSharp.text.Image]$img = [iTextSharp.text.Image]::GetInstance($File)
        if ($Centered) { $p.Alignment = [iTextSharp.text.Element]::ALIGN_CENTER }
        $img.ScalePercent(100)
        $Document.Add($img)
    }catch{
        Show-ExceptionDetails $_ -ShowStack
    }
}
        
# Add a table to the document with an array as the data, a number of columns, and optionally centered
function Add-TableToPdf {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory=$true,Position=0)]
        [iTextSharp.text.Document]$Document, 
        [Parameter(Mandatory=$true,Position=1)]
        [string[]]$Dataset, 
        [Parameter(Mandatory=$false)]
        [int32]$Cols = 3,
        [Parameter(Mandatory=$false)]
        [switch]$Centered
    ) 
    try{
        $t = New-Object iTextSharp.text.pdf.PDFPTable($Cols)
        $t.SpacingBefore = 5
        $t.SpacingAfter = 5
        if (!$Centered) { $t.HorizontalAlignment = 0 }
        foreach ($data in $Dataset)
        {
            $t.AddCell($data);
        }
        $Document.Add($t)
    }catch{
        Show-ExceptionDetails $_ -ShowStack
    }
}
