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
        [string]$Text
    )
    try{
        Register-ITextSharpLib
        $pdf = [iTextSharp.text.Document]::new()
        $Null = New-PdfDocument -Document $pdf -File "$Path" -TopMargin 20 -BottomMargin 20 -LeftMargin 20 -RightMargin 20 -Author "Guillaume Plante"
        $pdf.Open()
      
        $ret = Add-TitleToPdf -Document $pdf -Text "Crypto Data" -Color "Red" -Centered
        $ret = Add-TextToPdf -Document $pdf -Text "$Text"
    
        $pdf.Close()
    }
    catch{
        Write-Error $_
    }
}


function Write-PdfFile {

    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory=$true,Position=0)]
        [string]$Path,
        [Parameter(Mandatory=$true,Position=1)]
        [string]$Text
    )
    try{
        Register-ITextSharpLib
        $pdf = [iTextSharp.text.Document]::new()
        New-PdfDocument -Document $pdf -File "$Path" -TopMargin 20 -BottomMargin 20 -LeftMargin 20 -RightMargin 20 -Author "Guillaume Plante"
        $pdf.Open()
      
        Add-TitleToPdf -Document $pdf -Text "This Is the Title Test" -Color "Blue" -Centered
        Add-TextToPdf -Document $pdf -Text "$Text"
    
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
    
    $Document.SetPageSize([iTextSharp.text.PageSize]::A4)
    $Document.SetMargins($LeftMargin, $RightMargin, $TopMargin, $BottomMargin)
    [void][iTextSharp.text.pdf.PdfWriter]::GetInstance($Document, [System.IO.File]::Create($File))
    $Document.AddAuthor($Author)
    $Document;
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
        [string]$Color = "BLACK"
    )

    $p = New-Object iTextSharp.text.Paragraph
    $p.Font = [iTextSharp.text.FontFactory]::GetFont($FontName, $FontSize, [iTextSharp.text.Font]::NORMAL, [iTextSharp.text.BaseColor]::$Color)
    $p.SpacingBefore = 2
    $p.SpacingAfter = 2
    $p.Add($Text)
    $Document.Add($p)
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
        [Parameter(Mandatory=$false)]
        [string]$Color = "BLACK", 
        [Parameter(Mandatory=$false)]
        [switch]$Centered
    )    
    $p = New-Object iTextSharp.text.Paragraph
    $p.Font = [iTextSharp.text.FontFactory]::GetFont($FontName, $FontSize, [iTextSharp.text.Font]::BOLD, [iTextSharp.text.BaseColor]::$Color)
    if ($Centered) { $p.Alignment = [iTextSharp.text.Element]::ALIGN_CENTER }
    $p.SpacingBefore = 5
    $p.SpacingAfter = 5
    $p.Add($Text)
    $Document.Add($p)
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

    [iTextSharp.text.Image]$img = [iTextSharp.text.Image]::GetInstance($File)
    if ($Centered) { $p.Alignment = [iTextSharp.text.Element]::ALIGN_CENTER }
    $img.ScalePercent(100)
    $Document.Add($img)
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

    $t = New-Object iTextSharp.text.pdf.PDFPTable($Cols)
    $t.SpacingBefore = 5
    $t.SpacingAfter = 5
    if (!$Centered) { $t.HorizontalAlignment = 0 }
    foreach ($data in $Dataset)
    {
        $t.AddCell($data);
    }
    $Document.Add($t)
}

function Test-CreatePdf{
        
    Register-ITextSharpLib
    $pdf = [iTextSharp.text.Document]::new()
    New-PdfDocument -Document $pdf -File "C:\Tmp\Test.pdf" -TopMargin 20 -BottomMargin 20 -LeftMargin 20 -RightMargin 20 -Author "Guillaume Plante"
    $pdf.Open()
    # Add-Image -Document $pdf -File "$Flogo" -Centered
    $Ret = Add-TitleToPdf -Document $pdf -Text "THIS IS THE STORY OF THE LONELY VAGABOND" -Color "Magenta" -Centered
    $Ret = Add-TextToPdf -Document $pdf -Text "This would serve as a short paragraph,This would serve as a short paragraph,This would serve as a short paragraph,This would serve as a short paragraph"
    #Add-Table -Document $pdf -Dataset @('Name', "$first", "Login", "$userprinc", "Email", "$SamAccountName", "Password", "String") -Cols 2 -Centered
    Add-TextToPdf -Document $pdf -Text "We could use this space to show the Help Desk Ticket System"
    $pdf.Close()

}