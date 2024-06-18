#╔════════════════════════════════════════════════════════════════════════════════╗
#║                                                                                ║
#║   Write.ps1                                                                    ║
#║   Write utils                                                                  ║
#║                                                                                ║
#╟────────────────────────────────────────────────────────────────────────────────╢
#║   Guillaume Plante <guillaume.plante@luminator.com>                            ║
#║   Copyright (C) Luminator Technology Group.  All rights reserved.              ║
#╚════════════════════════════════════════════════════════════════════════════════╝


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

    }
    catch{
        Write-Error $_
    }
}
