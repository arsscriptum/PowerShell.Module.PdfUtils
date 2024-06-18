#╔════════════════════════════════════════════════════════════════════════════════╗
#║                                                                                ║
#║   Helpers.ps1                                                                  ║
#║   common functions                                                             ║
#║                                                                                ║
#╟────────────────────────────────────────────────────────────────────────────────╢
#║   Guillaume Plante <guillaume.plante@luminator.com>                            ║
#║   Copyright (C) Luminator Technology Group.  All rights reserved.              ║
#╚════════════════════════════════════════════════════════════════════════════════╝


function Get-PdfUtilsModuleInformation {

        $ModuleName = $ExecutionContext.SessionState.Module
        $ModuleScriptPath = $ScriptMyInvocation = $Script:MyInvocation.MyCommand.Path
        $ModuleScriptPath = (Get-Item "$ModuleScriptPath").DirectoryName
        $CurrentScriptName = $Script:MyInvocation.MyCommand.Name
        $ModuleInformation = @{
            Module        = $ModuleName
            ModuleScriptPath  = $ModuleScriptPath
            CurrentScriptName = $CurrentScriptName
        }
        return $ModuleInformation
}


function Get-ExportedLibsPath{   
    $ModPath = (Get-PdfUtilsModuleInformation).ModuleScriptPath
    $ExportsPath = Join-Path $ModPath 'exports'
    return $ExportsPath
}

function Register-ITextSharpLib{
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory = $false)]
        [Switch]$Force
    )

    $ExportsPath = Get-ExportedLibsPath
    $AllLibs = Get-ChildItem -Path "$ExportsPath" -File -Filter "*.dll"
    ForEach($dll in $AllLibs){
        $dllfn = $dll.Fullname
        $dllname = $dll.Name
        Write-Verbose "Importing $dllname"
        Add-Type -Path "$dllfn"
    }
}

function Test-ITextSharpLibRegistered{   

    if (!("iTextSharp.text.pdf.parser.PdfTextExtractor" -as [type])) {
        return $False
    }
    return $True
}