#╔════════════════════════════════════════════════════════════════════════════════╗
#║                                                                                ║
#║   Config.ps1                                                                   ║
#║   config module                                                                ║
#║                                                                                ║
#╟────────────────────────────────────────────────────────────────────────────────╢
#║   Guillaume Plante <guillaume.plante@luminator.com>                            ║
#║   Copyright (C) Luminator Technology Group.  All rights reserved.              ║
#╚════════════════════════════════════════════════════════════════════════════════╝

#===============================================================================
# Dependencies
#===============================================================================

$ModuleDependencies = @( 'PowerShell.Module.Core','PowerShell.Module.Github' )
$FunctionDependencies = @( 'Show-ExceptionDetails','Get-ScriptDirectory', 'New-Directory','ConvertFrom-Ctime','ConvertTo-Ctime' )
$EnvironmentVariable = @( 'OrganizationHKCU', 'OrganizationHKLM' )

try{
    $ScriptMyInvocation = $Script:MyInvocation.MyCommand.Path
    $CurrentScriptName = $Script:MyInvocation.MyCommand.Name
    $PSScriptRootValue = 'null' ; if($PSScriptRoot) { $PSScriptRootValue = $PSScriptRoot}
    $ModuleName = (Get-Item $PSScriptRootValue).Name
    Write-Host "===============================================================================" -f DarkRed
    Write-Host "MODULE $ModuleName BUILD CONFIGURATION AND VALIDATION" -f DarkYellow;
    Write-Host "===============================================================================" -f DarkRed    

    Write-Host "[CONFIG] " -f Blue -NoNewLine
    Write-Host "CHECKING ENVIRONMENT VARIABLE.."
    $EnvironmentVariable.ForEach({
        $EnvVar=$_
        $Var = [System.Environment]::GetEnvironmentVariable($EnvVar,[System.EnvironmentVariableTarget]::User)
        if($Var -eq $null){
            throw "ERROR: MISSING $EnvVar Environment Variable"
        }else{
            Write-Host "`t`t[OK]`t" -f DarkGreen -NoNewLine
            Write-Host "$EnvVar"
        }
    })
    Write-Host "[CONFIG] " -f Blue -NoNewLine
    Write-Host "CHECKING MODULE DEPENDENCIES..."
    $ModuleDependencies.ForEach({
        $ModuleName=$_
        $ModPtr = Get-Module "$ModuleName" -ErrorAction Ignore
        if($ModPtr -eq $null){
            throw "ERROR: MISSING $ModuleName module. Please import the dependency $ModuleName"
        }else{
            Write-Host "`t`t[OK]`t" -f DarkGreen -NoNewLine
            Write-Host "$ModuleName"
        }
    })
    Write-Host "[CONFIG] " -f Blue -NoNewLine
    Write-Host "CHECKING FUNCTION DEPENDENCIES..."
    $FunctionDependencies.ForEach({
        $Function=$_
        $FunctionPtr = Get-Command "$Function" -ErrorAction Ignore
        if($FunctionPtr -eq $null){
            throw "ERROR: MISSING $Function function. Please import the required dependencies"
        }else{
            Write-Host "`t`t[OK]`t" -f DarkGreen -NoNewLine
            Write-Host "$Function"
        }
    })
}catch{

    Write-Host "[ERROR] " -f DarkRed -NoNewLine
    Write-Host "$_"
    #throw "Error"

}
