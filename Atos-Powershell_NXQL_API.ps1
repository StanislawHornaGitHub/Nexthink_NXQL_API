<#
.SYNOPSIS
    Script to retrieve data from Nexthink via NXQL API

.DESCRIPTION
    GUI PowerShell script to to retrieve data from Nexthink via NXQL API.
    Create to use on multi engine Nexthink Experience environments.
    The result file will contains merged output from all connected engines,
    without any additional headers and blank lines.

.INPUTS
    Portal FQDN
    Username
    Password
    NXQL Query
.OUTPUTS
    Merged Nexthink engines output

.NOTES
    Version:            1.0
    Author:             Stanislaw Horna
    Mail:               stanislaw.horna@atos.net
    Creation Date:      16-Feb-2023
    ChangeLog:

    Date            Who                     What
    2023-02-18      Stanislaw Horna         Basic query validation added;
                                            Mechanism to keep credentials if valid;
                                            MoJ select environment GUI form added.

    2023-02-24      Stanislaw Horna         Show password button added;
                                            Connecting to portal with ENTER key;
                                            More accurate error handling;
                                            Possibility to change user, after establishing connection.
    
    2023-02-27      Stanislaw Horna         Better Handling if unable to create neccesary variables;
                                            Handling for running on unsupported environment;
                                            Handling for no write permission;
                                            Handling for running multiple apps at the same time.
                                            
#>

$ErrorActionPreference = 'Stop'

function Invoke-main {
    Invoke-FilesVerification
    Invoke-AdditionalClasses
    try {
        PowerShell.exe -WindowStyle hidden ./main/NXQL-main.ps1
    }
    catch {
        Write-Host "Not able to run application"
        Pause
        throw
    }
}
function Invoke-FilesVerification {
    if (!(Test-Path -Path ./main)) {
        Write-Host "Some catalogs are missing"
        Pause
        throw
    }
    if (!(Test-Path -Path ./main/NXQL-main.ps1)) {
        Write-Host "Main application files are missing"
        Pause
        throw
    }
    if (!(Test-Path -Path ./main/Job-Functions.psm1)) {
        Write-Host "Main application files are missing"
        Pause
        throw
    }        
}
function Invoke-AdditionalClasses {
    try {
        Add-Type -AssemblyName PresentationCore, PresentationFramework
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Web
        [System.Windows.Forms.Application]::EnableVisualStyles()
    }
    catch {
        Write-Host "Currently running environment is not supported"
        Pause
        throw
    }
}

Invoke-main