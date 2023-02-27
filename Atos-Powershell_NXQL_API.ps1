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
    2023-02-24      Stanislaw Horna         Show password button added;
                                            Connecting to portal with ENTER key;
                                            More accurate error handling;
                                            Possibility to change user, after establishing connection.
                                            
#>
$ErrorActionPreference = 'Stop'

function Invoke-main{
    Invoke-FilesVerification
    try {
        PowerShell.exe -WindowStyle hidden ./main/NXQL-main.ps1
    }
    catch {
        throw 'Not able to run application'
        Read-Host "Press any key to close this window"
    }
}

function Invoke-FilesVerification {
    if (!(Test-Path -Path ./main)) {
        throw 'Some catalogs are missing'
        Read-Host "Press any key to close this window"
    }
    if(!(Test-Path -Path ./main/NXQL-main.ps1)){
        throw 'Main application file is missing'
        Read-Host "Press any key to close this window"
    }    
}

Invoke-main