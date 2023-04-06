
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Web
Add-Type -AssemblyName PresentationCore, PresentationFramework
[System.Windows.Forms.Application]::EnableVisualStyles()

Class NexthinkOperations {
    [String]$Portal
    [pscredential]$Credentials
    [String[]]$Platforms
    [Int]$PortNumber
    [String]$OutputFormat = "csv"
    $Engines
    [String]$Query
    [String]$SyncPath
    [String]$LogPath
    [String]$DestinationPath
    [String]$QueryToValidate
    [String]$RandomServerName
    

    [Object]GetEnginesList() {
        $web = [Net.WebClient]::new()
        $web.Credentials = $this.Credentials
        $pair = [string]::Join(":", $web.Credentials.UserName, $web.Credentials.Password)
        $base64 = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($pair))
        $web.Headers.Add('Authorization', "Basic $base64")
        $baseUrl = "https://$($this.portal)/api/configuration/v1/engines"
        $result = $web.downloadString($baseUrl)
        $engineList = $result | ConvertFrom-Json
        $engineList = $engineList | Where-Object { $_.status -eq "CONNECTED" }
        # Listing Connected Engines only
        if ($this.Portal -eq 'ministryofjustice.eu.nexthink.cloud') {
            $FITS = ('engine-1', 'engine-2', 'engine-3', 'engine-4', 'engine-5', 'engine-6', 'engine-7', 'engine-8', 'engine-9')
            $MOJO = ('engine-10', 'engine-11', 'engine-12', 'engine-13', 'engine-14')
            Invoke-FormEnvironment
            if ($this.Environment -eq "FITS") {
                $engineList = $engineList | Where-Object { $_.name -in $FITS }
            }
            elseif ($this.Environment -eq "MoJo") {
                $engineList = $engineList | Where-Object { $_.name -in $MOJO }
            }
        }
        return $engineList
    }
    [Void]ReduceQueryOutputToValidation() {
        $this.QueryToValidate = ($this.Query -replace "\(limit [0-9]+\)", "(limit 0)")
    }
    [Void]GetEngineNameToValidate() {
        $this.RandomServerName = ($this.Engines | Select-Object -First 1).address
    }
    [Bool]ValidateQuery() {
        $this.ReduceQueryOutputToValidation()
        $this.RandomServerName()
        $PlaformsString = ""
        Foreach ($platform in $this.Platforms) {
            $PlaformsString += "&platform={0}" -f $platform
        }
        $EncodedNxqlQuery = [System.Web.HttpUtility]::UrlEncode($this.QueryToValidate)
        $Url = "https://{0}:{1}/2/query?query={2}&format={3}{4}" -f $this.RandomServerName, $this.PortNumber, $EncodedNxqlQuery, $this.OuputFormat, $PlaformsString
        #echo $Url
        try {
            [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.SecurityProtocolType]::Tls11
            [Net.ServicePointManager]::ServerCertificateValidationCallback = { $true } 
            $webclient = New-Object system.net.webclient
            $webclient.Credentials = New-Object System.Net.NetworkCredential($this.Credentials.UserName, $this.Credentials.GetNetworkCredential().Password)
            $webclient.DownloadString($Url) | Out-Null
        }
        catch [System.Net.WebException] {
            try {
                $streamReader = [System.IO.StreamReader]::new($_.Exception.Response.GetResponseStream())
            }
            catch {
                if ($this.Query -like "*#*") {
                    $Error_Output = @{"Error Message" = 'The dynamic field does not exist for selected table.' }
                }
                else {
                    $Error_Output = @{"Error Message" = 'Query is invalid, error unknown.' }
                }
                return $Error_Output
            }
            $responseContent = $streamReader.ReadToEnd()
            $streamReader.Dispose()
            $HTML = New-Object -Com "HTMLFile"
            $HTML.IHTMLDocument2_write($responseContent)
            $Error_Output = @{}
            $Error_Output.Add("Error Message", ($html.getElementById("error_message").IHTMLElement_innerText))
            $Error_Output.Add("Error Options", ($html.getElementById("error_options").IHTMLElement_innerText))
            if ($Error_Output.Count -eq 0) {
                return $null
            }
            return $Error_Output
        }
        catch {
            return 'Not able to retriev data'
        }
        return $null
    }
    [Bool]ValidateQuery([String]$Type) {
        if ($Type -eq "Light") {
            $this.ReduceQueryOutputToValidation()
            if ($this.QueryToValidate.Length -le 19) {
                return $null
            }
            # Check if all opened brackets in query are closed
            if (($this.Query.ToCharArray() | Where-Object { $_ -eq '(' } | Measure-Object).Count `
                    -ne `
                ($this.QueryToValidate.ToCharArray() | Where-Object { $_ -eq ')' } | Measure-Object).Count) {
                return "Some brackets are not closed !"
            }
            # Check if query is not empty
            if ($this.QueryToValidate.Length -le 1) {
                return "NXQL query can not be blank !"
            }
            # Check if select statement exists
            if (($this.QueryToValidate -notlike "*select*")) {
                return "There is no `"select`" statement !"
            }
            # Check if from statement exists
            if (($this.QueryToValidate -notlike "*from*")) {
                return "There is no `"from`" statement !"
            }
            # Check if limit statement exists
            if (($this.QueryToValidate -notlike "*limit*")) {
                return "There is no `"limit`" statement at the end of the query!"
            }
            return $null
        }
        else {
            return $null 
        }
    }
    [Object]GetNXQLExport() {
        $Function = {
            Add-Type -AssemblyName System.Web
            Function Invoke-Nxql {
                <#
                .SYNOPSIS
                Sends an NXQL query to a Nexthink engine.
            
                .DESCRIPTION
                 Sends an NXQL query to the Web API of Nexthink Engine as HTTP GET using HTTPS.
                 
                .PARAMETER ServerName
                 Nexthink Engine name or IP address.
            
                .PARAMETER PortNumber
                Port number of the Web API (default 1671).
            
                .PARAMETER UserName
                User name of the Finder account under which the query is executed.
            
                .PARAMETER UserPassword
                User password of the Finder account under which the query is executed.
            
                .PARAMETER NxqlQuery
                NXQL query.
            
                .PARAMETER FirstParamter
                Value of %1 in the NXQL query.
            
                .PARAMETER SecondParamter
                Value of %2 in the NXQL query.
            
                .PARAMETER OuputFormat
                NXQL query output format i.e. csv, xml, html, json (default csv).
            
                .PARAMETER Platforms
                Platforms on which the query applies i.e. windows, mac_os, mobile (default windows).
                
                .EXAMPLE
                Invoke-Nxql -ServerName 176.31.63.200 -UserName "admin" -UserPassword "admin" 
                -Platforms=windows,mac_os -NxqlQuery "(select (name) (from device))"
                #>
                Param(
                    [Parameter(Mandatory = $true)]
                    [string]$ServerName,
                    [Parameter(Mandatory = $true)]
                    [System.Management.Automation.PSCredential]$credentials,
                    [Parameter(Mandatory = $true)]
                    [string]$Query,
                    [Parameter(Mandatory = $false)]
                    [int]$PortNumber = 1671,
                    [Parameter(Mandatory = $false)]
                    [string]$OuputFormat = "csv",
                    [Parameter(Mandatory = $false)]
                    [string[]]$Platforms = "windows",
                    [Parameter(Mandatory = $false)]
                    [string]$FirstParameter,
                    [Parameter(Mandatory = $false)]
                    [string]$SecondParameter
                )
                $PlaformsString = ""
                Foreach ($platform in $Platforms) {
                    $PlaformsString += "&platform={0}" -f $platform
                }
                $EncodedNxqlQuery = [System.Web.HttpUtility]::UrlEncode($Query)
                $Url = "https://{0}:{1}/2/query?query={2}&format={3}{4}" -f $ServerName, $PortNumber, $EncodedNxqlQuery, $OuputFormat, $PlaformsString
                if ($FirstParameter) { 
                    $EncodedFirstParameter = [System.Web.HttpUtility]::UrlEncode($FirstParameter)
                    $Url = "{0}&p1={1}" -f $Url, $EncodedFirstParameter
                }
                if ($SecondParameter) { 
                    $EncodedSecondParameter = [System.Web.HttpUtility]::UrlEncode($SecondParameter)
                    $Url = "{0}&p2={1}" -f $Url, $EncodedSecondParameter
                }
                #echo $Url
                try {
                    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.SecurityProtocolType]::Tls11
                    [Net.ServicePointManager]::ServerCertificateValidationCallback = { $true } 
                    $webclient = New-Object system.net.webclient
                    $webclient.Credentials = New-Object System.Net.NetworkCredential($Credentials.UserName, $credentials.GetNetworkCredential().Password)
                    $webclient.DownloadString($Url)
                }
                catch [System.Net.WebException] {
                    $streamReader = [System.IO.StreamReader]::new($_.Exception.Response.GetResponseStream())
                    $responseContent = $streamReader.ReadToEnd()
                    $streamReader.Dispose()
                    $HTML = New-Object -Com "HTMLFile"
                    $HTML.IHTMLDocument2_write($responseContent)
                    throw ($html.getElementById("error_message").IHTMLElement_innerText)
                }
                catch {
                    throw 'Not able to retriev data'
                }
            }
        }
        if (Test-Path -Path "$($this.SyncPath)\Headers") {
            Remove-Item -Path "$($this.SyncPath)\Headers" -Confirm:$false -Force
        }
        if (Test-Path -Path "$($this.SyncPath)\Wait") {
            Remove-Item -Path "$($this.SyncPath)\Wait" -Confirm:$false -Force
        }
        if (Test-Path -Path $this.LogPath) {
            $LogsToDelete = (Get-ChildItem -Path $this.LogPath -Filter *.csv).FullName
            foreach ($file in $LogsToDelete) {
                Remove-Item -Path $file -Confirm:$false -Force
            }
        }
        if (Test-Path -Path "$($this.LogPath)\BadRequest") {
            Remove-Item -Path "$($this.LogPath)\BadRequest" -Confirm:$false -Force
        }
        # Create separate process for each engine in scope
        foreach ($Engine in $this.Engines) {
            $Name = "NXQL-" + $Engine.name
            $RandomWaitTime = Get-Random -Minimum 0 -Maximum 500
            Start-Job -Name $Name `
                -InitializationScript $Function `
                -ScriptBlock {
                param(
                    $EngineAddress,
                    $WebAPIPort,
                    $Platform,
                    [PSCredential]$credentials,
                    $Query,
                    $DestinationPath,
                    $SyncPath,
                    $RandomWaitTime,
                    $LogPath
                )
                "$(([System.DateTime]::Now).ToString('HH\:mm\:ss\.fff')) Thread Started" | Out-File "$LogPath\Log-$EngineAddress.csv" -Append
                $ErrorActionPreference = 'Stop'
                # Retrieve data from Nexthink
                try {
                    "$(([System.DateTime]::Now).ToString('HH\:mm\:ss\.fff')) Trying to retrieve data from Nexthink" | Out-File "$LogPath\Log-$EngineAddress.csv" -Append
                    $out = Invoke-Nxql `
                        -ServerName $EngineAddress `
                        -PortNumber $WebAPIPort `
                        -credentials $credentials `
                        -Query $Query `
                        -Platforms $Platform
                }
                catch {
                    $_.Exception.Message | Out-File "$LogPath\Log-$EngineAddress.csv" -Append
                    $_.Exception.Message | Out-File "$LogPath\BadRequest"
                    throw "Not able to collect data"
                }
                # Split to be able to remove unnecesary headers
                $out = $out.Split("`n")
                # Close job if the output from particular Engine is empty
                if ($out.count -le 2) {
                    "$(([System.DateTime]::Now).ToString('HH\:mm\:ss\.fff')) Output was empty" | Out-File -FilePath "$LogPath\Log-$EngineAddress.csv" -Append
                    return
                }
                # Check file lock and set File lock
                $Wait = $true
                while ($Wait) {
                    if (!(Test-Path -Path "$SyncPath\Wait")) {
                        try {
                            "$(([System.DateTime]::Now).ToString('HH\:mm\:ss\.fff')) Trying to set wait file" | Out-File -FilePath "$LogPath\Log-$EngineAddress.csv" -Append
                            New-Item -Path "$SyncPath\Wait"
                            "$(([System.DateTime]::Now).ToString('HH\:mm\:ss\.fff')) Wait file is set" | Out-File -FilePath "$LogPath\Log-$EngineAddress.csv" -Append
                            $Wait = $false
                        }
                        catch {
                            if ($Error.Exception.Message -like "*already exists.*") {
                                "$(([System.DateTime]::Now).ToString('HH\:mm\:ss\.fff')) Two threads tried to set wait file" | Out-File -FilePath "$LogPath\Log-$EngineAddress.csv" -Append
                            }
                            else {
                                "$(([System.DateTime]::Now).ToString('HH\:mm\:ss\.fff')) Not able to sync threads (error unknown)" | Out-File -FilePath "$LogPath\Log-$EngineAddress.csv" -Append
                                $_.Exception | Out-File "$LogPath\Log-$EngineAddress.csv" -Append
                            }
                        }
                    }
                    Start-Sleep -Milliseconds $RandomWaitTime
                }
                # Check if headers alredy exist in the output file
                if (Test-Path -Path "$SyncPath\Headers") {
                    # Append output file without headers
                    $out[1..($out.count - 2)] | Out-File $DestinationPath -Append
                    "$(([System.DateTime]::Now).ToString('HH\:mm\:ss\.fff')) Data was written to the result file" | Out-File "$LogPath\Log-$EngineAddress.csv" -Append
                    "$(([System.DateTime]::Now).ToString('HH\:mm\:ss\.fff')) $EngineAddress written data to the result file without headers" | Out-File "$LogPath\Log-Write-Order.csv" -Append
                }
                else {
                    # Set flag for headers
                    "$(([System.DateTime]::Now).ToString('HH\:mm\:ss\.fff')) Setting headers file" | Out-File "$LogPath\Log-$EngineAddress.csv" -Append
                    New-Item -Path "$SyncPath\Headers" 
                    # Write output file with headers
                    $out[0..($out.count - 2)] | Out-File $DestinationPath
                    "$(([System.DateTime]::Now).ToString('HH\:mm\:ss\.fff')) Data was written to the result file with headers" | Out-File "$LogPath\Log-$EngineAddress.csv" -Append
                    "$(([System.DateTime]::Now).ToString('HH\:mm\:ss\.fff')) $EngineAddress written data to the result file with headers" | Out-File "$LogPath\Log-Write-Order.csv" -Append
                }
                # Remove file lock
                "$(([System.DateTime]::Now).ToString('HH\:mm\:ss\.fff')) Removing wait file" | Out-File -FilePath "$LogPath\Log-$EngineAddress.csv" -Append
                Remove-Item -Path "$SyncPath\Wait" -Confirm:$false -Force
                "$(([System.DateTime]::Now).ToString('HH\:mm\:ss\.fff')) Thread Completed successfully" | Out-File "$LogPath\Log-$EngineAddress.csv" -Append
            } -ArgumentList $Engine.address, $this.PortNumber, $this.Platform, $this.Credentials, $this.Query, $this.DestinationPath, $this.SyncPath, $RandomWaitTime, $this.LogPath | Out-Null
        }
        $CompletedJobsCounter = 0
        # Wait until all jobs will be done
        while ((Get-Job -Name "NXQL*").State -contains "Running") {
            [System.Windows.Forms.Application]::DoEvents()
        }
        # Dynamically remove completed jobs
        $CompletedJobsCounter += (Get-Job -Name "NXQL*" | Where-Object { $_.State -eq "Completed" }).Count
        Get-Job -Name "NXQL*" | Remove-Job
        # Clear Thread synchronization files
        if (Test-Path -Path "$($this.SyncPath)\Headers") {
            Remove-Item -Path "$($this.SyncPath)\Headers" -Confirm:$false -Force
        }
        if (Test-Path -Path "$($this.SyncPath)\Wait") {
            Remove-Item -Path "$($this.SyncPath)\Wait" -Confirm:$false -Force
        }
        # Handling for environments with only one engine
        if ($null -eq $this.Engines.count) {
            $Number_of_engines = 1
        }
        else {
            $Number_of_engines = $this.Engines.Count
        }
        # Check if outputs from all engines are pasted to the result file
        if ($CompletedJobsCounter -eq $Number_of_engines) {
            return "Success!"
        }
        elseif (Test-Path "$($this.LogPath)\BadRequest") {
            $ErrorMessage = Get-Content -Path "$($this.LogPath)\BadRequest"
            return $ErrorMessage
        }
        else {
            return "Failed: Error unknown"
        }
    }
}
Class GUIComponents {

    [Object]Label (
        [String]$Text,
        [int]$Location_X,
        [int]$Location_Y,
        [String]$Font,
        [int]$FontSize,
        [bool]$Bold,
        [bool]$Visible,
        [MainGUIForm]$GUIFormClass) {

        $Label = New-Object system.Windows.Forms.Label
        $Label.text = $Text
        $Label.AutoSize = $true
        $Label.width = 25
        $Label.height = 10
        $Label.location = New-Object System.Drawing.Point($Location_X, $Location_Y)
        $Label.Font = New-Object System.Drawing.Font($Font, $FontSize)
        $GUIFormClass.Form.Controls.Add($Label)
        return $Label
    }

    [Object]CheckBox(
        [String]$Text,
        [int]$Location_X,
        [int]$Location_Y,
        [int]$Size_X,
        [int]$Size_Y,
        [bool]$Checked,
        [bool]$Visible,
        [scriptblock]$Action,
        [MainGUIForm]$GUIFormClass) {

        $Checkbox = New-Object System.Windows.Forms.Checkbox 
        $Checkbox.Location = New-Object System.Drawing.Size($Location_X, $Location_Y) 
        $Checkbox.Size = New-Object System.Drawing.Size($Size_X, $Size_Y)
        $Checkbox.Text = $Text
        $Checkbox.checked = $Checked
        $Checkbox.Visible = $Visible
        $Checkbox.TabIndex = 4
        $Checkbox.Add_Click($Action)
        $GUIFormClass.Form.Controls.Add($Checkbox)
        return $Checkbox
    }

}

Class MainGUIForm : NexthinkOperations {
    # Statuses
    [bool]$ValidQuery
    [bool]$AvailableOptionsAre
    [bool]$KeepCredentials
    # Variables
    $InputFolder
    # Main Form
    $Form
    # Labels
    $LabelPortal
    $LabelLogin
    $LabelPassword
    $LabelConnectionStatus
    $LabelConnectionStatusDetails 
    $LabelNumberOfEngines
    $LabelQuery
    $LabelPort
    $LabelFileName
    $LabelPath
    $LabelRunStatus
    $LabelPlatform
    $LabelLookup
    $LabelOptionsAre
    # Checkboxes
    $CheckboxWindows
    $CheckboxMac_OS
    $CheckboxMobile
    $CheckboxShowPassword
    # Boxes
    $BoxPortal
    $BoxLogin
    $BoxPassword
    $BoxQuery
    $BoxFileName
    $BoxPath
    $BoxPort
    $BoxLookfor
    $BoxErrorOptions
    # Buttons
    $ButtonConnect
    $ButtonPath
    $ButtonWebEditor
    $ButtonValidateQuery
    $ButtonRunQuery

    MainGUIForm() {
        ######################################################################
        #----------------------- GUI Forms Definition -----------------------#
        ######################################################################
        $this.Form = New-Object system.Windows.Forms.Form
        $this.Form.ClientSize = New-Object System.Drawing.Point(480, 150)
        $this.Form.text = "Powershell NXQL API"
        $this.Form.FormBorderStyle = 'FixedDialog'
        $this.Form.TopMost = $true
        # Handling if opened via Powershell ISE
        try {
            $p = (Get-Process powershell | Sort-Object -Property CPU -Descending | Select-Object -First 1).Path
            $this.Form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($p)
        }
        catch {
            $p = (Get-Process explorer | Sort-Object -Property CPU -Descending | Select-Object -First 1).Path
            $this.Form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($p)
        }
        $this.Form.Add_FormClosing({
                Remove-Item -Path $this.UniqueInstanceLock -Confirm:$false -Force
            })

        ######################################################################
        #-------------------------- Labels Section --------------------------#
        ######################################################################

        $this.LabelPortal = New-Object system.Windows.Forms.Label
        $this.LabelPortal.text = "Portal FQDN: "
        $this.LabelPortal.AutoSize = $true
        $this.LabelPortal.width = 25
        $this.LabelPortal.height = 10
        $this.LabelPortal.location = New-Object System.Drawing.Point(20, 0)
        $this.LabelPortal.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 12)
        $this.Form.Controls.Add($this.LabelPortal)

        $this.LabelLogin = New-Object system.Windows.Forms.Label
        $this.LabelLogin.text = "Login:"
        $this.LabelLogin.AutoSize = $true
        $this.LabelLogin.width = 25
        $this.LabelLogin.height = 10
        $this.LabelLogin.location = New-Object System.Drawing.Point(20, 30)
        $this.LabelLogin.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 12)
        $this.Form.Controls.Add($this.LabelLogin)

        $this.LabelPassword = New-Object system.Windows.Forms.Label
        $this.LabelPassword.text = "Password:"
        $this.LabelPassword.AutoSize = $true
        $this.LabelPassword.width = 25
        $this.LabelPassword.height = 10
        $this.LabelPassword.location = New-Object System.Drawing.Point(20, 60)
        $this.LabelPassword.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 12)
        $this.Form.Controls.Add($this.LabelPassword)

        $this.LabelConnectionStatus = New-Object system.Windows.Forms.Label
        $this.LabelConnectionStatus.text = ""
        $this.LabelConnectionStatus.AutoSize = $true
        $this.LabelConnectionStatus.width = 25
        $this.LabelConnectionStatus.height = 10
        $this.LabelConnectionStatus.location = New-Object System.Drawing.Point(150, 105)
        $this.LabelConnectionStatus.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)
        $this.LabelConnectionStatus.Visible = $false
        $this.Form.Controls.Add($this.LabelConnectionStatus)

        $this.LabelConnectionStatusDetails = New-Object system.Windows.Forms.Label
        $this.LabelConnectionStatusDetails.text = ""
        $this.LabelConnectionStatusDetails.AutoSize = $true
        $this.LabelConnectionStatusDetails.width = 25
        $this.LabelConnectionStatusDetails.height = 10
        $this.LabelConnectionStatusDetails.location = New-Object System.Drawing.Point(260, 105)
        $this.LabelConnectionStatusDetails.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10, [System.Drawing.FontStyle]::Bold)
        $this.LabelConnectionStatusDetails.Visible = $true
        $this.Form.Controls.Add($this.LabelConnectionStatusDetails)

        $this.LabelNumberOfEngines = New-Object system.Windows.Forms.Label
        $this.LabelNumberOfEngines.text = "Number of engines: "
        $this.LabelNumberOfEngines.AutoSize = $true
        $this.LabelNumberOfEngines.width = 25
        $this.LabelNumberOfEngines.height = 10
        $this.LabelNumberOfEngines.location = New-Object System.Drawing.Point(350, 105)
        $this.LabelNumberOfEngines.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)
        $this.LabelNumberOfEngines.Visible = $false
        $this.Form.Controls.Add($this.LabelNumberOfEngines)

        $this.LabelQuery = New-Object system.Windows.Forms.Label
        $this.LabelQuery.text = "NXQL query:"
        $this.LabelQuery.AutoSize = $true
        $this.LabelQuery.width = 25
        $this.LabelQuery.height = 10
        $this.LabelQuery.location = New-Object System.Drawing.Point(20, 140)
        $this.LabelQuery.Visible = $false
        $this.LabelQuery.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 12)
        $this.Form.Controls.Add($this.LabelQuery)

        $this.LabelPort = New-Object system.Windows.Forms.Label
        $this.LabelPort.text = "NXQL port:"
        $this.LabelPort.AutoSize = $true
        $this.LabelPort.width = 25
        $this.LabelPort.height = 10
        $this.LabelPort.location = New-Object System.Drawing.Point(550, 140)
        $this.LabelPort.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 12)
        $this.LabelPort.Visible = $false
        $this.Form.Controls.Add($this.LabelPort)

        $this.LabelFileName = New-Object System.Windows.Forms.Label
        $this.LabelFileName.Text = "Destination File Name"
        $this.LabelFileName.AutoSize = $true
        $this.LabelFileName.width = 25
        $this.LabelFileName.height = 10
        $this.LabelFileName.location = New-Object System.Drawing.Point(20, 480)
        $this.LabelFileName.Visible = $false
        $this.LabelFileName.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 12)
        $this.Form.Controls.Add($this.LabelFileName)

        $this.LabelPath = New-Object System.Windows.Forms.Label
        $this.LabelPath.Text = "Destination Path"
        $this.LabelPath.AutoSize = $true
        $this.LabelPath.width = 25
        $this.LabelPath.height = 10
        $this.LabelPath.location = New-Object System.Drawing.Point(20, 510)
        $this.LabelPath.Visible = $false
        $this.LabelPath.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 12)
        $this.Form.Controls.Add($this.LabelPath)

        $this.LabelRunStatus = New-Object System.Windows.Forms.Label
        $this.LabelRunStatus.Text = ""
        $this.LabelRunStatus.AutoSize = $true
        $this.LabelRunStatus.TextAlign = "MiddleCenter"
        $this.LabelRunStatus.MaximumSize = New-Object System.Drawing.Size(435, 60)
        $this.LabelRunStatus.width = 25
        $this.LabelRunStatus.height = 10
        $this.LabelRunStatus.location = New-Object System.Drawing.Point(150, 540)
        $this.LabelRunStatus.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10, [System.Drawing.FontStyle]::Bold)
        $this.LabelRunStatus.Visible = $false
        $this.Form.Controls.Add($this.LabelRunStatus)

        $this.LabelPlatform = New-Object System.Windows.Forms.Label
        $this.LabelPlatform.Text = "Select Platform:"
        $this.LabelPlatform.AutoSize = $true
        $this.LabelPlatform.width = 25
        $this.LabelPlatform.height = 10
        $this.LabelPlatform.location = New-Object System.Drawing.Point(550, 0)
        $this.LabelPlatform.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 12)
        $this.LabelPlatform.Visible = $false
        $this.Form.Controls.Add($this.LabelPlatform)

        $this.LabelLookup = New-Object System.Windows.Forms.Label
        $this.LabelLookup.Text = "Look for:"
        $this.LabelLookup.AutoSize = $true
        $this.LabelLookup.width = 25
        $this.LabelLookup.height = 10
        $this.LabelLookup.location = New-Object System.Drawing.Point(695, 0)
        $this.LabelLookup.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 12)
        $this.LabelLookup.Visible = $false
        $this.Form.Controls.Add($this.LabelLookup)

        $this.LabelOptionsAre = New-Object System.Windows.Forms.Label
        $this.LabelOptionsAre.Text = "Options are:"
        $this.LabelOptionsAre.AutoSize = $true
        $this.LabelOptionsAre.width = 25
        $this.LabelOptionsAre.height = 10
        $this.LabelOptionsAre.location = New-Object System.Drawing.Point(695, 35)
        $this.LabelOptionsAre.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 12)
        $this.LabelOptionsAre.Visible = $false
        $this.Form.Controls.Add($this.LabelOptionsAre)

        ######################################################################
        #------------------------ Checkboxes Section ------------------------#
        ######################################################################
        $this.CheckboxWindows = [GUIComponents]::CheckBox(
            "Windows",
            555,
            25,
            100,
            20,
            $true,
            $false,
            {Invoke-CheckboxPlatform},
            $this
        )
        # $this.CheckboxWindows = New-Object System.Windows.Forms.Checkbox 
        # $this.CheckboxWindows.Location = New-Object System.Drawing.Size(555, 25) 
        # $this.CheckboxWindows.Size = New-Object System.Drawing.Size(100, 20)
        # $this.CheckboxWindows.Text = "Windows"
        # $this.CheckboxWindows.checked = $true
        # $this.CheckboxWindows.Visible = $false
        # $this.CheckboxWindows.TabIndex = 4
        # $this.CheckboxWindows.Add_Click({ Invoke-CheckboxPlatform })
        # $this.Form.Controls.Add($this.CheckboxWindows)

        $this.CheckboxMac_OS = New-Object System.Windows.Forms.Checkbox 
        $this.CheckboxMac_OS.Location = New-Object System.Drawing.Size(555, 45) 
        $this.CheckboxMac_OS.Size = New-Object System.Drawing.Size(100, 20)
        $this.CheckboxMac_OS.Text = "Mac OS"
        $this.CheckboxMac_OS.Visible = $false
        $this.CheckboxMac_OS.TabIndex = 4
        $this.CheckboxMac_OS.Add_Click({ Invoke-CheckboxPlatform })
        $this.Form.Controls.Add($this.CheckboxMac_OS)

        $this.CheckboxMobile = New-Object System.Windows.Forms.Checkbox 
        $this.CheckboxMobile.Location = New-Object System.Drawing.Size(555, 65) 
        $this.CheckboxMobile.Size = New-Object System.Drawing.Size(100, 20)
        $this.CheckboxMobile.Text = "Mobile"
        $this.CheckboxMobile.Visible = $false
        $this.CheckboxMobile.TabIndex = 4
        $this.CheckboxMobile.Add_Click({ Invoke-CheckboxPlatform })
        $this.Form.Controls.Add($this.CheckboxMobile)

        $this.CheckboxShowPassword = New-Object System.Windows.Forms.Checkbox 
        $this.CheckboxShowPassword.Location = New-Object System.Drawing.Size(140, 80) 
        $this.CheckboxShowPassword.Size = New-Object System.Drawing.Size(200, 20)
        $this.CheckboxShowPassword.Text = "Show Password"
        $this.CheckboxShowPassword.Visible = $true
        $this.CheckboxShowPassword.TabIndex = 4
        $this.CheckboxShowPassword.checked = $false
        $this.CheckboxShowPassword.Add_Click({ Invoke-CheckboxShowPassword })
        $this.Form.Controls.Add($this.CheckboxShowPassword)

        ######################################################################
        #------------------------- TextBoxes Section ------------------------#
        ######################################################################

        $this.BoxPortal = New-Object System.Windows.Forms.TextBox 
        $this.BoxPortal.Multiline = $false
        $this.BoxPortal.Location = New-Object System.Drawing.Size(140, 0) 
        $this.BoxPortal.Size = New-Object System.Drawing.Size(300, 20)
        $this.BoxPortal.Add_TextChanged({ Invoke-BoxPortal })
        $this.Form.Controls.Add($this.BoxPortal)

        $this.BoxLogin = New-Object System.Windows.Forms.TextBox 
        $this.BoxLogin.Multiline = $false
        $this.BoxLogin.Location = New-Object System.Drawing.Size(140, 30) 
        $this.BoxLogin.Size = New-Object System.Drawing.Size(300, 20)
        $this.BoxLogin.Add_TextChanged({ Invoke-CredentialCleanup })
        $this.Form.Controls.Add($this.BoxLogin)

        $this.BoxPassword = New-Object System.Windows.Forms.MaskedTextBox 
        $this.BoxPassword.passwordchar = "*"
        $this.BoxPassword.Multiline = $false
        $this.BoxPassword.Location = New-Object System.Drawing.Size(140, 60) 
        $this.BoxPassword.Size = New-Object System.Drawing.Size(300, 20)
        $this.Form.Controls.Add($this.BoxPassword)

        $this.BoxQuery = New-Object System.Windows.Forms.TextBox 
        $this.BoxQuery.Multiline = $true
        $this.BoxQuery.Location = New-Object System.Drawing.Size(20, 165) 
        $this.BoxQuery.Size = New-Object System.Drawing.Size(660, 300)
        $this.BoxQuery.Scrollbars = 'Vertical'
        $this.BoxQuery.Visible = $false
        $this.BoxQuery.Add_TextChanged({ Invoke-BoxQuery })
        $this.Form.Controls.Add($this.BoxQuery)

        $this.BoxFileName = New-Object System.Windows.Forms.TextBox 
        $this.BoxFileName.Multiline = $false
        $this.BoxFileName.Location = New-Object System.Drawing.Size(190, 480) 
        $this.BoxFileName.Size = New-Object System.Drawing.Size(495, 20)
        $this.BoxFileName.Text = (Get-Date).ToString("yyyy-MM-dd")
        $this.BoxFileName.Visible = $false
        $this.Form.Controls.Add($this.BoxFileName)

        $this.BoxPath = New-Object System.Windows.Forms.TextBox 
        $this.BoxPath.Multiline = $false
        $this.BoxPath.Location = New-Object System.Drawing.Size(150, 510) 
        $this.BoxPath.Size = New-Object System.Drawing.Size(430, 20)
        $this.BoxPath.Text = Get-BoxPathLocation
        $this.BoxPath.Visible = $false
        $this.Form.Controls.Add($this.BoxPath)

        $this.BoxPort = New-Object System.Windows.Forms.TextBox 
        $this.BoxPort.Multiline = $false
        $this.BoxPort.Location = New-Object System.Drawing.Size(640, 140) 
        $this.BoxPort.Size = New-Object System.Drawing.Size(40, 20)
        $this.BoxPort.Text = ""
        $this.BoxPort.Visible = $false
        $this.Form.Controls.Add($this.BoxPort)

        $this.BoxLookfor = New-Object System.Windows.Forms.TextBox 
        $this.BoxLookfor.Text = ""
        $this.BoxLookfor.AutoSize = $true
        $this.BoxLookfor.width = 25
        $this.BoxLookfor.height = 10
        $this.BoxLookfor.location = New-Object System.Drawing.Point(765, 0)
        $this.BoxLookfor.Size = New-Object System.Drawing.Size(315, 20)
        $this.BoxLookfor.Visible = $false
        $this.BoxLookfor.Add_TextChanged({ Invoke-BoxLookFor })
        $this.Form.Controls.Add($this.BoxLookfor)

        $this.BoxErrorOptions = New-Object System.Windows.Forms.TextBox 
        $this.BoxErrorOptions.Text = ""
        $this.BoxErrorOptions.AutoSize = $true
        $this.BoxErrorOptions.Multiline = $true
        $this.BoxErrorOptions.Scrollbars = 'Vertical'
        $this.BoxErrorOptions.ReadOnly = $true
        $this.BoxErrorOptions.width = 25
        $this.BoxErrorOptions.height = 10
        $this.BoxErrorOptions.location = New-Object System.Drawing.Point(700, 60)
        $this.BoxErrorOptions.Size = New-Object System.Drawing.Size(380, 525)
        $this.BoxErrorOptions.Visible = $false
        $this.Form.Controls.Add($this.BoxErrorOptions)

        ######################################################################
        #------------------------- Buttons Section --------------------------#
        ######################################################################

        $this.ButtonConnect = New-Object System.Windows.Forms.Button
        $this.ButtonConnect.Location = New-Object System.Drawing.Point(20, 100)
        $this.ButtonConnect.Size = New-Object System.Drawing.Size(100, 30)
        $this.ButtonConnect.Text = 'Connect to portal'
        $this.ButtonConnect.Add_Click({ Invoke-ButtonConnectToPortal })
        $this.Form.Controls.Add($this.ButtonConnect)

        $this.ButtonPath = New-Object System.Windows.Forms.Button
        $this.ButtonPath.Location = New-Object System.Drawing.Point(585, 505)
        $this.ButtonPath.Size = New-Object System.Drawing.Size(100, 30)
        $this.ButtonPath.Text = 'Select'
        $this.ButtonPath.Visible = $false
        $this.ButtonPath.Add_Click({ $this.BoxPath.Text = Invoke-ButtonSelectPath -inputFolder $this.BoxPath.Text })
        $this.Form.Controls.Add($this.ButtonPath)

        $this.ButtonWebEditor = New-Object System.Windows.Forms.Button
        $this.ButtonWebEditor.Location = New-Object System.Drawing.Point(585, 540)
        $this.ButtonWebEditor.Size = New-Object System.Drawing.Size(100, 50)
        $this.ButtonWebEditor.Visible = $false
        $this.ButtonWebEditor.Text = 'Open Web Query Editor'
        $this.ButtonWebEditor.Add_Click({ Invoke-ButtonQueryWebEditor })
        $this.Form.Controls.Add($this.ButtonWebEditor)

        $this.ButtonValidateQuery = New-Object System.Windows.Forms.Button
        $this.ButtonValidateQuery.Location = New-Object System.Drawing.Point(20, 540)
        $this.ButtonValidateQuery.Size = New-Object System.Drawing.Size(120, 50)
        $this.ButtonValidateQuery.Text = 'Validate Query'
        $this.ButtonValidateQuery.Visible = $false
        $this.ButtonValidateQuery.ForeColor = 'orange'
        $this.ButtonValidateQuery.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10, [System.Drawing.FontStyle]::Bold)
        $this.ButtonValidateQuery.Add_Click({ Invoke-ButtonValidateQuery })
        $this.Form.Controls.Add($this.ButtonValidateQuery)

        $this.ButtonRunQuery = New-Object System.Windows.Forms.Button
        $this.ButtonRunQuery.Location = New-Object System.Drawing.Point(20, 540)
        $this.ButtonRunQuery.Size = New-Object System.Drawing.Size(120, 50)
        $this.ButtonRunQuery.Text = 'Run NXQL Query'
        $this.ButtonRunQuery.Visible = $false
        $this.ButtonRunQuery.ForeColor = 'green'
        $this.ButtonRunQuery.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10, [System.Drawing.FontStyle]::Bold)
        $this.ButtonRunQuery.Add_Click({ Invoke-ButtonRunNXQLQuery })
        $this.Form.Controls.Add($this.ButtonRunQuery)
    }
    BigGUI() {
        $this.Form.TopMost = $false
        $this.Form.ClientSize = New-Object System.Drawing.Point(700, 600)
        $this.BoxErrorOptions.visible = $False
        $this.LabelQuery.Visible = $true
        $this.BoxQuery.Visible = $true
        $this.LabelFileName.Visible = $true
        $this.BoxFileName.Visible = $true
        $this.LabelPath.Visible = $true
        $this.BoxPath.Visible = $true
        $this.ButtonPath.Visible = $true
        $this.LabelPlatform.Visible = $true
        $this.CheckboxWindows.Visible = $true
        $this.CheckboxMac_OS.Visible = $true
        $this.CheckboxMobile.Visible = $true
        $this.ButtonWebEditor.Visible = $true
        $this.LabelPort.Visible = $true
        $this.BoxPort.Visible = $true
        if ($this.ValidQuery) {
            $this.ButtonValidateQuery.visible = $false
            $this.ButtonRunQuery.visible = $true
        }
        else {
            $this.ButtonRunQuery.visible = $false
            $this.ButtonValidateQuery.visible = $true
        }
        if ($this.AvailableOptionsAre) {
            $this.Form.TopMost = $false
            $this.Form.ClientSize = New-Object System.Drawing.Point(1100, 600)
            $this.BoxErrorOptions.visible = $true
        }
    }
    SmallGUI() {
        $this.Form.TopMost = $true
        $this.Form.ClientSize = New-Object System.Drawing.Point(480, 150)
        $this.LabelQuery.Visible = $false
        $this.BoxQuery.Visible = $false
        $this.LabelFileName.Visible = $false
        $this.BoxFileName.Visible = $false
        $this.LabelPath.Visible = $false
        $this.BoxPath.Visible = $false
        $this.ButtonPath.Visible = $false
        $this.LabelPlatform.Visible = $false
        $this.CheckboxWindows.Visible = $false
        $this.CheckboxMac_OS.Visible = $false
        $this.CheckboxMobile.Visible = $false
        $this.ButtonWebEditor.Visible = $false
        $this.LabelNumberOfEngines.Visible = $false
        $this.ButtonRunQuery.Visible = $false
        $this.LabelRunStatus.Visible = $false
        $this.LabelPort.Visible = $false
        $this.BoxPort.Visible = $false
    }
    GUI_FailLoginMessage([String]$Message) {
        $this.BoxPassword.text = ""
        $this.LabelConnectionStatus.Text = "Login and password can not be empty !"
        $this.LabelConnectionStatus.ForeColor = "red"
        $this.LabelConnectionStatus.Visible = $true
        $this.BoxPassword.text = ""
    }
    GUI_SuccessLoginMessage() {
        # Handling for the environments with only one engine
        if ($null -eq $this.Engines.count) {
            $Number_of_engines = 1
        }
        else {
            $Number_of_engines = $this.Engines.Count
        }
        # Set state and number of Engines
        $this.LabelConnectionStatusDetails.Visible = $true
        $this.LabelConnectionStatusDetails.text = "Connected"
        $this.LabelConnectionStatusDetails.ForeColor = "green"
        $this.LabelNumberOfEngines.text = "Number of engines: $Number_of_engines"
        $this.BigGUI()
    }
    GUI_ClearConnectionStatus() {
        $this.LabelConnectionStatus.Visible = $true
        $this.LabelNumberOfEngines.Visible = $false
        $this.LabelConnectionStatus.Text = "Connection state:"
        $this.LabelConnectionStatus.ForeColor = "black"
        $this.LabelNumberOfEngines.text = ""
        $this.LabelConnectionStatusDetails.text = "-"
        $this.LabelConnectionStatusDetails.ForeColor = "black"
    }
    EnableButtons() {
        $this.ButtonConnect.enabled = $true
        $this.ButtonRunQuery.enabled = $true
        $this.ButtonPath.enabled = $true
        $this.BoxPath.enabled = $true
        $this.BoxFileName.enabled = $true
        $this.BoxPort.enabled = $true
        $this.CheckboxWindows.enabled = $true
        $this.CheckboxMac_OS.enabled = $true
        $this.CheckboxMobile.enabled = $true
    }
    DisableButtons() {
        $this.ButtonConnect.enabled = $false
        $this.ButtonRunQuery.enabled = $false
        $this.ButtonPath.enabled = $false
        $this.BoxPath.enabled = $false
        $this.BoxFileName.enabled = $false
        $this.BoxPort.enabled = $false
        $this.CheckboxWindows.enabled = $false
        $this.CheckboxMac_OS.enabled = $false
        $this.CheckboxMobile.enabled = $false
    }
    ShowPassword() {
        if ($this.CheckboxShowPassword.checked) {
            $this.BoxPassword.passwordchar = $null
        }
        else {
            $this.BoxPassword.passwordchar = "*"
        }
    }
    ShowQueryOptions() {
        Invoke-FormMainResize -Big -Options
        $this.LabelLookup.visible = $true
        $this.BoxLookfor.visible = $true
        $this.BoxErrorOptions.visible = $true
        $this.LabelOptionsAre.Visible = $true
        $ErrorOptions = (($this.ErrorInformation."Error Options").Split("`n") | Select-Object -Skip 1) -join "`n" 
        $this.BoxErrorOptions.Text = $ErrorOptions
        $ErrorMessage = $this.ErrorInformation."Error Message"
        $this.BoxLookfor.text = ($ErrorMessage.Split("'*'")[1])
    }
    CleanupCredentialsOnUsernameChange() {
        $this.SmallGUI()
        $this.ButtonRunQuery.Visible = $false
        $this.ButtonWebEditor.Visible = $false
        $this.LabelConnectionStatus.Visible = $false
        $this.LabelConnectionStatusDetails.Visible = $false
        $this.LabelNumberOfEngines.Visible = $false
        $this.LabelRunStatus.Visible = $false
        $this.BoxPassword.text = ""
        $this.LabelConnectionStatusDetails.Text = ""
    }
    CleanupCredentialsOnPortalChange() {
        $this.SmallGUI()
        $this.ButtonRunQuery.Visible = $false
        $this.ButtonWebEditor.Visible = $false
        $this.LabelConnectionStatus.Visible = $false
        $this.LabelConnectionStatusDetails.Visible = $false
        $this.LabelNumberOfEngines.Visible = $false
        $this.LabelRunStatus.Visible = $false
        $this.BoxPassword.text = ""
        $this.BoxLogin.text = "" 
    }

    ActionButtonConnect() {
        # Remember password while Button is clicked multiple times without changing anything
        $this.KeepCredentials = $false
        if ($this.LabelConnectionStatusDetails.Text -eq "Connected" -and 
            $this.PortalFQDN -eq $this.BoxPortal.Text -and
            $this.Login -eq $this.BoxLogin.Text) {
            $this.KeepCredentials = $true
        }
        else {
            $this.PortalFQDN = $this.BoxPortal.text
            $this.Login = $this.BoxLogin.Text
        }
        # Clear additional info to hide any information from previous run, which maybe misleading
        # Hide additional fields if button clicked multiple times
        $this.SmallGUI()
        # Fill in the first part of output name and format portal connection details
        $this.BoxFileName.Text = (Get-Date).ToString("yyyy-MM-dd")
        # Check if fields are not empty if yes exit
        if ($this.BoxLogin.Text.Length -lt 1 -or
            $this.BoxPassword.Text.Length -lt 1) {
            $this.GUI_FailLoginMessage("Login and password can not be empty !")
            return
        }
        # Clear additional info to hide any information from previous run, which maybe misleading
        $this.GUI_ClearConnectionStatus()
        $this.Portal = $this.BoxPortal.Text
        # if login portal etc. are not changed and worked previously
        # use those credentials once agin
        if ($this.KeepCredentials -eq $false) {
            $Username = $this.BoxLogin.text
            $Password = $this.BoxPassword.text
            $Password = ConvertTo-SecureString $Password -AsPlainText -Force
            $this.Credentials = New-Object System.Management.Automation.PSCredential ($Username, $Password)
        }
        $this.BoxPassword.text = "********************"
        try {
            $this.Engines = Get-EngineList -portal $this.Portal -credentials $this.Credentials
        }
        catch {
            $this.GUI_FailLoginMessage("Not connected")
            return
        }
        # Display additional fields
        $this.GUI_SuccessLoginMessage()

        $this.LabelNumberOfEngines.Visible = $true
        $this.ButtonRunQuery.Visible = $true
        # Check environment type SAAS / On-prem
        # Set additional details based on it
        if ($this.BoxPortal.Text -notlike "*.nexthink.cloud" ) {
            $this.BoxPort.Text = "1671"
            $this.BoxFileName.Text += " - <Customer_Name> - "
        }
        else {
            $this.CustomerName = $this.BoxPortal.Text.Split(".")[0]
            $this.BoxPort.Text = "443"
            $this.BoxFileName.Text += " - $this.CustomerName - "
        }
    }
    ActionButtonPath() {
        [Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
        $browse = New-Object System.Windows.Forms.FolderBrowserDialog
        $browse.SelectedPath = $this.InputFolder
        $browse.ShowNewFolderButton = $true
        $browse.Description = "Select a directory"
        $loop = $true
        while ($loop) {
            if ($browse.ShowDialog() -eq "OK") {
                $loop = $false
            
            }
            else {
                return
            }
        }
        $browse.SelectedPath
        $browse.Dispose()
        $this.BoxPath.Text = $browse.SelectedPath
        return 
    }
    ActionButtonValidateQuery() {
        $this.ValidQuery = Invoke-QueryValidation -Query $this.BoxQuery.Text 
        if ($null -ne $this.ValidQuery) {
            $this.LabelRunStatus.Visible = $true
            $this.LabelRunStatus.ForeColor = "red"
            $this.LabelRunStatus.Text = $this.ValidQuery.'Error message'
            if (($this.ValidQuery.'Error Options').count -ne 0) {
                $this.ShowQueryOptions()
            }
            else {
                $this.BigGUI()
            }
        }
        else {
            $this.LabelRunStatus.Visible = $false
            $this.LabelRunStatus.ForeColor = "black"
            $this.BigGUI()
        }
    }
    ActionButtonRunNXQLQuery() {
        # Update Export status
        $this.LabelRunStatus.ForeColor = "orange"
        $this.LabelRunStatus.Text = "Proccessing..."
        $this.LabelRunStatus.Visible = $true
        # Disable Query Box for user to be unable to modify before reading it to variable
        $this.BoxQuery.enabled = $false
        $this.DisableButtons()
        [String]$this.Query = $this.BoxQuery.Text
        $this.BoxQuery.enabled = $true
        $FileName = $this.BoxFileName.Text
        # Check Platform
        $this.Platforms = @()
        if ($this.CheckboxWindows.checked) {
            $this.Platforms += "windows"
        }
        if ($this.CheckboxMac_OS.checked) {
            $this.Platforms += "mac_os"
        }
        if ($this.CheckboxMobile.checked) {
            $this.Platforms += "mobile"
        }
        # Handling if no platform is selected
        if ($null -eq $this.Platforms) {
            $this.LabelRunStatus.Visible = $true
            $this.LabelRunStatus.ForeColor = "red"
            $this.LabelRunStatus.Text = "There is no platform selected"
            Invoke-Buttons -Enable
            return
        }
        # Invoke Basic Query validation
        Invoke-QueryValidation -Query $this.Query
        if ($null -ne $this.ErrorInformation) {
            $this.LabelRunStatus.Visible = $true
            $this.LabelRunStatus.ForeColor = "red"
            $this.LabelRunStatus.Text = $this.ErrorInformation."Error message"
            Invoke-Buttons -Enable
            return
        }
        # Check if the file name does not contain any "/" or "\"
        if (($FileName.ToCharArray() | Where-Object { $_ -eq '/' } | Measure-Object).Count -gt 0 `
                -or `
            ($FileName.ToCharArray() | Where-Object { $_ -eq '\' } | Measure-Object).Count -gt 0 ) {
        
            $this.LabelFileName.ForeColor = "red"
            $this.LabelFileName.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 12, [System.Drawing.FontStyle]::Bold)
            Invoke-Buttons -Enable
            return
        }
        # If user set inccorret filename and on the next run it is correct remove red label
        $this.LabelFileName.ForeColor = "black"
        $this.LabelFileName.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 12)
        $WebAPIPort = $this.BoxPort.text
        $Path = $this.BoxPath.text
        $FileName = $this.BoxFileName.text
        if (!(Test-Path -Path $Path)) {
            New-Item -Path $Path -ItemType Directory
        }
        # Check if user added file extension
        if (($FileName -notlike "*.csv") -and ($FileName -notlike "*.txt")) {
            $FilePath = "$Path\$Filename.csv"
        }
        else {
            $FilePath = "$Path\$Filename"
        }
        if (!(Test-Path -Path $this.LogPath)) {
            New-Item -ItemType Directory -Path $this.LogPath
        }
        # Retrieve Nexthink data
        $result = Get-NxqlExport `
            -Query $this.Query `
            -credentials $this.Credentials `
            -webapiPort $WebAPIPort `
            -Platform $this.Platforms `
            -EngineList $this.Engines `
            -DestinationPath $FilePath `
            -SyncPath $this.LogPath `
            -LogPath $this.LogPath
        # Check if any data was returned
        if ($result -eq "Success!") {
            $this.LabelRunStatus.Visible = $true
            $this.LabelRunStatus.ForeColor = "green"
            $this.LabelRunStatus.Text = $result
            Invoke-Popup -title "NXQL Export" -description "NXQL Export for $this.CustomerName is ready!"
        }
        else {
            $this.LabelRunStatus.Visible = $true
            $this.LabelRunStatus.ForeColor = "red"
            $this.LabelRunStatus.Text = $result
            $result = $result.Split(":")
            Invoke-Popup -title "FAIL NXQL Export" -description "NXQL Export for $this.CustomerName failed with error: $result"
        }
        Invoke-Buttons -Enable
    }
    ActionButtonWebEditor() {
        # Select one of the engines
        $engine = ($this.Engines | Select-Object -First 1).address
        # Create a link to NXQL web editor
        $WebEditorAddress = "https://$engine/2/editor/nxql_editor.html"
        # Run the link
        Start-Process "$WebEditorAddress"
    }
    ActionBoxPortal() {
        if ($this.BoxPortal.text -like "https://*") {
            $FQDN = $this.BoxPortal.text
            $this.BoxPortal.text = $FQDN.Substring("https://".Length, ($FQDN.Length - ("https://".Length))).Split("/")[0]
        }
        Invoke-CredentialCleanup -Portal
    }
    ActionBoxQuery() {
        # Invoke Basic Query validation
        $this.ButtonRunQuery.visible = $false
        $this.ButtonValidateQuery.visible = $true
                
        $this.ValidQuery = Invoke-QueryValidation -Query $this.BoxQuery.Text -Ligth
        if (($null -ne $this.ValidQuery)) {
            $this.LabelRunStatus.Visible = $true
            $this.LabelRunStatus.ForeColor = "red"
            $this.LabelRunStatus.Text = $this.ValidQuery.'Error message'
        }
    }
    ActionCheckboxPlatform() {
        Invoke-ButtonValidateQuery
    }

}

Class EnvironmentForm {
    # Operations
    [String]$Environment
    # Form
    $EnvSelect
    # Label
    $LabelQuestion
    # Buttons
    $ButtonFITS
    $ButtonMoJo
    $ButtonAll
    
    EnvironmentForm() {
        $this.EnvSelect = New-Object system.Windows.Forms.Form
        $this.EnvSelect.ClientSize = New-Object System.Drawing.Point(390, 100)
        $this.EnvSelect.text = "Powershell NXQL API"
        $this.EnvSelect.TopMost = $true
        $this.EnvSelect.FormBorderStyle = 'FixedDialog'
        try {
            $p = (Get-Process powershell | Sort-Object -Property CPU -Descending | Select-Object -First 1).Path
            $this.EnvSelect.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($p)
        }
        catch {
            $p = (Get-Process explorer | Sort-Object -Property CPU -Descending | Select-Object -First 1).Path
            $this.EnvSelect.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($p)
        }
    
        ######################################################################
        #-------------------------- Labels Section --------------------------#
        ######################################################################
        
        $this.LabelQuestion = New-Object system.Windows.Forms.Label
        $this.LabelQuestion.text = "On which environment do you want to run query?"
        $this.LabelQuestion.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
        $this.LabelQuestion.AutoSize = $true
        $this.LabelQuestion.width = 370
        $this.LabelQuestion.height = 10
        $this.LabelQuestion.location = New-Object System.Drawing.Point(10, 10)
        $this.LabelQuestion.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 12)
        $this.EnvSelect.Controls.Add($this.LabelQuestion)
    
        ######################################################################
        #------------------------- Buttons Section --------------------------#
        ######################################################################
    
        $this.ButtonFITS = New-Object System.Windows.Forms.Button
        $this.ButtonFITS.Location = New-Object System.Drawing.Point(55, 50)
        $this.ButtonFITS.Size = New-Object System.Drawing.Size(80, 30)
        $this.ButtonFITS.Text = 'FITS EUCS'
        $this.ButtonFITS.Add_Click({
                $this.EnvSelect.Add_FormClosing({
                        $this.Environment = "FITS" })
                $this.EnvSelect.Close() })
        $this.EnvSelect.Controls.Add($this.ButtonFITS)
    
        $this.ButtonMoJo = New-Object System.Windows.Forms.Button
        $this.ButtonMoJo.Location = New-Object System.Drawing.Point(150, 50)
        $this.ButtonMoJo.Size = New-Object System.Drawing.Size(80, 30)
        $this.ButtonMoJo.Text = 'MoJo'
        $this.ButtonMoJo.Add_Click({
                $this.EnvSelect.Add_FormClosing({
                        $this.Environment = "MoJo" })
                $this.EnvSelect.Close() })
        $this.EnvSelect.Controls.Add($this.ButtonMoJo)
    
        $this.ButtonAll = New-Object System.Windows.Forms.Button
        $this.ButtonAll.Location = New-Object System.Drawing.Point(250, 50)
        $this.ButtonAll.Size = New-Object System.Drawing.Size(80, 30)
        $this.ButtonAll.Text = 'All'
        $this.ButtonAll.Add_Click({
                $this.EnvSelect.Add_FormClosing({
                        $this.Environment = "All" })
                $this.EnvSelect.Close() })
        $this.EnvSelect.Controls.Add($this.ButtonAll)
        [void]$this.EnvSelect.ShowDialog()
    }
}