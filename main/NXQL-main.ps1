Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$ErrorActionPreference = 'Stop'
try {
    New-Variable -Name 'Engines' -Value @() -Scope Script
    New-Variable -Name 'Credentials' -value "-" -Scope Script
    New-Variable -Name 'PortalFQDN' -value "-" -Scope Script
    New-Variable -Name 'Login' -Value "-" -Scope Script
    New-Variable -Name 'Environment' -Value "-" -Scope Script
    New-Variable -Name 'LogPath' -Value "$((Get-Location).Path)/Logs" -Scope Script
    New-Variable -Name 'Platform' -Value @() -Scope Script
}
catch {}


function Invoke-main {

    $Form = New-Object system.Windows.Forms.Form
    $Form.ClientSize = New-Object System.Drawing.Point(480, 150)
    $Form.text = "Powershell NXQL API"
    $Form.TopMost = $true
    # Handling if opened via Powershell ISE
    if ($null -ne (Get-Process powershell)) {
        $p = (Get-Process powershell | Sort-Object -Property CPU -Descending | Select-Object -First 1).Path
        $Form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($p)
    }

    ######################################################################
    #-------------------------- Labels Section --------------------------#
    ######################################################################

    $LabelPortal = New-Object system.Windows.Forms.Label
    $LabelPortal.text = "Portal FQDN: "
    $LabelPortal.AutoSize = $true
    $LabelPortal.width = 25
    $LabelPortal.height = 10
    $LabelPortal.location = New-Object System.Drawing.Point(20, 0)
    $LabelPortal.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 12)
    $Form.Controls.Add($LabelPortal)

    $LabelLogin = New-Object system.Windows.Forms.Label
    $LabelLogin.text = "Login:"
    $LabelLogin.AutoSize = $true
    $LabelLogin.width = 25
    $LabelLogin.height = 10
    $LabelLogin.location = New-Object System.Drawing.Point(20, 30)
    $LabelLogin.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 12)
    $Form.Controls.Add($LabelLogin)

    $LabelPassword = New-Object system.Windows.Forms.Label
    $LabelPassword.text = "Password:"
    $LabelPassword.AutoSize = $true
    $LabelPassword.width = 25
    $LabelPassword.height = 10
    $LabelPassword.location = New-Object System.Drawing.Point(20, 60)
    $LabelPassword.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 12)
    $Form.Controls.Add($LabelPassword)

    $LabelConnectionStatus = New-Object system.Windows.Forms.Label
    $LabelConnectionStatus.text = ""
    $LabelConnectionStatus.AutoSize = $true
    $LabelConnectionStatus.width = 25
    $LabelConnectionStatus.height = 10
    $LabelConnectionStatus.location = New-Object System.Drawing.Point(150, 105)
    $LabelConnectionStatus.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)
    $LabelConnectionStatus.Visible = $false
    $Form.Controls.Add($LabelConnectionStatus)

    $LabelConnectionStatusDetails = New-Object system.Windows.Forms.Label
    $LabelConnectionStatusDetails.text = ""
    $LabelConnectionStatusDetails.AutoSize = $true
    $LabelConnectionStatusDetails.width = 25
    $LabelConnectionStatusDetails.height = 10
    $LabelConnectionStatusDetails.location = New-Object System.Drawing.Point(260, 105)
    $LabelConnectionStatusDetails.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10, [System.Drawing.FontStyle]::Bold)
    $LabelConnectionStatusDetails.Visible = $true
    $Form.Controls.Add($LabelConnectionStatusDetails)

    $LabelNumberOfEngines = New-Object system.Windows.Forms.Label
    $LabelNumberOfEngines.text = "Number of engines: "
    $LabelNumberOfEngines.AutoSize = $true
    $LabelNumberOfEngines.width = 25
    $LabelNumberOfEngines.height = 10
    $LabelNumberOfEngines.location = New-Object System.Drawing.Point(350, 105)
    $LabelNumberOfEngines.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)
    $LabelNumberOfEngines.Visible = $false
    $Form.Controls.Add($LabelNumberOfEngines)

    $LabelQuery = New-Object system.Windows.Forms.Label
    $LabelQuery.text = "NXQL query:"
    $LabelQuery.AutoSize = $true
    $LabelQuery.width = 25
    $LabelQuery.height = 10
    $LabelQuery.location = New-Object System.Drawing.Point(20, 140)
    $LabelQuery.Visible = $false
    $LabelQuery.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 12)
    $Form.Controls.Add($LabelQuery)

    $LabelPort = New-Object system.Windows.Forms.Label
    $LabelPort.text = "NXQL port:"
    $LabelPort.AutoSize = $true
    $LabelPort.width = 25
    $LabelPort.height = 10
    $LabelPort.location = New-Object System.Drawing.Point(550, 140)
    $LabelPort.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 12)
    $LabelPort.Visible = $false
    $Form.Controls.Add($LabelPort)

    $LabelFileName = New-Object System.Windows.Forms.Label
    $LabelFileName.Text = "Destination File Name"
    $LabelFileName.AutoSize = $true
    $LabelFileName.width = 25
    $LabelFileName.height = 10
    $LabelFileName.location = New-Object System.Drawing.Point(20, 480)
    $LabelFileName.Visible = $false
    $LabelFileName.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 12)
    $Form.Controls.Add($LabelFileName)

    $LabelPath = New-Object System.Windows.Forms.Label
    $LabelPath.Text = "Destination Path"
    $LabelPath.AutoSize = $true
    $LabelPath.width = 25
    $LabelPath.height = 10
    $LabelPath.location = New-Object System.Drawing.Point(20, 510)
    $LabelPath.Visible = $false
    $LabelPath.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 12)
    $Form.Controls.Add($LabelPath)

    $LabelRunStatus = New-Object System.Windows.Forms.Label
    $LabelRunStatus.Text = ""
    $LabelRunStatus.AutoSize = $true
    $LabelRunStatus.width = 25
    $LabelRunStatus.height = 10
    $LabelRunStatus.location = New-Object System.Drawing.Point(150, 545)
    $LabelRunStatus.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10, [System.Drawing.FontStyle]::Bold)
    $LabelRunStatus.Visible = $false
    $Form.Controls.Add($LabelRunStatus)

    $LabelPlatform = New-Object System.Windows.Forms.Label
    $LabelPlatform.Text = "Select Platform:"
    $LabelPlatform.AutoSize = $true
    $LabelPlatform.width = 25
    $LabelPlatform.height = 10
    $LabelPlatform.location = New-Object System.Drawing.Point(550, 0)
    $LabelPlatform.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 12)
    $LabelPlatform.Visible = $false
    $Form.Controls.Add($LabelPlatform)

    ######################################################################
    #------------------------ Checkboxes Section ------------------------#
    ######################################################################

    $CheckboxWindows = New-Object System.Windows.Forms.Checkbox 
    $CheckboxWindows.Location = New-Object System.Drawing.Size(555, 25) 
    $CheckboxWindows.Size = New-Object System.Drawing.Size(100, 20)
    $CheckboxWindows.Text = "Windows"
    $CheckboxWindows.checked = $true
    $CheckboxWindows.Visible = $false
    $CheckboxWindows.TabIndex = 4
    $Form.Controls.Add($CheckboxWindows)

    $CheckboxMac_OS = New-Object System.Windows.Forms.Checkbox 
    $CheckboxMac_OS.Location = New-Object System.Drawing.Size(555, 45) 
    $CheckboxMac_OS.Size = New-Object System.Drawing.Size(100, 20)
    $CheckboxMac_OS.Text = "Mac OS"
    $CheckboxMac_OS.Visible = $false
    $CheckboxMac_OS.TabIndex = 4
    $Form.Controls.Add($CheckboxMac_OS)

    $CheckboxMobile = New-Object System.Windows.Forms.Checkbox 
    $CheckboxMobile.Location = New-Object System.Drawing.Size(555, 65) 
    $CheckboxMobile.Size = New-Object System.Drawing.Size(100, 20)
    $CheckboxMobile.Text = "Mobile"
    $CheckboxMobile.Visible = $false
    $CheckboxMobile.TabIndex = 4
    $Form.Controls.Add($CheckboxMobile)

    $CheckboxShowPassword = New-Object System.Windows.Forms.Checkbox 
    $CheckboxShowPassword.Location = New-Object System.Drawing.Size(140, 80) 
    $CheckboxShowPassword.Size = New-Object System.Drawing.Size(200, 20)
    $CheckboxShowPassword.Text = "Show Password"
    $CheckboxShowPassword.Visible = $true
    $CheckboxShowPassword.TabIndex = 4
    $CheckboxShowPassword.checked = $false
    $CheckboxShowPassword.Add_Click({
            if ($CheckboxShowPassword.checked) {
                $BoxPassword.passwordchar = $null
            }
            else {
                $BoxPassword.passwordchar = "*"
            }
        
        })
    $Form.Controls.Add($CheckboxShowPassword)

    ######################################################################
    #------------------------- TextBoxes Section ------------------------#
    ######################################################################

    $BoxPortal = New-Object System.Windows.Forms.TextBox 
    $BoxPortal.Multiline = $false
    $BoxPortal.Location = New-Object System.Drawing.Size(140, 0) 
    $BoxPortal.Size = New-Object System.Drawing.Size(300, 20)
    $BoxPortal.Add_TextChanged({ Invoke-CredentialCleanup })
    $Form.Controls.Add($BoxPortal)

    $BoxLogin = New-Object System.Windows.Forms.TextBox 
    $BoxLogin.Multiline = $false
    $BoxLogin.Location = New-Object System.Drawing.Size(140, 30) 
    $BoxLogin.Size = New-Object System.Drawing.Size(300, 20)
    $BoxLogin.Add_TextChanged({ Invoke-UsernameChange })
    $Form.Controls.Add($BoxLogin)

    $BoxPassword = New-Object System.Windows.Forms.MaskedTextBox 
    $BoxPassword.passwordchar = "*"
    $BoxPassword.Multiline = $false
    $BoxPassword.Location = New-Object System.Drawing.Size(140, 60) 
    $BoxPassword.Size = New-Object System.Drawing.Size(300, 20)
    $Form.Controls.Add($BoxPassword)

    $BoxQuery = New-Object System.Windows.Forms.TextBox 
    $BoxQuery.Multiline = $true
    $BoxQuery.Location = New-Object System.Drawing.Size(20, 165) 
    $BoxQuery.Size = New-Object System.Drawing.Size(660, 300)
    $BoxQuery.Scrollbars = 'Vertical'
    $BoxQuery.Visible = $false
    $BoxQuery.Add_TextChanged({
            # Invoke Basic Query validation
            $Status = Invoke-QueryValidation -Query $BoxQuery.Text -Ligth
            if ($LabelRunStatus.Text -eq "Proccessing...") {
                return
            }
            if (($null -ne $Status)) {
                $LabelRunStatus.Visible = $true
                $LabelRunStatus.ForeColor = "red"
                $LabelRunStatus.Text = $Status
            }
            else {
                $LabelRunStatus.Visible = $false
            }
        })
    $Form.Controls.Add($BoxQuery)

    $BoxFileName = New-Object System.Windows.Forms.TextBox 
    $BoxFileName.Multiline = $false
    $BoxFileName.Location = New-Object System.Drawing.Size(190, 480) 
    $BoxFileName.Size = New-Object System.Drawing.Size(495, 20)
    $BoxFileName.Text = (Get-Date).ToString("yyyy-MM-dd")
    $BoxFileName.Visible = $false
    $Form.Controls.Add($BoxFileName)

    $BoxPath = New-Object System.Windows.Forms.TextBox 
    $BoxPath.Multiline = $false
    $BoxPath.Location = New-Object System.Drawing.Size(150, 510) 
    $BoxPath.Size = New-Object System.Drawing.Size(430, 20)
    $BoxPath.Text = (Get-Location).Path
    $BoxPath.Visible = $false
    $Form.Controls.Add($BoxPath)

    $BoxPort = New-Object System.Windows.Forms.TextBox 
    $BoxPort.Multiline = $false
    $BoxPort.Location = New-Object System.Drawing.Size(640, 140) 
    $BoxPort.Size = New-Object System.Drawing.Size(40, 20)
    $BoxPort.Text = ""
    $BoxPort.Visible = $false
    $Form.Controls.Add($BoxPort)

    ######################################################################
    #------------------------- Buttons Section --------------------------#
    ######################################################################

    $ButtonConnect = New-Object System.Windows.Forms.Button
    $ButtonConnect.Location = New-Object System.Drawing.Point(20, 100)
    $ButtonConnect.Size = New-Object System.Drawing.Size(100, 30)
    $ButtonConnect.Text = 'Connect to portal'
    $ButtonConnect.Add_Click({ Invoke-PortalConnection })
    $Form.Controls.Add($ButtonConnect)

    $ButtonPath = New-Object System.Windows.Forms.Button
    $ButtonPath.Location = New-Object System.Drawing.Point(585, 505)
    $ButtonPath.Size = New-Object System.Drawing.Size(100, 30)
    $ButtonPath.Text = 'Select'
    $ButtonPath.Visible = $false
    $ButtonPath.Add_Click({ $BoxPath.Text = Get-Folder -inputFolder $BoxPath.Text })
    $Form.Controls.Add($ButtonPath)

    $ButtonWebEditor = New-Object System.Windows.Forms.Button
    $ButtonWebEditor.Location = New-Object System.Drawing.Point(585, 540)
    $ButtonWebEditor.Size = New-Object System.Drawing.Size(100, 50)
    $ButtonWebEditor.Visible = $false
    $ButtonWebEditor.Text = 'Open Web Query Editor'
    $ButtonWebEditor.Add_Click({ Invoke-WebQueryEditor })
    $Form.Controls.Add($ButtonWebEditor)

    $ButtonRunQuery = New-Object System.Windows.Forms.Button
    $ButtonRunQuery.Location = New-Object System.Drawing.Point(20, 540)
    $ButtonRunQuery.Size = New-Object System.Drawing.Size(120, 50)
    $ButtonRunQuery.Text = 'Run NXQL Query'
    $ButtonRunQuery.Visible = $false
    $ButtonRunQuery.Add_Click({ Invoke-NXQLQueryRun })
    $Form.Controls.Add($ButtonRunQuery)

    $Form.AcceptButton = $ButtonConnect
    [void]$Form.ShowDialog()
}

function Invoke-CredentialCleanup {
    Invoke-FormResize
    $ButtonRunQuery.Visible = $false
    $ButtonWebEditor.Visible = $false
    $LabelConnectionStatus.Visible = $false
    $LabelConnectionStatusDetails.Visible = $false
    $LabelNumberOfEngines.Visible = $false
    $LabelRunStatus.Visible = $false
    $BoxPassword.text = ""
    $BoxLogin.text = ""    
}
function Invoke-UsernameChange {
    Invoke-FormResize
    $ButtonRunQuery.Visible = $false
    $ButtonWebEditor.Visible = $false
    $LabelConnectionStatus.Visible = $false
    $LabelConnectionStatusDetails.Visible = $false
    $LabelNumberOfEngines.Visible = $false
    $LabelRunStatus.Visible = $false
    $BoxPassword.text = ""
    $LabelConnectionStatusDetails.Text = ""
}

function Invoke-PortalConnection {
    # Remember password while Button is clicked multiple times without changing anything
    $KeepCredentials = $false
    if ($LabelConnectionStatusDetails.Text -eq "Connected" -and 
        $script:PortalFQDN -eq $BoxPortal.Text -and
        $script:Login -eq $BoxLogin.Text) {
        $KeepCredentials = $true
    }
    else {
        $script:PortalFQDN = $BoxPortal.text
        $script:Login = $BoxLogin.Text
    }
    # Clear additional info to hide any information from previous run, which maybe misleading
    $LabelNumberOfEngines.Visible = $false
    $ButtonRunQuery.Visible = $false
    $LabelRunStatus.Visible = $false
    $LabelPort.Visible = $false
    $BoxPort.Visible = $false
    # Hide additional fields if button clicked multiple times
    Invoke-FormResize
    # Fill in the first part of output name and format portal connection details
    $BoxFileName.Text = (Get-Date).ToString("yyyy-MM-dd")
    $LabelConnectionStatus.Text = "Connection state:"
    $LabelConnectionStatus.ForeColor = "black"
    # Check if fields are not empty if yes exit
    if ($BoxLogin.Text.Length -lt 1 -or
        $BoxPassword.Text.Length -lt 1) {
        $BoxPassword.text = ""
        $LabelConnectionStatus.Text = "Login and password can not be empty !"
        $LabelConnectionStatus.ForeColor = "red"
        $LabelConnectionStatus.Visible = $true
        return
    }
    # Clear additional info to hide any information from previous run, which maybe misleading
    $LabelConnectionStatus.Visible = $true
    $LabelNumberOfEngines.Visible = $false
    $LabelNumberOfEngines.text = ""
    $LabelConnectionStatusDetails.text = "-"
    $LabelConnectionStatusDetails.ForeColor = "black"
    $Portal = $BoxPortal.Text
    # if login portal etc. are not changed and worked previously
    # use those credentials once agin
    if ($KeepCredentials -eq $false) {
        $Username = $BoxLogin.text
        $Password = $BoxPassword.text
        $Password = ConvertTo-SecureString $Password -AsPlainText -Force
        $script:Credentials = New-Object System.Management.Automation.PSCredential ($Username, $Password)
    }
    $BoxPassword.text = "********************"
    $script:Engines = Get-EngineList -portal $Portal -credentials $script:Credentials
    # Check if Engine list is not null if yes exit
    if ($null -eq $Engines) {
        $LabelConnectionStatusDetails.text = "Not connected"
        $LabelConnectionStatusDetails.ForeColor = "red"
        $LabelConnectionStatusDetails.Visible = $true
        $BoxPassword.text = ""
        return
    }
    # Handling for the environments with only one engine
    if ($null -eq $Engines.count) {
        $Number_of_engines = 1
    }
    else {
        $Number_of_engines = $Engines.Count
    }
    # Set state and number of Engines
    $LabelConnectionStatusDetails.Visible = $true
    $LabelConnectionStatusDetails.text = "Connected"
    $LabelConnectionStatusDetails.ForeColor = "green"
    $LabelNumberOfEngines.text = "Number of engines: $Number_of_engines"
    # Display additional fields
    Invoke-FormResize -Big
    $LabelNumberOfEngines.Visible = $true
    $ButtonRunQuery.Visible = $true
    # Check environment type SAAS / On-prem
    # Set additional details based on it
    if ($BoxPortal.Text -notlike "*.nexthink.cloud" ) {
        $BoxPort.Text = "1671"
        $BoxFileName.Text += " - <Customer_Name> - "
    }
    else {
        $CustomerName = $BoxPortal.Text.Split(".")[0]
        $BoxPort.Text = "443"
        $BoxFileName.Text += " - $CustomerName - "
    }
    # Display additional components of the GUI
    $ButtonWebEditor.Visible = $true
    $LabelPort.Visible = $true
    $BoxPort.Visible = $true
}
Function Get-Folder {
    param(
        $inputFolder
    )
    [Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $browse = New-Object System.Windows.Forms.FolderBrowserDialog
    $browse.SelectedPath = $inputFolder
    $browse.ShowNewFolderButton = $true
    $browse.Description = "Select a directory"
    $loop = $true
    while ($loop) {
        if ($browse.ShowDialog() -eq "OK") {
            $loop = $false
		
        }
        else {
            return $inputFolder
        }
    }
    $browse.SelectedPath
    $browse.Dispose()
    return $FolderBrowserDialog.SelectedPath
}
function Invoke-NXQLQueryRun {
    # Update Export status
    $LabelRunStatus.ForeColor = "orange"
    $LabelRunStatus.Text = "Proccessing..."
    $LabelRunStatus.Visible = $true
    # Disable Query Box for user to be unable to modify before reading it to variable
    $BoxQuery.enabled = $false
    Invoke-Buttons
    [String]$Query = $BoxQuery.Text
    $BoxQuery.enabled = $true
    $FileName = $BoxFileName.Text
    # Check Platform
    $Platform = @()
    if ($CheckboxWindows.checked) {
        $Platform += "windows"
    }
    if ($CheckboxMac_OS.checked) {
        $Platform += "mac_os"
    }
    if ($CheckboxMobile.checked) {
        $Platform += "mobile"
    }
    # Handling if no platform is selected
    if ($null -eq $Platform) {
        $LabelRunStatus.Visible = $true
        $LabelRunStatus.ForeColor = "red"
        $LabelRunStatus.Text = "There is no platform selected"
        Invoke-Buttons -Enable
        return
    }
    # Invoke Basic Query validation
    $Status = Invoke-QueryValidation -Query $Query
    if ($null -ne $Status) {
        $LabelRunStatus.Visible = $true
        $LabelRunStatus.ForeColor = "red"
        $LabelRunStatus.Text = $Status
        Invoke-Buttons -Enable
        return
    }
    # Check if the file name does not contain any "/" or "\"
    if (($FileName.ToCharArray() | Where-Object { $_ -eq '/' } | Measure-Object).Count -gt 0 `
            -or `
        ($FileName.ToCharArray() | Where-Object { $_ -eq '\' } | Measure-Object).Count -gt 0 ) {
        
        $LabelFileName.ForeColor = "red"
        $LabelFileName.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 12, [System.Drawing.FontStyle]::Bold)
        Invoke-Buttons -Enable
        return
    }
    # If user set inccorret filename and on the next run it is correct remove red label
    $LabelFileName.ForeColor = "black"
    $LabelFileName.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 12)
    $WebAPIPort = $BoxPort.text
    $Path = $BoxPath.text
    $FileName = $BoxFileName.text
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
    if (!(Test-Path -Path $script:LogPath)) {
        New-Item -ItemType Directory -Path $script:LogPath
    }
    # Retrieve Nexthink data
    $result = Get-NxqlExport `
        -Query $Query `
        -credentials $script:Credentials `
        -webapiPort $WebAPIPort `
        -Platform $Platform `
        -EngineList $script:Engines `
        -DestinationPath $FilePath `
        -SyncPath $script:LogPath `
        -LogPath $script:LogPath
    # Check if any data was returned
    if ($result -eq "Success!") {
        $LabelRunStatus.Visible = $true
        $LabelRunStatus.ForeColor = "green"
        $LabelRunStatus.Text = $result
    }
    else {
        $LabelRunStatus.Visible = $true
        $LabelRunStatus.ForeColor = "red"
        $LabelRunStatus.Text = $result
    }
    Invoke-Buttons -Enable
}
function Invoke-QueryValidation {
    param (
        [String]$Query,
        [Switch]$Ligth
    )
    # Check if all opened brackets in query are closed
    if (($Query.ToCharArray() | Where-Object { $_ -eq '(' } | Measure-Object).Count `
            -ne `
        ($Query.ToCharArray() | Where-Object { $_ -eq ')' } | Measure-Object).Count) {
        if ($Ligth) {
            return "Some brackets are not closed !"
        }
        else {
            return "Failed: Some brackets are not closed !"
        }
    }
    if (($Query.ToCharArray() | Where-Object { $_ -eq '`"' } | Measure-Object).Count `
            -ne `
        ($Query.ToCharArray() | Where-Object { $_ -eq '`"' } | Measure-Object).Count) {
        if ($Ligth) {
            return "Some quotes are not closed !"
        }
        else {
            return "Failed: Some quotes are not closed !"
        }
    }
    if (($Query.ToCharArray() | Where-Object { $_ -eq "`'" } | Measure-Object).Count `
            -ne `
        ($Query.ToCharArray() | Where-Object { $_ -eq "`'" } | Measure-Object).Count) {
        if ($Ligth) {
            return "Some quotes are not closed !"
        }
        else {
            return "Failed: Some quotes are not closed !"
        }
    }
    if ($Ligth) {
        return $null
    }
    # Check if query is not empty
    if ($Query.Length -le 1) {
        return "Failed: NXQL query can not be blank !"
    }
    # Check if select statement exists
    if (($Query -notlike "*select*")) {
        return "Failed: There is no `"select`" statement !"
    }
    # Check if from statement exists
    if (($Query -notlike "*from*")) {
        return "Failed: There is no `"from`" statement !"
    }
    # Check if limit statement exists
    if (($Query -notlike "*limit*")) {
        return "Failed: There is no `"limit`" statement !"
    }
    return $null
}
function Invoke-WebQueryEditor {
    # Select one of the engines
    $engine = ($script:Engines | Select-Object -First 1).address
    # Create a link to NXQL web editor
    $WebEditorAddress = "https://$engine/2/editor/nxql_editor.html"
    # Run the link
    Start-Process "$WebEditorAddress"
}
function Get-EngineList {
    <#
.SYNOPSIS
Returns list of Engines connected to Nexthink Portal

.DESCRIPTION
Connets to Nexthink Portal and retrieves list of all engines,
next select only connected ones.

.PARAMETER portal
The Nexthink Portal DNS Name to retrieve connected engines

.PARAMETER credentials
Nexthink account authorised to extract list of engines

.EXAMPLE
Get-EngineList -portal "test.eu.nexthink.cloud" -credentials <Account_UserName>

.INPUTS
String

.OUTPUTS
Hastable

.NOTES
    Author:  Stanislaw Horna
#>
    param (
        [Parameter(Mandatory = $true)]
        [string] $portal,
        [Parameter(Mandatory = $true)]
        [System.Management.Automation.PSCredential]$credentials
    )
    $web = [Net.WebClient]::new()
    $web.Credentials = $credentials
    $pair = [string]::Join(":", $web.Credentials.UserName, $web.Credentials.Password)
    $base64 = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($pair))
    $web.Headers.Add('Authorization', "Basic $base64")
    $baseUrl = "https://$portal/api/configuration/v1/engines"
    $result = $web.downloadString($baseUrl)
    $engineList = $result | ConvertFrom-Json
    $engineList = $engineList | Where-Object { $_.status -eq "CONNECTED" }
    # Listing Connected Engines only
    if ($portal -eq 'ministryofjustice.eu.nexthink.cloud') {
        $FITS = ('engine-1', 'engine-2', 'engine-3', 'engine-4', 'engine-5', 'engine-6', 'engine-7', 'engine-8', 'engine-9')
        $MOJO = ('engine-10', 'engine-11', 'engine-12', 'engine-13', 'engine-14')
        Invoke-EnvironmentSelection
        if ($script:Environment -eq "FITS") {
            $engineList = $engineList | Where-Object { $_.name -in $FITS }
        }
        elseif ($script:Environment -eq "MoJo") {
            $engineList = $engineList | Where-Object { $_.name -in $MOJO }
        }
    }
    return $engineList
}
function Invoke-EnvironmentSelection {

    $EnvSelect = New-Object system.Windows.Forms.Form
    $EnvSelect.ClientSize = New-Object System.Drawing.Point(390, 100)
    $EnvSelect.text = "Powershell NXQL API"
    $EnvSelect.TopMost = $true
    if ($null -ne (Get-Process powershell)) {
        $p = (Get-Process powershell | Sort-Object -Property CPU -Descending | Select-Object -First 1).Path
    }
    if ($null -ne $p) {
        $EnvSelect.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($p)
    }

    $LabelQuestion = New-Object system.Windows.Forms.Label
    $LabelQuestion.text = "On which environment do you want to run query?"
    $LabelQuestion.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $LabelQuestion.AutoSize = $true
    $LabelQuestion.width = 370
    $LabelQuestion.height = 10
    $LabelQuestion.location = New-Object System.Drawing.Point(10, 10)
    $LabelQuestion.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 12)
    $EnvSelect.Controls.Add($LabelQuestion)

    $ButtonFITS = New-Object System.Windows.Forms.Button
    $ButtonFITS.Location = New-Object System.Drawing.Point(55, 50)
    $ButtonFITS.Size = New-Object System.Drawing.Size(80, 30)
    $ButtonFITS.Text = 'FITS EUCS'
    $ButtonFITS.Add_Click({
            $EnvSelect.Add_FormClosing({
                    $script:Environment = "FITS" })
            $EnvSelect.Close() })
    $EnvSelect.Controls.Add($ButtonFITS)

    $ButtonMoJo = New-Object System.Windows.Forms.Button
    $ButtonMoJo.Location = New-Object System.Drawing.Point(150, 50)
    $ButtonMoJo.Size = New-Object System.Drawing.Size(80, 30)
    $ButtonMoJo.Text = 'MoJo'
    $ButtonMoJo.Add_Click({
            $EnvSelect.Add_FormClosing({
                    $script:Environment = "MoJo" })
            $EnvSelect.Close() })
    $EnvSelect.Controls.Add($ButtonMoJo)

    $ButtonAll = New-Object System.Windows.Forms.Button
    $ButtonAll.Location = New-Object System.Drawing.Point(250, 50)
    $ButtonAll.Size = New-Object System.Drawing.Size(80, 30)
    $ButtonAll.Text = 'All'
    $ButtonAll.Add_Click({
            $EnvSelect.Add_FormClosing({
                    $script:Environment = "All" })
            $EnvSelect.Close() })
    $EnvSelect.Controls.Add($ButtonAll)
    [void]$EnvSelect.ShowDialog()
}
Function Get-NxqlExport {
    param (
        [Parameter(Mandatory = $true)]
        [String] $Query,
        [Parameter(Mandatory = $true)]
        [System.Management.Automation.PSCredential] $credentials,
        [Parameter(Mandatory = $true)]
        $EngineList,
        [Parameter(Mandatory = $false)]
        [String]$webapiPort,
        [Parameter(Mandatory = $false)]
        [String[]]$Platform,
        [Parameter(Mandatory = $false)]
        [string]$DestinationPath,
        [Parameter(Mandatory = $false)]
        [string]$SyncPath,
        [Parameter(Mandatory = $false)]
        [string]$LogPath
    )
    $Functions = {
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
            catch {
                throw
            }
        };
    }
    if (! $webapiPort) {
        $webapiPort = "443"
    }
    if (! $credentials) {
        $credentials = Get-Credential
    }
    if (!$EngineList) {
        $portal = Read-Host "Enter Portal DNS name"
        $EngineList = Get-EngineList -portal $portal -credentials $credentials
    }

    if (Test-Path -Path "$SyncPath\Headers") {
        Remove-Item -Path "$SyncPath\Headers" -Confirm:$false -Force
    }
    if (Test-Path -Path "$SyncPath\Wait") {
        Remove-Item -Path "$SyncPath\Wait" -Confirm:$false -Force
    }
    if (Test-Path -Path $LogPath) {
        $LogsToDelete = (Get-ChildItem -Path $LogPath).FullName
        foreach ($file in $LogsToDelete) {
            Remove-Item -Path $file -Confirm:$false -Force
        }
    }
    # Create separate process for each engine in scope
    foreach ($Engine in $EngineList) {
        $Name = "NXQL-" + $Engine.name
        $RandomWaitTime = Get-Random -Minimum 0 -Maximum 500
        Start-Job -Name $Name `
            -InitializationScript $Functions `
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
                $_.Exception | Out-File "$LogPath\Log-$EngineAddress.csv" -Append
                "Probably bad query" | Out-File "$LogPath\BadRequest"
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
        } -ArgumentList $Engine.address, $webapiPort, $Platform, $credentials, $Query, $DestinationPath, $SyncPath, $RandomWaitTime, $LogPath | Out-Null
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
    if (Test-Path -Path "$SyncPath\Headers") {
        Remove-Item -Path "$SyncPath\Headers" -Confirm:$false -Force
    }
    if (Test-Path -Path "$SyncPath\Wait") {
        Remove-Item -Path "$SyncPath\Wait" -Confirm:$false -Force
    }
    # Handling for environments with only one engine
    if ($null -eq $Engines.count) {
        $Number_of_engines = 1
    }
    else {
        $Number_of_engines = $Engines.Count
    }
    # Check if outputs from all engines are pasted to the result file
    if ($CompletedJobsCounter -eq $Number_of_engines) {
        return "Success!"
    }
    elseif (Test-Path "$LogPath\BadRequest") {
        return "Failed: Invalid NXQL query"
    }
    else {
        return "Failed: Error unknown"
    }
}
function Invoke-FormResize {
    # Function to change window mode
    param (
        [switch]$Big
    )
    if ($Big) {
        $Form.TopMost = $false
        $Form.ClientSize = New-Object System.Drawing.Point(700, 600)
        $LabelQuery.Visible = $true
        $BoxQuery.Visible = $true
        $LabelFileName.Visible = $true
        $BoxFileName.Visible = $true
        $LabelPath.Visible = $true
        $BoxPath.Visible = $true
        $ButtonPath.Visible = $true
        $LabelPlatform.Visible = $true
        $CheckboxWindows.Visible = $true
        $CheckboxMac_OS.Visible = $true
        $CheckboxMobile.Visible = $true
    }
    else {
        $Form.TopMost = $true
        $Form.ClientSize = New-Object System.Drawing.Point(480, 150)
        $LabelQuery.Visible = $false
        $BoxQuery.Visible = $false
        $LabelFileName.Visible = $false
        $BoxFileName.Visible = $false
        $LabelPath.Visible = $false
        $BoxPath.Visible = $false
        $ButtonPath.Visible = $false
        $LabelPlatform.Visible = $false
        $CheckboxWindows.Visible = $false
        $CheckboxMac_OS.Visible = $false
        $CheckboxMobile.Visible = $false
        $ButtonWebEditor.Visible = $false
    }
}
function Invoke-Buttons {
    # Function to disable and enable action buttons
    param (
        [switch]$Enable
    )
    if ($Enable) {
        $ButtonConnect.enabled = $true
        $ButtonRunQuery.enabled = $true
        $ButtonPath.enabled = $true
        $BoxPath.enabled = $true
        $BoxFileName.enabled = $true
        $BoxPort.enabled = $true
        $CheckboxWindows.enabled = $true
        $CheckboxMac_OS.enabled = $true
        $CheckboxMobile.enabled = $true
    }
    else {
        $ButtonConnect.enabled = $false
        $ButtonRunQuery.enabled = $false
        $ButtonPath.enabled = $false
        $BoxPath.enabled = $false
        $BoxFileName.enabled = $false
        $BoxPort.enabled = $false
        $CheckboxWindows.enabled = $false
        $CheckboxMac_OS.enabled = $false
        $CheckboxMobile.enabled = $false
    }
}


Invoke-main

#commit to push