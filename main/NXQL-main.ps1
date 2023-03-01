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
    Version:            1.04
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

    2023-02-28      Stanislaw Horna         Error Handling for invalid query - 
                                            returns the same message as NXQL WebEditor.
                                            
#>

function Invoke-main {
    Invoke-GettingStarted
    Invoke-FormMain
    $script:Form.AcceptButton = $script:ButtonConnect
    [void]$script:Form.ShowDialog()
}
function Invoke-GettingStarted {
    $ErrorActionPreference = 'Stop'
    # Loading additional .NET classes
    try {
        Add-Type -AssemblyName PresentationCore, PresentationFramework
    }
    catch {
        Write-Host "Currently running environment is not supported"
        Pause
        throw
    }

    try {
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Web
        [System.Windows.Forms.Application]::EnableVisualStyles()
    }
    catch {
        Invoke-FormError -Message "Currently running environment is not supported"
    }

    # Creating neccessary variables
    try {
        New-Variable -Name 'Engines' -Value @() -Scope Script
        New-Variable -Name 'Credentials' -value "-" -Scope Script
        New-Variable -Name 'PortalFQDN' -value "-" -Scope Script
        New-Variable -Name 'Login' -Value "-" -Scope Script
        New-Variable -Name 'Environment' -Value "-" -Scope Script
        New-Variable -Name 'Platform' -Value @() -Scope Script
        New-Variable -Name 'ErrorInformation' -Value @{} -Scope Script
    }
    catch {}

    # Handling if someone open directly NXQL-main.ps1
    if ((Get-Location).Path -like "*\main") {
        try {
            New-Variable -Name 'LogPath' -Value "$((Get-Location).Path)\Logs" -Scope Script
        }
        catch {}
    }
    else {
        try {
            New-Variable -Name 'LogPath' -Value "$((Get-Location).Path)\main\Logs" -Scope Script
        }
        catch {}
    }

    try {
        New-Variable -Name 'UniqueInstanceLock' -Value "$LogPath\UniqueInstance" -Scope Script
    }
    catch {}
    
    # Create Log directory
    if (!(Test-Path -Path $script:LogPath)) {
        try {
            New-Item -Path $script:LogPath -ItemType Directory | Out-Null
        }
        catch {
            Invoke-FormError -Message "No write permission in the main catalog."
        }
    }

    # Creating lock file to prevent running multiple instances from the same location
    if (!(Test-Path -Path $script:UniqueInstanceLock)) {
        try {
            New-Item -Path $script:UniqueInstanceLock | Out-Null
        }
        catch {
            Invoke-FormError -Message "No write permission in the main catalog."
        }
        
    }
    else {
        Invoke-FormError -Message "Another instance of this application is already running."
    }
}
######################################################################
#----------------------- GUI Forms Definition -----------------------#
######################################################################
function Invoke-FormMain {
    $script:Form = New-Object system.Windows.Forms.Form
    $script:Form.ClientSize = New-Object System.Drawing.Point(480, 150)
    $script:Form.text = "Powershell NXQL API"
    $script:Form.FormBorderStyle = 'FixedDialog'
    $script:Form.TopMost = $true
    # Handling if opened via Powershell ISE
    try {
        $p = (Get-Process powershell | Sort-Object -Property CPU -Descending | Select-Object -First 1).Path
        $script:Form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($p)
    }
    catch {
        $p = (Get-Process explorer | Sort-Object -Property CPU -Descending | Select-Object -First 1).Path
        $script:Form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($p)
    }
    $script:Form.Add_FormClosing({
            Remove-Item -Path $script:UniqueInstanceLock -Confirm:$false -Force
        })

    ######################################################################
    #-------------------------- Labels Section --------------------------#
    ######################################################################

    $script:LabelPortal = New-Object system.Windows.Forms.Label
    $script:LabelPortal.text = "Portal FQDN: "
    $script:LabelPortal.AutoSize = $true
    $script:LabelPortal.width = 25
    $script:LabelPortal.height = 10
    $script:LabelPortal.location = New-Object System.Drawing.Point(20, 0)
    $script:LabelPortal.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 12)
    $script:Form.Controls.Add($script:LabelPortal)

    $script:LabelLogin = New-Object system.Windows.Forms.Label
    $script:LabelLogin.text = "Login:"
    $script:LabelLogin.AutoSize = $true
    $script:LabelLogin.width = 25
    $script:LabelLogin.height = 10
    $script:LabelLogin.location = New-Object System.Drawing.Point(20, 30)
    $script:LabelLogin.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 12)
    $script:Form.Controls.Add($script:LabelLogin)

    $script:LabelPassword = New-Object system.Windows.Forms.Label
    $script:LabelPassword.text = "Password:"
    $script:LabelPassword.AutoSize = $true
    $script:LabelPassword.width = 25
    $script:LabelPassword.height = 10
    $script:LabelPassword.location = New-Object System.Drawing.Point(20, 60)
    $script:LabelPassword.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 12)
    $script:Form.Controls.Add($script:LabelPassword)

    $script:LabelConnectionStatus = New-Object system.Windows.Forms.Label
    $script:LabelConnectionStatus.text = ""
    $script:LabelConnectionStatus.AutoSize = $true
    $script:LabelConnectionStatus.width = 25
    $script:LabelConnectionStatus.height = 10
    $script:LabelConnectionStatus.location = New-Object System.Drawing.Point(150, 105)
    $script:LabelConnectionStatus.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)
    $script:LabelConnectionStatus.Visible = $false
    $script:Form.Controls.Add($script:LabelConnectionStatus)

    $script:LabelConnectionStatusDetails = New-Object system.Windows.Forms.Label
    $script:LabelConnectionStatusDetails.text = ""
    $script:LabelConnectionStatusDetails.AutoSize = $true
    $script:LabelConnectionStatusDetails.width = 25
    $script:LabelConnectionStatusDetails.height = 10
    $script:LabelConnectionStatusDetails.location = New-Object System.Drawing.Point(260, 105)
    $script:LabelConnectionStatusDetails.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10, [System.Drawing.FontStyle]::Bold)
    $script:LabelConnectionStatusDetails.Visible = $true
    $script:Form.Controls.Add($script:LabelConnectionStatusDetails)

    $script:LabelNumberOfEngines = New-Object system.Windows.Forms.Label
    $script:LabelNumberOfEngines.text = "Number of engines: "
    $script:LabelNumberOfEngines.AutoSize = $true
    $script:LabelNumberOfEngines.width = 25
    $script:LabelNumberOfEngines.height = 10
    $script:LabelNumberOfEngines.location = New-Object System.Drawing.Point(350, 105)
    $script:LabelNumberOfEngines.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)
    $script:LabelNumberOfEngines.Visible = $false
    $script:Form.Controls.Add($script:LabelNumberOfEngines)

    $script:LabelQuery = New-Object system.Windows.Forms.Label
    $script:LabelQuery.text = "NXQL query:"
    $script:LabelQuery.AutoSize = $true
    $script:LabelQuery.width = 25
    $script:LabelQuery.height = 10
    $script:LabelQuery.location = New-Object System.Drawing.Point(20, 140)
    $script:LabelQuery.Visible = $false
    $script:LabelQuery.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 12)
    $script:Form.Controls.Add($script:LabelQuery)

    $script:LabelPort = New-Object system.Windows.Forms.Label
    $script:LabelPort.text = "NXQL port:"
    $script:LabelPort.AutoSize = $true
    $script:LabelPort.width = 25
    $script:LabelPort.height = 10
    $script:LabelPort.location = New-Object System.Drawing.Point(550, 140)
    $script:LabelPort.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 12)
    $script:LabelPort.Visible = $false
    $script:Form.Controls.Add($script:LabelPort)

    $script:LabelFileName = New-Object System.Windows.Forms.Label
    $script:LabelFileName.Text = "Destination File Name"
    $script:LabelFileName.AutoSize = $true
    $script:LabelFileName.width = 25
    $script:LabelFileName.height = 10
    $script:LabelFileName.location = New-Object System.Drawing.Point(20, 480)
    $script:LabelFileName.Visible = $false
    $script:LabelFileName.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 12)
    $script:Form.Controls.Add($script:LabelFileName)

    $script:LabelPath = New-Object System.Windows.Forms.Label
    $script:LabelPath.Text = "Destination Path"
    $script:LabelPath.AutoSize = $true
    $script:LabelPath.width = 25
    $script:LabelPath.height = 10
    $script:LabelPath.location = New-Object System.Drawing.Point(20, 510)
    $script:LabelPath.Visible = $false
    $script:LabelPath.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 12)
    $script:Form.Controls.Add($script:LabelPath)

    $script:LabelRunStatus = New-Object System.Windows.Forms.Label
    $script:LabelRunStatus.Text = ""
    $script:LabelRunStatus.AutoSize = $true
    $script:LabelRunStatus.TextAlign = "MiddleCenter"
    $script:LabelRunStatus.MaximumSize = New-Object System.Drawing.Size(435, 60)
    $script:LabelRunStatus.width = 25
    $script:LabelRunStatus.height = 10
    $script:LabelRunStatus.location = New-Object System.Drawing.Point(150, 540)
    $script:LabelRunStatus.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10, [System.Drawing.FontStyle]::Bold)
    $script:LabelRunStatus.Visible = $false
    $script:Form.Controls.Add($script:LabelRunStatus)

    $script:LabelPlatform = New-Object System.Windows.Forms.Label
    $script:LabelPlatform.Text = "Select Platform:"
    $script:LabelPlatform.AutoSize = $true
    $script:LabelPlatform.width = 25
    $script:LabelPlatform.height = 10
    $script:LabelPlatform.location = New-Object System.Drawing.Point(550, 0)
    $script:LabelPlatform.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 12)
    $script:LabelPlatform.Visible = $false
    $script:Form.Controls.Add($script:LabelPlatform)

    $script:LabelLookup = New-Object System.Windows.Forms.Label
    $script:LabelLookup.Text = "Look for:"
    $script:LabelLookup.AutoSize = $true
    $script:LabelLookup.width = 25
    $script:LabelLookup.height = 10
    $script:LabelLookup.location = New-Object System.Drawing.Point(700, 0)
    $script:LabelLookup.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 12)
    $script:LabelLookup.Visible = $false
    $script:Form.Controls.Add($script:LabelLookup)

    ######################################################################
    #------------------------ Checkboxes Section ------------------------#
    ######################################################################

    $script:CheckboxWindows = New-Object System.Windows.Forms.Checkbox 
    $script:CheckboxWindows.Location = New-Object System.Drawing.Size(555, 25) 
    $script:CheckboxWindows.Size = New-Object System.Drawing.Size(100, 20)
    $script:CheckboxWindows.Text = "Windows"
    $script:CheckboxWindows.checked = $true
    $script:CheckboxWindows.Visible = $false
    $script:CheckboxWindows.TabIndex = 4
    $script:CheckboxWindows.Add_Click({ Invoke-CheckboxPlatform })
    $script:Form.Controls.Add($script:CheckboxWindows)

    $script:CheckboxMac_OS = New-Object System.Windows.Forms.Checkbox 
    $script:CheckboxMac_OS.Location = New-Object System.Drawing.Size(555, 45) 
    $script:CheckboxMac_OS.Size = New-Object System.Drawing.Size(100, 20)
    $script:CheckboxMac_OS.Text = "Mac OS"
    $script:CheckboxMac_OS.Visible = $false
    $script:CheckboxMac_OS.TabIndex = 4
    $script:CheckboxMac_OS.Add_Click({ Invoke-CheckboxPlatform })
    $script:Form.Controls.Add($script:CheckboxMac_OS)

    $script:CheckboxMobile = New-Object System.Windows.Forms.Checkbox 
    $script:CheckboxMobile.Location = New-Object System.Drawing.Size(555, 65) 
    $script:CheckboxMobile.Size = New-Object System.Drawing.Size(100, 20)
    $script:CheckboxMobile.Text = "Mobile"
    $script:CheckboxMobile.Visible = $false
    $script:CheckboxMobile.TabIndex = 4
    $script:CheckboxMobile.Add_Click({ Invoke-CheckboxPlatform })
    $script:Form.Controls.Add($script:CheckboxMobile)

    $script:CheckboxShowPassword = New-Object System.Windows.Forms.Checkbox 
    $script:CheckboxShowPassword.Location = New-Object System.Drawing.Size(140, 80) 
    $script:CheckboxShowPassword.Size = New-Object System.Drawing.Size(200, 20)
    $script:CheckboxShowPassword.Text = "Show Password"
    $script:CheckboxShowPassword.Visible = $true
    $script:CheckboxShowPassword.TabIndex = 4
    $script:CheckboxShowPassword.checked = $false
    $script:CheckboxShowPassword.Add_Click({ Invoke-CheckboxShowPassword })
    $script:Form.Controls.Add($script:CheckboxShowPassword)

    ######################################################################
    #------------------------- TextBoxes Section ------------------------#
    ######################################################################

    $script:BoxPortal = New-Object System.Windows.Forms.TextBox 
    $script:BoxPortal.Multiline = $false
    $script:BoxPortal.Location = New-Object System.Drawing.Size(140, 0) 
    $script:BoxPortal.Size = New-Object System.Drawing.Size(300, 20)
    $script:BoxPortal.Add_TextChanged({ Invoke-CredentialCleanup -Portal })
    $script:Form.Controls.Add($script:BoxPortal)

    $script:BoxLogin = New-Object System.Windows.Forms.TextBox 
    $script:BoxLogin.Multiline = $false
    $script:BoxLogin.Location = New-Object System.Drawing.Size(140, 30) 
    $script:BoxLogin.Size = New-Object System.Drawing.Size(300, 20)
    $script:BoxLogin.Add_TextChanged({ Invoke-CredentialCleanup })
    $script:Form.Controls.Add($script:BoxLogin)

    $script:BoxPassword = New-Object System.Windows.Forms.MaskedTextBox 
    $script:BoxPassword.passwordchar = "*"
    $script:BoxPassword.Multiline = $false
    $script:BoxPassword.Location = New-Object System.Drawing.Size(140, 60) 
    $script:BoxPassword.Size = New-Object System.Drawing.Size(300, 20)
    $script:Form.Controls.Add($script:BoxPassword)

    $script:BoxQuery = New-Object System.Windows.Forms.TextBox 
    $script:BoxQuery.Multiline = $true
    $script:BoxQuery.Location = New-Object System.Drawing.Size(20, 165) 
    $script:BoxQuery.Size = New-Object System.Drawing.Size(660, 300)
    $script:BoxQuery.Scrollbars = 'Vertical'
    $script:BoxQuery.Visible = $false
    $script:BoxQuery.Add_TextChanged({ Invoke-BoxQuery })
    $script:Form.Controls.Add($script:BoxQuery)

    $script:BoxFileName = New-Object System.Windows.Forms.TextBox 
    $script:BoxFileName.Multiline = $false
    $script:BoxFileName.Location = New-Object System.Drawing.Size(190, 480) 
    $script:BoxFileName.Size = New-Object System.Drawing.Size(495, 20)
    $script:BoxFileName.Text = (Get-Date).ToString("yyyy-MM-dd")
    $script:BoxFileName.Visible = $false
    $script:Form.Controls.Add($script:BoxFileName)

    $script:BoxPath = New-Object System.Windows.Forms.TextBox 
    $script:BoxPath.Multiline = $false
    $script:BoxPath.Location = New-Object System.Drawing.Size(150, 510) 
    $script:BoxPath.Size = New-Object System.Drawing.Size(430, 20)
    $script:BoxPath.Text = Get-BoxPathLocation
    $script:BoxPath.Visible = $false
    $script:Form.Controls.Add($script:BoxPath)

    $script:BoxPort = New-Object System.Windows.Forms.TextBox 
    $script:BoxPort.Multiline = $false
    $script:BoxPort.Location = New-Object System.Drawing.Size(640, 140) 
    $script:BoxPort.Size = New-Object System.Drawing.Size(40, 20)
    $script:BoxPort.Text = ""
    $script:BoxPort.Visible = $false
    $script:Form.Controls.Add($script:BoxPort)

    $script:BoxLookfor = New-Object System.Windows.Forms.TextBox 
    $script:BoxLookfor.Text = ""
    $script:BoxLookfor.AutoSize = $true
    $script:BoxLookfor.width = 25
    $script:BoxLookfor.height = 10
    $script:BoxLookfor.location = New-Object System.Drawing.Point(780, 0)
    $script:BoxLookfor.Size = New-Object System.Drawing.Size(300, 20)
    $script:BoxLookfor.Visible = $false
    $script:Form.Controls.Add($script:BoxLookfor)

    $script:BoxErrorOptions = New-Object System.Windows.Forms.TextBox 
    $script:BoxErrorOptions.Text = ""
    $script:BoxErrorOptions.AutoSize = $true
    $script:BoxErrorOptions.Multiline = $true
    $script:BoxErrorOptions.Scrollbars = 'Vertical'
    $script:BoxErrorOptions.ReadOnly = $true
    $script:BoxErrorOptions.width = 25
    $script:BoxErrorOptions.height = 10
    $script:BoxErrorOptions.location = New-Object System.Drawing.Point(700, 80)
    $script:BoxErrorOptions.Size = New-Object System.Drawing.Size(380, 505)
    $script:BoxErrorOptions.Visible = $false
    $script:Form.Controls.Add($script:BoxErrorOptions)

    ######################################################################
    #------------------------- Buttons Section --------------------------#
    ######################################################################

    $script:ButtonConnect = New-Object System.Windows.Forms.Button
    $script:ButtonConnect.Location = New-Object System.Drawing.Point(20, 100)
    $script:ButtonConnect.Size = New-Object System.Drawing.Size(100, 30)
    $script:ButtonConnect.Text = 'Connect to portal'
    $script:ButtonConnect.Add_Click({ Invoke-ButtonConnectToPortal })
    $script:Form.Controls.Add($script:ButtonConnect)

    $script:ButtonPath = New-Object System.Windows.Forms.Button
    $script:ButtonPath.Location = New-Object System.Drawing.Point(585, 505)
    $script:ButtonPath.Size = New-Object System.Drawing.Size(100, 30)
    $script:ButtonPath.Text = 'Select'
    $script:ButtonPath.Visible = $false
    $script:ButtonPath.Add_Click({ $script:BoxPath.Text = Invoke-ButtonSelectPath -inputFolder $script:BoxPath.Text })
    $script:Form.Controls.Add($script:ButtonPath)

    $script:ButtonWebEditor = New-Object System.Windows.Forms.Button
    $script:ButtonWebEditor.Location = New-Object System.Drawing.Point(585, 540)
    $script:ButtonWebEditor.Size = New-Object System.Drawing.Size(100, 50)
    $script:ButtonWebEditor.Visible = $false
    $script:ButtonWebEditor.Text = 'Open Web Query Editor'
    $script:ButtonWebEditor.Add_Click({ Invoke-ButtonQueryWebEditor })
    $script:Form.Controls.Add($script:ButtonWebEditor)

    $script:ButtonValidateQuery = New-Object System.Windows.Forms.Button
    $script:ButtonValidateQuery.Location = New-Object System.Drawing.Point(20, 540)
    $script:ButtonValidateQuery.Size = New-Object System.Drawing.Size(120, 50)
    $script:ButtonValidateQuery.Text = 'Validate Query'
    $script:ButtonValidateQuery.Visible = $false
    $script:ButtonValidateQuery.ForeColor = 'orange'
    $script:ButtonValidateQuery.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10, [System.Drawing.FontStyle]::Bold)
    $script:ButtonValidateQuery.Add_Click({ Invoke-ButtonValidateQuery })
    $script:Form.Controls.Add($script:ButtonValidateQuery)

    $script:ButtonRunQuery = New-Object System.Windows.Forms.Button
    $script:ButtonRunQuery.Location = New-Object System.Drawing.Point(20, 540)
    $script:ButtonRunQuery.Size = New-Object System.Drawing.Size(120, 50)
    $script:ButtonRunQuery.Text = 'Run NXQL Query'
    $script:ButtonRunQuery.Visible = $false
    $script:ButtonRunQuery.ForeColor = 'green'
    $script:ButtonRunQuery.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10, [System.Drawing.FontStyle]::Bold)
    $script:ButtonRunQuery.Add_Click({ Invoke-ButtonRunNXQLQuery })
    $script:Form.Controls.Add($script:ButtonRunQuery)
}
function Invoke-FormEnvironment {

    $EnvSelect = New-Object system.Windows.Forms.Form
    $EnvSelect.ClientSize = New-Object System.Drawing.Point(390, 100)
    $EnvSelect.text = "Powershell NXQL API"
    $EnvSelect.TopMost = $true
    $EnvSelect.FormBorderStyle = 'FixedDialog'
    try {
        $p = (Get-Process powershell | Sort-Object -Property CPU -Descending | Select-Object -First 1).Path
        $EnvSelect.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($p)
    }
    catch {
        $p = (Get-Process explorer | Sort-Object -Property CPU -Descending | Select-Object -First 1).Path
        $EnvSelect.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($p)
    }

    ######################################################################
    #-------------------------- Labels Section --------------------------#
    ######################################################################
    
    $LabelQuestion = New-Object system.Windows.Forms.Label
    $LabelQuestion.text = "On which environment do you want to run query?"
    $LabelQuestion.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $LabelQuestion.AutoSize = $true
    $LabelQuestion.width = 370
    $LabelQuestion.height = 10
    $LabelQuestion.location = New-Object System.Drawing.Point(10, 10)
    $LabelQuestion.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 12)
    $EnvSelect.Controls.Add($LabelQuestion)

    ######################################################################
    #------------------------- Buttons Section --------------------------#
    ######################################################################

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
function Invoke-FormError {
    param(
        [Parameter(Mandatory = $true)]
        [String]$Message
    )
    [System.Windows.MessageBox]::Show($Message, 'Application Error', 'OK', 'Error') | Out-Null
    throw
}
######################################################################
#--------------------------- Forms Modes ----------------------------#
######################################################################
function Invoke-FormMainResize {
    # Function to change window mode
    param (
        [switch]$Big,
        [switch]$Options,
        [switch]$ValidQuery
    )
    if ($Big) {
        $script:Form.TopMost = $false
        $script:Form.ClientSize = New-Object System.Drawing.Point(700, 600)
        $script:BoxErrorOptions.visible = $False
        $script:LabelQuery.Visible = $true
        $script:BoxQuery.Visible = $true
        $script:LabelFileName.Visible = $true
        $script:BoxFileName.Visible = $true
        $script:LabelPath.Visible = $true
        $script:BoxPath.Visible = $true
        $script:ButtonPath.Visible = $true
        $script:LabelPlatform.Visible = $true
        $script:CheckboxWindows.Visible = $true
        $script:CheckboxMac_OS.Visible = $true
        $script:CheckboxMobile.Visible = $true
        if ($ValidQuery) {
            $script:ButtonValidateQuery.visible = $false
            $script:ButtonRunQuery.visible = $true
        }
        else {
            $script:ButtonRunQuery.visible = $false
            $script:ButtonValidateQuery.visible = $true
        }
        if ($options) {
            $script:Form.TopMost = $false
            $script:Form.ClientSize = New-Object System.Drawing.Point(1100, 600)
            $script:BoxErrorOptions.visible = $true
        }
    }
    else {
        $script:Form.TopMost = $true
        $script:Form.ClientSize = New-Object System.Drawing.Point(480, 150)
        $script:LabelQuery.Visible = $false
        $script:BoxQuery.Visible = $false
        $script:LabelFileName.Visible = $false
        $script:BoxFileName.Visible = $false
        $script:LabelPath.Visible = $false
        $script:BoxPath.Visible = $false
        $script:ButtonPath.Visible = $false
        $script:LabelPlatform.Visible = $false
        $script:CheckboxWindows.Visible = $false
        $script:CheckboxMac_OS.Visible = $false
        $script:CheckboxMobile.Visible = $false
        $script:ButtonWebEditor.Visible = $false
    }
}
function Invoke-Buttons {
    # Function to disable and enable action buttons
    param (
        [switch]$Enable
    )
    if ($Enable) {
        $script:ButtonConnect.enabled = $true
        $script:ButtonRunQuery.enabled = $true
        $script:ButtonPath.enabled = $true
        $script:BoxPath.enabled = $true
        $script:BoxFileName.enabled = $true
        $script:BoxPort.enabled = $true
        $script:CheckboxWindows.enabled = $true
        $script:CheckboxMac_OS.enabled = $true
        $script:CheckboxMobile.enabled = $true
    }
    else {
        $script:ButtonConnect.enabled = $false
        $script:ButtonRunQuery.enabled = $false
        $script:ButtonPath.enabled = $false
        $script:BoxPath.enabled = $false
        $script:BoxFileName.enabled = $false
        $script:BoxPort.enabled = $false
        $script:CheckboxWindows.enabled = $false
        $script:CheckboxMac_OS.enabled = $false
        $script:CheckboxMobile.enabled = $false
    }
}
######################################################################
#--------------------------- Forms Logic ----------------------------#
######################################################################
function Invoke-ButtonConnectToPortal {
    # Remember password while Button is clicked multiple times without changing anything
    $KeepCredentials = $false
    if ($script:LabelConnectionStatusDetails.Text -eq "Connected" -and 
        $script:PortalFQDN -eq $script:BoxPortal.Text -and
        $script:Login -eq $script:BoxLogin.Text) {
        $KeepCredentials = $true
    }
    else {
        $script:PortalFQDN = $script:BoxPortal.text
        $script:Login = $script:BoxLogin.Text
    }
    # Clear additional info to hide any information from previous run, which maybe misleading
    $script:LabelNumberOfEngines.Visible = $false
    $script:ButtonRunQuery.Visible = $false
    $script:LabelRunStatus.Visible = $false
    $script:LabelPort.Visible = $false
    $script:BoxPort.Visible = $false
    # Hide additional fields if button clicked multiple times
    Invoke-FormMainResize
    # Fill in the first part of output name and format portal connection details
    $script:BoxFileName.Text = (Get-Date).ToString("yyyy-MM-dd")
    $script:LabelConnectionStatus.Text = "Connection state:"
    $script:LabelConnectionStatus.ForeColor = "black"
    # Check if fields are not empty if yes exit
    if ($script:BoxLogin.Text.Length -lt 1 -or
        $script:BoxPassword.Text.Length -lt 1) {
        $script:BoxPassword.text = ""
        $script:LabelConnectionStatus.Text = "Login and password can not be empty !"
        $script:LabelConnectionStatus.ForeColor = "red"
        $script:LabelConnectionStatus.Visible = $true
        return
    }
    # Clear additional info to hide any information from previous run, which maybe misleading
    $script:LabelConnectionStatus.Visible = $true
    $script:LabelNumberOfEngines.Visible = $false
    $script:LabelNumberOfEngines.text = ""
    $script:LabelConnectionStatusDetails.text = "-"
    $script:LabelConnectionStatusDetails.ForeColor = "black"
    $Portal = $script:BoxPortal.Text
    # if login portal etc. are not changed and worked previously
    # use those credentials once agin
    if ($KeepCredentials -eq $false) {
        $Username = $script:BoxLogin.text
        $Password = $script:BoxPassword.text
        $Password = ConvertTo-SecureString $Password -AsPlainText -Force
        $script:Credentials = New-Object System.Management.Automation.PSCredential ($Username, $Password)
    }
    $script:BoxPassword.text = "********************"
    try {
        $script:Engines = Get-EngineList -portal $Portal -credentials $script:Credentials
    }
    catch {
        $script:LabelConnectionStatusDetails.text = "Not connected"
        $script:LabelConnectionStatusDetails.ForeColor = "red"
        $script:LabelConnectionStatusDetails.Visible = $true
        $script:BoxPassword.text = ""
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
    $script:LabelConnectionStatusDetails.Visible = $true
    $script:LabelConnectionStatusDetails.text = "Connected"
    $script:LabelConnectionStatusDetails.ForeColor = "green"
    $script:LabelNumberOfEngines.text = "Number of engines: $Number_of_engines"
    # Display additional fields
    Invoke-FormMainResize -Big
    $script:LabelNumberOfEngines.Visible = $true
    $script:ButtonRunQuery.Visible = $true
    # Check environment type SAAS / On-prem
    # Set additional details based on it
    if ($script:BoxPortal.Text -notlike "*.nexthink.cloud" ) {
        $script:BoxPort.Text = "1671"
        $script:BoxFileName.Text += " - <Customer_Name> - "
    }
    else {
        $script:CustomerName = $script:BoxPortal.Text.Split(".")[0]
        $script:BoxPort.Text = "443"
        $script:BoxFileName.Text += " - $script:CustomerName - "
    }
    # Display additional components of the GUI
    $script:ButtonWebEditor.Visible = $true
    $script:LabelPort.Visible = $true
    $script:BoxPort.Visible = $true
}
Function Invoke-ButtonSelectPath {
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
            return
        }
    }
    $browse.SelectedPath
    $browse.Dispose()
    $script:BoxPath.Text = $FolderBrowserDialog.SelectedPath
    return 
}
function Invoke-ButtonValidateQuery {
    Invoke-FormMainResize -Big 
    $Status = Invoke-QueryValidation -Query $script:BoxQuery.Text 
    if ($null -ne $Status) {
        Write-Host $Status
        $script:LabelRunStatus.Visible = $true
        $script:LabelRunStatus.ForeColor = "red"
        $script:LabelRunStatus.Text = $Status.'Error message'
        if (($Status.'Error Options').count -ne 0) {
            Write-Host $Status.'Error Options'
            Invoke-QueryOptions
        }
    }
    else {
        Write-Host "Valid Query"
        Invoke-FormMainResize -Big -ValidQuery
    }
}
function Invoke-CredentialCleanup {
    param(
        [switch]$Portal
    )
    Invoke-FormMainResize
    if ($Portal) {
        $script:ButtonRunQuery.Visible = $false
        $script:ButtonWebEditor.Visible = $false
        $script:LabelConnectionStatus.Visible = $false
        $script:LabelConnectionStatusDetails.Visible = $false
        $script:LabelNumberOfEngines.Visible = $false
        $script:LabelRunStatus.Visible = $false
        $script:BoxPassword.text = ""
        $script:BoxLogin.text = ""  
    }
    else {
        $script:ButtonRunQuery.Visible = $false
        $script:ButtonWebEditor.Visible = $false
        $script:LabelConnectionStatus.Visible = $false
        $script:LabelConnectionStatusDetails.Visible = $false
        $script:LabelNumberOfEngines.Visible = $false
        $script:LabelRunStatus.Visible = $false
        $script:BoxPassword.text = ""
        $script:LabelConnectionStatusDetails.Text = ""
    }
  
}
function Invoke-ButtonRunNXQLQuery {
    # Update Export status
    $script:LabelRunStatus.ForeColor = "orange"
    $script:LabelRunStatus.Text = "Proccessing..."
    $script:LabelRunStatus.Visible = $true
    # Disable Query Box for user to be unable to modify before reading it to variable
    $script:BoxQuery.enabled = $false
    Invoke-Buttons
    [String]$Query = $script:BoxQuery.Text
    $script:BoxQuery.enabled = $true
    $FileName = $script:BoxFileName.Text
    # Check Platform
    $Platform = @()
    if ($script:CheckboxWindows.checked) {
        $Platform += "windows"
    }
    if ($script:CheckboxMac_OS.checked) {
        $Platform += "mac_os"
    }
    if ($script:CheckboxMobile.checked) {
        $Platform += "mobile"
    }
    # Handling if no platform is selected
    if ($null -eq $Platform) {
        $script:LabelRunStatus.Visible = $true
        $script:LabelRunStatus.ForeColor = "red"
        $script:LabelRunStatus.Text = "There is no platform selected"
        Invoke-Buttons -Enable
        return
    }
    # Invoke Basic Query validation
    Invoke-QueryValidation -Query $Query
    if ($null -ne $script:ErrorInformation) {
        $script:LabelRunStatus.Visible = $true
        $script:LabelRunStatus.ForeColor = "red"
        $script:LabelRunStatus.Text = $script:ErrorInformation."Error message"
        Invoke-Buttons -Enable
        return
    }
    # Check if the file name does not contain any "/" or "\"
    if (($FileName.ToCharArray() | Where-Object { $_ -eq '/' } | Measure-Object).Count -gt 0 `
            -or `
        ($FileName.ToCharArray() | Where-Object { $_ -eq '\' } | Measure-Object).Count -gt 0 ) {
        
        $script:LabelFileName.ForeColor = "red"
        $script:LabelFileName.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 12, [System.Drawing.FontStyle]::Bold)
        Invoke-Buttons -Enable
        return
    }
    # If user set inccorret filename and on the next run it is correct remove red label
    $script:LabelFileName.ForeColor = "black"
    $script:LabelFileName.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 12)
    $WebAPIPort = $script:BoxPort.text
    $Path = $script:BoxPath.text
    $FileName = $script:BoxFileName.text
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
        $script:LabelRunStatus.Visible = $true
        $script:LabelRunStatus.ForeColor = "green"
        $script:LabelRunStatus.Text = $result
        Invoke-Popup -title "NXQL Export" -description "NXQL Export for $script:CustomerName is ready!"
    }
    else {
        $script:LabelRunStatus.Visible = $true
        $script:LabelRunStatus.ForeColor = "red"
        $script:LabelRunStatus.Text = $result
        $result = $result.Split(":")
        Invoke-Popup -title "FAIL NXQL Export" -description "NXQL Export for $script:CustomerName failed with error: $result"
    }
    Invoke-Buttons -Enable
}
function Invoke-ButtonQueryWebEditor {
    # Select one of the engines
    $engine = ($script:Engines | Select-Object -First 1).address
    # Create a link to NXQL web editor
    $WebEditorAddress = "https://$engine/2/editor/nxql_editor.html"
    # Run the link
    Start-Process "$WebEditorAddress"
}
function Invoke-BoxQuery {
    # Invoke Basic Query validation
    $script:ButtonRunQuery.visible = $false
    $script:ButtonValidateQuery.visible = $true
                
    $Status = Invoke-QueryValidation -Query $script:BoxQuery.Text -Ligth
    if (($null -ne $Status)) {
        $script:LabelRunStatus.Visible = $true
        $script:LabelRunStatus.ForeColor = "red"
        $script:LabelRunStatus.Text = $Status.'Error message'
    }
    else {
        $script:LabelLookup.visible = $false
        $script:BoxLookfor.visible = $false
        $script:BoxErrorOptions.visible = $false
        $script:LabelRunStatus.Visible = $false
    }
}
function Get-BoxPathLocation {
    if ((Get-Location).Path -like "*main") {
        $path = (Get-Location).Path.Split("\")[0..((Get-Location).Path.Split("\").count - 2)] -join "\"
    }
    else {
        $path = (Get-Location).Path
    }
    return $path
}
function Invoke-CheckboxPlatform {
    $Status = Invoke-QueryValidation -Query $script:BoxQuery.Text 
    if ($script:LabelRunStatus.Text -eq "Proccessing...") {
        return
    }
    if (($null -ne $Status)) {
        $script:LabelRunStatus.Visible = $true
        $script:LabelRunStatus.ForeColor = "red"
        $script:LabelRunStatus.Text = $Status
    }
    else {
        $script:LabelRunStatus.Visible = $false
    }
}
function Invoke-CheckboxShowPassword {
    if ($script:CheckboxShowPassword.checked) {
        $script:BoxPassword.passwordchar = $null
    }
    else {
        $script:BoxPassword.passwordchar = "*"
    }
}
######################################################################
#----------------------- Operation Functions ------------------------#
######################################################################
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
        Invoke-FormEnvironment
        if ($script:Environment -eq "FITS") {
            $engineList = $engineList | Where-Object { $_.name -in $FITS }
        }
        elseif ($script:Environment -eq "MoJo") {
            $engineList = $engineList | Where-Object { $_.name -in $MOJO }
        }
    }
    return $engineList
}
function Invoke-QueryValidation {
    param (
        [String]$Query,
        [Switch]$Ligth
    )
    if ($Query.Length -le 19) {
        return $null
    }
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
    # Check if query is not empty
    if ($Query.Length -le 1) {
        return "NXQL query can not be blank !"
    }
    # Check if select statement exists
    if (($Query -notlike "*select*")) {
        return "There is no `"select`" statement !"
    }
    # Check if from statement exists
    if (($Query -notlike "*from*")) {
        return "There is no `"from`" statement !"
    }
    # Check if limit statement exists
    if (($Query -notlike "*limit*")) {
        return "There is no `"limit`" statement at the end of the query!"
    }
    if ($Ligth) {
        return $null
    }
    # Check Platform
    $Platform = @()
    if ($script:CheckboxWindows.checked) {
        $Platform += "windows"
    }
    if ($script:CheckboxMac_OS.checked) {
        $Platform += "mac_os"
    }
    if ($script:CheckboxMobile.checked) {
        $Platform += "mobile"
    }
    [String]$Query = $script:BoxQuery.Text
    $Query = $Query -replace "\(limit [0-9]+\)", "(limit 0)"
    $Engine = ($script:Engines | Select-Object -First 1).address
    $WebAPIPort = $script:BoxPort.text
  
    $script:ErrorInformation = Invoke-NXTEngineQueryValidation `
        -ServerName $Engine `
        -PortNumber $WebAPIPort `
        -credentials $script:Credentials `
        -Query $Query `
        -Platforms $Platform
    
    return $script:ErrorInformation
}
Function Invoke-NXTEngineQueryValidation {
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
        $webclient.DownloadString($Url) | Out-Null
    }
    catch [System.Net.WebException] {
        $streamReader = [System.IO.StreamReader]::new($_.Exception.Response.GetResponseStream())
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
        throw 'Not able to retriev data'
    }
    return $null
}
function Invoke-QueryOptions {
    Invoke-FormMainResize -Big -Options
    $script:LabelLookup.visible = $true
    $script:BoxLookfor.visible = $true
    $script:BoxErrorOptions.visible = $true
    $script:BoxErrorOptions.Text = ($script:ErrorInformation."Error Options")
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
        [String]$webapiPort = "443",
        [Parameter(Mandatory = $false)]
        [String[]]$Platform,
        [Parameter(Mandatory = $false)]
        [string]$DestinationPath,
        [Parameter(Mandatory = $false)]
        [string]$SyncPath,
        [Parameter(Mandatory = $false)]
        [string]$LogPath
    )
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
    if (!$credentials) {
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
        $LogsToDelete = (Get-ChildItem -Path $LogPath -Filter *.csv).FullName
        foreach ($file in $LogsToDelete) {
            Remove-Item -Path $file -Confirm:$false -Force
        }
    }
    if (Test-Path -Path "$LogPath\BadRequest") {
        Remove-Item -Path "$LogPath\BadRequest" -Confirm:$false -Force
    }
    # Create separate process for each engine in scope
    foreach ($Engine in $EngineList) {
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
        $ErrorMessage = Get-Content -Path "$LogPath\BadRequest"
        return $ErrorMessage
    }
    else {
        return "Failed: Error unknown"
    }
}
Function Invoke-Popup {
    <#
.SYNOPSIS
Display windows 10 pop-up message

.DESCRIPTION
Display Windows 10 pop-up message based on the title and description provided
Powershell icon will be visible in the pop-up

.PARAMETER title
Message title 1 line

.PARAMETER description
Message description multiple lines

.EXAMPLE
Invoke-Popup -title "Report "ready" `
			 -description "Automatically created report is ready"


.INPUTS
String

.OUTPUTS
None

.NOTES
    Author:  Stanislaw Horna
#>
    param (
        [Parameter(Mandatory = $true)]
        [string] $title,
        [Parameter(Mandatory = $true)]
        [String] $description
    )
    $global:endmsg = New-Object System.Windows.Forms.Notifyicon
    $endmsg.BalloonTipTitle = $title
    $endmsg.BalloonTipText = $description
    $endmsg.Visible = $true
    try {
        $p = (Get-Process powershell | Sort-Object -Property CPU -Descending | Select-Object -First 1).Path
        $endmsg.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($p)
    }
    catch {
        $p = (Get-Process explorer | Sort-Object -Property CPU -Descending | Select-Object -First 1).Path
        $endmsg.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($p)
    }
    $endmsg.ShowBalloonTip(10)
}
Invoke-main