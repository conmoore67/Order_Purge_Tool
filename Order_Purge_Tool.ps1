
########## Verrify Script Version ##########
# Add Windows Forms reference
Add-Type -AssemblyName System.Windows.Forms

# Current script version
$localVersion = '1.3.0'  # Example version format

# GitHub URL for the raw script
$githubUrl = 'https://raw.githubusercontent.com/conmoore67/Order_Purge_Tool/5d6e0b3e70dafd6650221a7f88068d77c8a8b6f6/Order_Purge_Tool.ps1'

# Function to extract version number from script content
function Get-ScriptVersion($scriptContent) {
    $versionLine = $scriptContent -match "\$version = '(\d+\.\d+\.\d+)'" | Out-Null
    $scriptVersion = $matches[1]
    return $scriptVersion
}

# Get the latest script content from GitHub
$latestScriptContent = Invoke-RestMethod -Uri $githubUrl

# Extract version from the latest script
$latestVersion = Get-ScriptVersion -scriptContent $latestScriptContent

# Function to create OK button
function CreateOkButton($form) {
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = 'OK'
    $okButton.Size = New-Object System.Drawing.Size(75, 23)
    $okButton.Location = New-Object System.Drawing.Point(100, 100)
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $okButton
    $form.Controls.Add($okButton)
}

# Compare versions and show message in a form
if ([version]$localVersion -lt [version]$latestVersion) {
    # Create and configure the form for the update message
    $formUpdate = New-Object System.Windows.Forms.Form
    $formUpdate.Text = "Update Available"
    $formUpdate.Size = New-Object System.Drawing.Size(300, 200)
    $formUpdate.StartPosition = 'CenterScreen'

    # Create the label for the update message
    $labelUpdate = New-Object System.Windows.Forms.Label
    $labelUpdate.Location = New-Object System.Drawing.Point(10, 10)
    $labelUpdate.Size = New-Object System.Drawing.Size(280, 80)
    $labelUpdate.Text = "A newer version ($latestVersion) is available. Updating..."
    $formUpdate.Controls.Add($labelUpdate)

    # Create OK button
    CreateOkButton -form $formUpdate

    # Show the form as a dialog
    $formUpdate.ShowDialog()

    # Download and save the new script
    $newScriptPath = "C:\temp\Order_Purge_Tool.ps1"  # Specify the path for the new script
    Invoke-WebRequest -Uri $githubUrl -OutFile $newScriptPath

    # Execute the new script
    Start-Process $newScriptPath
    exit
} else {
    # Create and configure the form for the "latest version" message
    $formLatest = New-Object System.Windows.Forms.Form
    $formLatest.Text = "Latest Version"
    $formLatest.Size = New-Object System.Drawing.Size(300, 200)
    $formLatest.StartPosition = 'CenterScreen'

    # Create the label for the "latest version" message
    $labelLatest = New-Object System.Windows.Forms.Label
    $labelLatest.Location = New-Object System.Drawing.Point(10, 10)
    $labelLatest.Size = New-Object System.Drawing.Size(280, 80)
    $labelLatest.Text = "You're running the latest version ($localVersion)."
    $formLatest.Controls.Add($labelLatest)

    # Create OK button
    CreateOkButton -form $formLatest

    # Show the form as a dialog
    $formLatest.ShowDialog()
}


########## Check/Install/Import needed modules ##########
# Load necessary assembly for Windows Forms
Add-Type -AssemblyName System.Windows.Forms

# Create a new form for the progress bar
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Installation Progress'
$form.Size = New-Object System.Drawing.Size(300, 100)
$form.StartPosition = 'CenterScreen'

# Create a progress bar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(10, 10)
$progressBar.Size = New-Object System.Drawing.Size(260, 20)
$progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
$form.Controls.Add($progressBar)

# Function to update progress bar
function Update-ProgressBar {
    param($value)
    $progressBar.Value = $value
    $form.Refresh()
}

# Show form in a separate thread to prevent script from pausing
$form.Show()

# Function to check and install modules
function CheckAndInstallModule {
    param($moduleName)
    if (-not (Get-Module -Name $moduleName -ListAvailable -ErrorAction SilentlyContinue)) {
        Install-Module -Name $moduleName -Force -Scope CurrentUser -Confirm:$false -ErrorAction SilentlyContinue
    }
}

Set-ExecutionPolicy Bypass -force -Confirm:$false

# NuGet Package Provider
Update-ProgressBar 10
CheckAndInstallModule -moduleName 'NuGet'

# SharePointPnPPowerShellOnline Module
Update-ProgressBar 30
CheckAndInstallModule -moduleName 'SharePointPnPPowerShellOnline'

# Microsoft.Online.SharePoint.PowerShell Module
Update-ProgressBar 60
CheckAndInstallModule -moduleName 'Microsoft.Online.SharePoint.PowerShell'

# SQLSERVER Module
Update-ProgressBar 90
CheckAndInstallModule -moduleName 'SQLSERVER'

# Complete
Update-ProgressBar 100
[System.Windows.Forms.MessageBox]::Show("Installation Complete", "Info", [System.Windows.Forms.MessageBoxButtons]::OK)
$form.Close()
 
 
########## Check/Install/Import needed modules ##########

# Load necessary assembly for Windows Forms
Add-Type -AssemblyName System.Windows.Forms

# Create a new form for the progress bar
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Installation Progress'
$form.Size = New-Object System.Drawing.Size(300, 100)
$form.StartPosition = 'CenterScreen'

# Create a progress bar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(10, 10)
$progressBar.Size = New-Object System.Drawing.Size(260, 20)
$progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
$form.Controls.Add($progressBar)

# Function to update progress bar
function Update-ProgressBar {
    param($value)
    $progressBar.Value = $value
    $form.Refresh()
}

# Show form in a separate thread to prevent script from pausing
$form.Show()

# Function to check and install modules
function CheckAndInstallModule {
    param($moduleName)
    if (-not (Get-Module -Name $moduleName -ListAvailable -ErrorAction SilentlyContinue)) {
        Install-Module -Name $moduleName -Force -Scope CurrentUser -Confirm:$false -ErrorAction Stop
    }
    if (-not (Get-Module -Name $moduleName -ListAvailable -ErrorAction SilentlyContinue)) {
        throw "Failed to install module: $moduleName"
    }
}

# Check and install modules with error handling
try {
    # Set execution policy with error handling
    Set-ExecutionPolicy Bypass -force -Confirm:$false -ErrorAction Stop

    # NuGet Package Provider
    Update-ProgressBar 10
    CheckAndInstallModule -moduleName 'NuGet'

    # SharePointPnPPowerShellOnline Module
    Update-ProgressBar 30
    CheckAndInstallModule -moduleName 'SharePointPnPPowerShellOnline'

    # Microsoft.Online.SharePoint.PowerShell Module
    Update-ProgressBar 60
    CheckAndInstallModule -moduleName 'Microsoft.Online.SharePoint.PowerShell'

    # SQLSERVER Module
    Update-ProgressBar 90
    CheckAndInstallModule -moduleName 'SQLSERVER'

    # Complete
    Update-ProgressBar 100
    [System.Windows.Forms.MessageBox]::Show("Installation Complete", "Info", [System.Windows.Forms.MessageBoxButtons]::OK)
} catch {
    # Close form and display admin permissions message
    $form.Close()
    [System.Windows.Forms.MessageBox]::Show("Please rerun the script with admin permissions.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    Exit
} finally {
    $form.Close()
}




########## Run SQL Query/Backup/Purge ##########
# Define server, database, and movenumber

$serverName = "SQL00.otls1.otl-upt.com"
 
# SQL Query to retrieve data
$sqlSearchQuery = @"
DECLARE @MOVNUMBER INT
SET @MOVNUMBER = '$moveNumber'
 
-- Selecting from ORDERHEADER
SELECT *
FROM ORDERHEADER ORD (NOLOCK)
WHERE ORD.MOV_NUMBER = @MOVNUMBER
 
-- Selecting from LEGHEADER
SELECT *
FROM LEGHEADER L (NOLOCK)
WHERE L.MOV_NUMBER = @MOVNUMBER
 
-- Selecting from STOPS
SELECT *
FROM STOPS (NOLOCK)
WHERE MOV_NUMBER = @MOVNUMBER
ORDER BY STOPS.STP_ARRIVALDATE
 
-- Selecting from FREIGHTDETAIL joined with STOPS
SELECT F.*
FROM FREIGHTDETAIL F (NOLOCK)
JOIN STOPS S (NOLOCK) ON F.STP_NUMBER = S.STP_NUMBER
WHERE S.MOV_NUMBER = @MOVNUMBER
ORDER BY F.FGT_SEQUENCE, F.FGT_DISPLAY_SEQUENCE
 
-- Selecting from FREIGHT_BY_COMPARTMENT joined with FREIGHTDETAIL and STOPS
SELECT FBC.*
FROM FREIGHT_BY_COMPARTMENT FBC (NOLOCK)
JOIN FREIGHTDETAIL F (NOLOCK) ON FBC.FGT_NUMBER = F.FGT_NUMBER
JOIN STOPS S (NOLOCK) ON F.STP_NUMBER = S.STP_NUMBER
WHERE S.MOV_NUMBER = @MOVNUMBER
 
-- Selecting from OILFIELDREADINGS joined with FREIGHTDETAIL and STOPS
SELECT OFR.*
FROM OILFIELDREADINGS OFR (NOLOCK)
JOIN FREIGHTDETAIL F (NOLOCK) ON OFR.FGT_NUMBER = F.FGT_NUMBER
JOIN STOPS S (NOLOCK) ON F.STP_NUMBER = S.STP_NUMBER
WHERE S.MOV_NUMBER = @MOVNUMBER
 
-- Selecting from EVENT joined with STOPS
SELECT E.*
FROM EVENT E (NOLOCK)
JOIN STOPS S (NOLOCK) ON E.STP_NUMBER = S.STP_NUMBER
WHERE S.MOV_NUMBER = @MOVNUMBER
 
-- Selecting from SHIFTSCHEDULES joined with LEGHEADER
SELECT SS.*
FROM SHIFTSCHEDULES SS (NOLOCK)
JOIN LEGHEADER L (NOLOCK) ON SS.SS_ID = L.SHIFT_SS_ID
WHERE L.MOV_NUMBER = @MOVNUMBER
 
-- Selecting from ASSETASSIGNMENT
SELECT *
FROM ASSETASSIGNMENT (NOLOCK)
WHERE MOV_NUMBER = @MOVNUMBER
 
-- Selecting from REFERENCENUMBER joined with ORDERHEADER
SELECT R.*
FROM REFERENCENUMBER R (NOLOCK)
JOIN ORDERHEADER O (NOLOCK) ON R.ORD_HDRNUMBER = O.ORD_HDRNUMBER
WHERE O.MOV_NUMBER = @MOVNUMBER
 
-- Selecting from INVOICEHEADER
SELECT *
FROM INVOICEHEADER IH (NOLOCK)
WHERE IH.MOV_NUMBER = @MOVNUMBER
 
-- Selecting from INVOICEDETAIL joined with INVOICEHEADER
SELECT ID.*
FROM INVOICEDETAIL ID (NOLOCK)
JOIN INVOICEHEADER IH (NOLOCK) ON ID.IVH_HDRNUMBER = IH.IVH_HDRNUMBER
WHERE IH.MOV_NUMBER = @MOVNUMBER
"@
 
# Retrieve data from SQL Server
$data = Invoke-Sqlcmd -Query $sqlSearchQuery -ServerInstance $serverName -Database $databaseName
 
# Export data to CSV
$CurrentDate = Get-Date -Format "yyyy-MM-dd"
$csvFileName = "$databaseName $CurrentDate $moveNumber.csv"
$localPath = "C:\temp\$csvFileName"
$data | Export-Csv -Path $localPath -NoTypeInformation
 
# Upload CSV to SharePoints
# Ensures SharePoint PnP PowerShell is installed and configured
# SharePoint details
$PNPLEGACYMESSAGE = $false
$sharePointSite = "https://otlupt.sharepoint.com/sites/InformationServices"
$sharePointFolder = "/Shared Documents/Jamie Mitchell/Purges/2023"
 
 
# Add Windows Forms reference
Add-Type -AssemblyName System.Windows.Forms

# Upload backup to SharePoint
Connect-PnPOnline -Url $sharePointSite -UseWebLogin -WarningAction Ignore
Add-PnPFile -Path $localPath -Folder $sharePointFolder -WarningAction Ignore
 
# Confirm if the file was uploaded
$fileUploaded = Get-PnPFile -Url "$sharePointFolder/$csvFileName" -ErrorAction SilentlyContinue -WarningAction Ignore

if ($fileUploaded) {
    # Progress bar setup
    $progress = @{
        Activity = "Purging SQL Data"
        Status = "In progress"
        PercentComplete = 0
    }

    Write-Progress @progress

    # Purge SQL data
    $purgeQuery = "PURGE_DELETE $moveNumber ,0"
    
    # Execute SQLCMD with progress
    Invoke-Sqlcmd -Query $purgeQuery -ServerInstance $serverName -Database $databaseName

    # Update progress bar to complete
    $progress.PercentComplete = 100
    Write-Progress @progress
    Start-Sleep -Seconds 2 # Pause to show the completed progress bar

    # Display completion message
    [System.Windows.Forms.MessageBox]::Show("Data for move number $moveNumber has been backed up to the UPT IT SharePoint and purged from SQL00.", "Backup and Purge Completed")

} else {
    # Add Windows Forms reference
    Add-Type -AssemblyName System.Windows.Forms
    
    # Create and configure the form for the message box
    $formError = New-Object System.Windows.Forms.Form
    $formError.Text = "Error!!"
    $formError.Size = New-Object System.Drawing.Size(400, 250)  # Adjust the size as needed
    $formError.StartPosition = 'CenterScreen'
    
    # Create the bold label for the "Backup Unsuccessful!!" message
    $labelBoldError = New-Object System.Windows.Forms.Label
    $labelBoldError.Location = New-Object System.Drawing.Point(10, 10)
    $labelBoldError.Size = New-Object System.Drawing.Size(380, 30)  # Adjust the size as needed
    $labelBoldError.Text = "Backup Unsuccessful!!"
    $labelBoldError.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
    $formError.Controls.Add($labelBoldError)
    
    # Create the standard label for the description
    $labelDescription = New-Object System.Windows.Forms.Label
    $labelDescription.Location = New-Object System.Drawing.Point(10, 50)
    $labelDescription.Size = New-Object System.Drawing.Size(380, 120)  # Adjust the size as needed
    $labelDescription.Text = "Data purge halted to prevent the loss of order files. Please run the script with your Admin account and enter your credentials on the SharePoint screen. Should this issue persist, clear your cached credentials and try the order purge process again."
    $labelDescription.Font = New-Object System.Drawing.Font("Arial", 10)
    $formError.Controls.Add($labelDescription)
    
    # Create an OK button for the user to close the message box
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(150, 180)  # Adjusted position
    $okButton.Size = New-Object System.Drawing.Size(100, 23)
    $okButton.Text = "OK"
    $okButton.Add_Click({$formError.Close()})
    $formError.Controls.Add($okButton)
    
    # Show the form as a dialog
    $formError.ShowDialog()
}
