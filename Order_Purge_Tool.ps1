#
##Version Check
#$scriptVersion = '1.0.0'  # Example version format
#$githubVersionUrl = 'https://github.com/conmoore67/Order_Purge_Tool/blob/main/Order_Purge_Tool.ps1'  
#$latestVersion = Invoke-RestMethod -Uri $githubVersionUrl
#function Compare-Version($version1, $version2) {
#    $v1 = New-Object -TypeName System.Version -ArgumentList $version1
#    $v2 = New-Object -TypeName System.Version -ArgumentList $version2
#    return $v1.CompareTo($v2)
#}
#
#$isUpdateRequired = Compare-Version -version1 $scriptVersion -version2 $latestVersion -lt 0
#
#if ($isUpdateRequired) {
#    $newScriptUrl = 'https://raw.githubusercontent.com/username/repository/branch/YourScript.ps1'  # Replace with actual URL
#    $newScriptPath = $MyInvocation.MyCommand.Path  # Gets the path of the running script
#
#    Invoke-WebRequest -Uri $newScriptUrl -OutFile $newScriptPath
#
#    Write-Host "Script updated to latest version. Please rerun the script."
#    exit
#} else {
#    # Notify the user that the application is up to date
#    Add-Type -AssemblyName System.Windows.Forms
#
#    $form = New-Object System.Windows.Forms.Form
#    $form.Text = 'Application Up to Date'
#    $form.Size = New-Object System.Drawing.Size(300, 150)
#    $form.StartPosition = 'CenterScreen'
#
#    $label = New-Object System.Windows.Forms.Label
#    $label.Location = New-Object System.Drawing.Point(10, 10)
#    $label.Size = New-Object System.Drawing.Size(280, 80)
#    $label.Text = "The application is already at the most current version ($scriptVersion)."
#    $form.Controls.Add($label)
#
#    $okButton = New-Object System.Windows.Forms.Button
#    $okButton.Location = New-Object System.Drawing.Point(100, 100)
#    $okButton.Size = New-Object System.Drawing.Size(100, 23)
#    $okButton.Text = 'OK'
#    $okButton.Add_Click({ $form.Close() })
#    $form.Controls.Add($okButton)
#
#    $form.ShowDialog()
#}


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
 
 
########## Create windows Forms and pull needed data ##########
 
# Add Windows Forms reference
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

# Create the form
$form = New-Object System.Windows.Forms.Form
$form.Text = 'UPT Order Purge Tool'
$form.Size = New-Object System.Drawing.Size(450, 550)  # Adjusted form height
$form.StartPosition = 'CenterScreen'

# Create the banner label
$bannerLabel = New-Object System.Windows.Forms.Label
$bannerLabel.Location = New-Object System.Drawing.Point(10, 10)
$bannerLabel.Size = New-Object System.Drawing.Size(430, 20)
$bannerLabel.Text = 'UPT Order Purge Tool'
$bannerLabel.Font = New-Object System.Drawing.Font('Arial', 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($bannerLabel)

# Create UPT Logo
$base64Image = "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"
$imageBytes = [Convert]::FromBase64String($base64Image)
$memoryStream = New-Object System.IO.MemoryStream($imageBytes, 0, $imageBytes.Length)
$memoryStream.Position = 0
$image = [System.Drawing.Image]::FromStream($memoryStream)

# Create the PictureBox
$pictureBox = New-Object System.Windows.Forms.PictureBox
$pictureBox.SizeMode = 'Zoom'
$pictureBox.Location = New-Object System.Drawing.Point(10, 40)
$pictureBox.Size = New-Object System.Drawing.Size(430, 100)
$pictureBox.Image = $image
$form.Controls.Add($pictureBox)

# Create the Move Number label and textbox
$labelMoveNumber = New-Object System.Windows.Forms.Label
$labelMoveNumber.Location = New-Object System.Drawing.Point(10, 150)
$labelMoveNumber.Size = New-Object System.Drawing.Size(430, 20)
$labelMoveNumber.Text = 'Enter Move/Order Number:'
$form.Controls.Add($labelMoveNumber)

$textboxMoveNumber = New-Object System.Windows.Forms.TextBox
$textboxMoveNumber.Location = New-Object System.Drawing.Point(10, 180)
$textboxMoveNumber.Size = New-Object System.Drawing.Size(420, 20)  # Narrowed width
$form.Controls.Add($textboxMoveNumber)

# Create the Confirm Move Number label and textbox
$labelConfirmMoveNumber = New-Object System.Windows.Forms.Label
$labelConfirmMoveNumber.Location = New-Object System.Drawing.Point(10, 210)
$labelConfirmMoveNumber.Size = New-Object System.Drawing.Size(430, 20)
$labelConfirmMoveNumber.Text = 'Confirm Move/Order Number:'
$form.Controls.Add($labelConfirmMoveNumber)

$textboxConfirmMoveNumber = New-Object System.Windows.Forms.TextBox
$textboxConfirmMoveNumber.Location = New-Object System.Drawing.Point(10, 240)
$textboxConfirmMoveNumber.Size = New-Object System.Drawing.Size(420, 20)  # Narrowed width
$form.Controls.Add($textboxConfirmMoveNumber)

# Create the Database Selection label and ComboBox
$labelDatabase = New-Object System.Windows.Forms.Label
$labelDatabase.Location = New-Object System.Drawing.Point(10, 270)
$labelDatabase.Size = New-Object System.Drawing.Size(430, 20)
$labelDatabase.Text = 'Select Database:'
$form.Controls.Add($labelDatabase)

$comboBoxDatabase = New-Object System.Windows.Forms.ComboBox
$comboBoxDatabase.Location = New-Object System.Drawing.Point(10, 300)
$comboBoxDatabase.Size = New-Object System.Drawing.Size(420, 20)  # Narrowed width
$comboBoxDatabase.Items.Add('TMW_UPT_PROD')
$comboBoxDatabase.Items.Add('TMW_UPT_TEST')
$form.Controls.Add($comboBoxDatabase)

# Create the OK button
$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(175, 340)  # Centered position
$okButton.Size = New-Object System.Drawing.Size(100, 23)
$okButton.Text = 'Purge Order'
$form.Controls.Add($okButton)

# Add Click event for OK button
$okButton.Add_Click({
    if ($textboxMoveNumber.Text -eq $textboxConfirmMoveNumber.Text) {
        $databaseName = $comboBoxDatabase.SelectedItem
        # Additional logic for processing the confirmed move number and selected database
        $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.Close()
    } else {
        [System.Windows.Forms.MessageBox]::Show("The entered Move Numbers do not match. Please re-enter.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK)
    }
})

# Show the form as a dialog
$form.ShowDialog()
$form.Topmost = $true
$result = $form.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
    $moveNumber = $textboxMoveNumber.Text
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
$csvFileName = "$CurrentDate $moveNumber.csv"
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
