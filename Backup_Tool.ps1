Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[Windows.Forms.Application]::EnableVisualStyles()

$form = New-Object Windows.Forms.Form
$form.Text = "Backup Tool"
$form.Size = New-Object Drawing.Size(400, 220)
$form.StartPosition = "CenterScreen"

$label1 = New-Object Windows.Forms.Label
$label1.Location = New-Object Drawing.Point(10, 20)
$label1.Size = New-Object Drawing.Size(150, 20)
$label1.Text = "Source Path:"
$form.Controls.Add($label1)

$textBox1 = New-Object Windows.Forms.TextBox
$textBox1.Location = New-Object Drawing.Point(160, 20)
$textBox1.Size = New-Object Drawing.Size(150, 20)
$form.Controls.Add($textBox1)

$buttonBrowseSource = New-Object Windows.Forms.Button
$buttonBrowseSource.Location = New-Object Drawing.Point(320, 20)
$buttonBrowseSource.Size = New-Object Drawing.Size(30, 20)
$buttonBrowseSource.Text = "..."
$buttonBrowseSource.Add_Click({
    $folderBrowserDialog = New-Object Windows.Forms.FolderBrowserDialog
    $folderBrowserDialog.Description = "Select Source Folder"
    $folderBrowserDialog.RootFolder = "MyComputer"
    if ($folderBrowserDialog.ShowDialog() -eq [Windows.Forms.DialogResult]::OK) {
        $textBox1.Text = $folderBrowserDialog.SelectedPath
    }
})
$form.Controls.Add($buttonBrowseSource)

$label2 = New-Object Windows.Forms.Label
$label2.Location = New-Object Drawing.Point(10, 50)
$label2.Size = New-Object Drawing.Size(150, 20)
$label2.Text = "Destination Path:"
$form.Controls.Add($label2)

$textBox2 = New-Object Windows.Forms.TextBox
$textBox2.Location = New-Object Drawing.Point(160, 50)
$textBox2.Size = New-Object Drawing.Size(150, 20)
$form.Controls.Add($textBox2)

$buttonBrowseDestination = New-Object Windows.Forms.Button
$buttonBrowseDestination.Location = New-Object Drawing.Point(320, 50)
$buttonBrowseDestination.Size = New-Object Drawing.Size(30, 20)
$buttonBrowseDestination.Text = "..."
$buttonBrowseDestination.Add_Click({
    $folderBrowserDialog = New-Object Windows.Forms.FolderBrowserDialog
    $folderBrowserDialog.Description = "Select Destination Folder"
    $folderBrowserDialog.RootFolder = "MyComputer"
    if ($folderBrowserDialog.ShowDialog() -eq [Windows.Forms.DialogResult]::OK) {
        $textBox2.Text = $folderBrowserDialog.SelectedPath
    }
})
$form.Controls.Add($buttonBrowseDestination)

$label4 = New-Object Windows.Forms.Label
$label4.Location = New-Object Drawing.Point(10, 80)
$label4.Size = New-Object Drawing.Size(150, 20)
$label4.Text = "Number of Versions:"
$form.Controls.Add($label4)

$numericUpDown2 = New-Object Windows.Forms.NumericUpDown
$numericUpDown2.Location = New-Object Drawing.Point(160, 80)
$numericUpDown2.Size = New-Object Drawing.Size(80, 20)
$form.Controls.Add($numericUpDown2)

$label3 = New-Object Windows.Forms.Label
$label3.Location = New-Object Drawing.Point(10, 110)
$label3.Size = New-Object Drawing.Size(150, 20)
$label3.Text = "Interval (Minutes):"
$form.Controls.Add($label3)

$numericUpDown = New-Object Windows.Forms.NumericUpDown
$numericUpDown.Location = New-Object Drawing.Point(160, 110)
$numericUpDown.Size = New-Object Drawing.Size(80, 20)
$form.Controls.Add($numericUpDown)

$labelDescription = New-Object Windows.Forms.Label
$labelDescription.Location = New-Object Drawing.Point(270, 82)
$labelDescription.Size = New-Object Drawing.Size(150, 20)
$labelDescription.Text = "ZIP-Archiv:"
$labelDescription.AutoSize = $true
$form.Controls.Add($labelDescription)

$checkBox = New-Object Windows.Forms.CheckBox
$checkBox.Location = New-Object Drawing.Point(335, 80)
$checkBox.Size = New-Object Drawing.Size(80, 20)
$checkBox.Checked = $true
$form.Controls.Add($checkBox)

$button1 = New-Object Windows.Forms.Button
$button1.Location = New-Object Drawing.Point(10, 140)
$button1.Size = New-Object Drawing.Size(180, 30)
$button1.Text = "Create Backup Job"
$button1.Add_Click({
    # Insert code for creating the backup job here
    $sourcePath = $textBox1.Text.Trim()
    $destinationPath = $textBox2.Text.Trim()
    $numVersions = $numericUpDown2.Value
    $interval = $numericUpDown.Value
    $useZip = $checkBox.Checked

    # Input validation
    if (-not $sourcePath -or -not $destinationPath) {
        [Windows.Forms.MessageBox]::Show("Invalid paths. Please check your inputs.")
        return
    }

    try {
        # Save PowerShell script in .ps1 file in the backup tool folder
        if ($useZip) {
            $psScriptContent = @"
            `$sourcePath = '$sourcePath'
            `$destinationPath = '$destinationPath'
            `$maxVersions = $numVersions
            `$currentDateTime = Get-Date -Format 'yyyyMMdd-HHmmss'
            `$newFolderName = 'Backup_' + `$currentDateTime
            `$newFolderPath = Join-Path -Path `$destinationPath -ChildPath `$newFolderName
            New-Item -Path `$newFolderPath -ItemType Directory
            robocopy `$sourcePath `$newFolderPath /MIR /V /LOG+:C:\BackupLog.txt

            # Dateien in ein Zip-Archiv komprimieren
            `$zipFileName = Join-Path -Path `$destinationPath -ChildPath "`$newFolderName.zip"
            Compress-Archive -Path `$newFolderPath -DestinationPath `$zipFileName -Force

            # Ordner nach dem Erstellen des Zip-Archivs löschen
            Remove-Item -Path `$newFolderPath -Recurse -Force

            # Prüfen und löschen älterer Zip-Archive, wenn die maximale Anzahl erreicht ist
            `$existingZips = Get-ChildItem -Path `$destinationPath -Filter 'Backup_*.zip' | Sort-Object LastWriteTime -Descending
            if (`$existingZips.Count -ge `$maxVersions) {
                `$zipsToDelete = `$existingZips | Select-Object -Skip `$maxVersions
                foreach (`$zip in `$zipsToDelete) {
                    Remove-Item -Path `$zip.FullName -Force
                }
            }
"@
        } else {
            $psScriptContent = @"
            `$sourcePath = '$sourcePath'
            `$destinationPath = '$destinationPath'
            `$maxVersions = $numVersions
            `$currentDateTime = Get-Date -Format 'yyyyMMdd-HHmmss'
            `$existingFolders = Get-ChildItem -Path `$destinationPath -Directory | Where-Object { `$_.Name -match '^Backup_\d{8}-\d{6}$' } | Sort-Object Name -Descending
            if (`$existingFolders.Count -ge `$maxVersions) {
                `$foldersToDelete = `$existingFolders | Select-Object -Skip `$maxVersions
                foreach (`$folder in `$foldersToDelete) {
                    Remove-Item -Path `$folder.FullName -Recurse -Force
                }
            }
            `$newFolderName = 'Backup_' + `$currentDateTime
            `$newFolderPath = Join-Path -Path `$destinationPath -ChildPath `$newFolderName
            New-Item -Path `$newFolderPath -ItemType Directory
            robocopy `$sourcePath `$newFolderPath /MIR /V /LOG+:C:\BackupLog.txt
"@
        }
        # Create folder for the backup tool in %AppData% if not available
        $appDataFolder = [System.IO.Path]::Combine($env:APPDATA, "BackupTool")
        if (-not (Test-Path $appDataFolder)) {
            New-Item -Path $appDataFolder -ItemType Directory | Out-Null
        }

        # Save PowerShell script in .ps1 file in the backup tool folder
        $psScriptPath = [System.IO.Path]::Combine($appDataFolder, "BackupScript.ps1")
        Set-Content -Path $psScriptPath -Value $psScriptContent

        # Configure Task Scheduler to run the script in the backup tool folder
        schtasks /create /tn "BackupJob" /tr "powershell.exe -File `"$psScriptPath`"" /sc minute /mo $interval /ru SYSTEM

        [Windows.Forms.MessageBox]::Show("Backup job created!")
    } catch {
        [Windows.Forms.MessageBox]::Show("Error creating backup job: $_")
    }
})

$form.Controls.Add($button1)

$button2 = New-Object Windows.Forms.Button
$button2.Location = New-Object Drawing.Point(200, 140)
$button2.Size = New-Object Drawing.Size(160, 30)
$button2.Text = "Delete Backup Job"
$button2.Add_Click({
    try {
        # Delete Task Scheduler job
        Unregister-ScheduledTask -TaskName "BackupJob" -Confirm:$false

        # Delete PowerShell script file in the backup tool folder
        $appDataFolder = [System.IO.Path]::Combine($env:APPDATA, "BackupTool")
        $psScriptPath = [System.IO.Path]::Combine($appDataFolder, "BackupScript.ps1")
        Remove-Item -Path $psScriptPath -Force -ErrorAction SilentlyContinue

        # Delete backup tool folder under %AppData%, if present
        if (Test-Path $appDataFolder) {
            Remove-Item -Path $appDataFolder -Force -Recurse -ErrorAction SilentlyContinue
        }

        [Windows.Forms.MessageBox]::Show("Backup job deleted!")
    } catch {
        [Windows.Forms.MessageBox]::Show("Error deleting backup job: $_")
    }
})

$form.Controls.Add($button2)

[Windows.Forms.Application]::Run($form)
