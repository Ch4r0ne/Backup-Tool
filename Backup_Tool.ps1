Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[Windows.Forms.Application]::EnableVisualStyles()

$form = New-Object Windows.Forms.Form
$form.Text = "BackupJobSchedulerGUI"
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
$numericUpDown2.Maximum = 10000
$form.Controls.Add($numericUpDown2)

$label3 = New-Object Windows.Forms.Label
$label3.Location = New-Object Drawing.Point(10, 110)
$label3.Size = New-Object Drawing.Size(150, 20)
$label3.Text = "Interval (Minutes):"
$form.Controls.Add($label3)

$numericUpDown = New-Object Windows.Forms.NumericUpDown
$numericUpDown.Location = New-Object Drawing.Point(160, 110)
$numericUpDown.Size = New-Object Drawing.Size(80, 20)
$numericUpDown.Maximum = 10000
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

    # Check whether the user has administrator authorisations
    $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

    if (-not $isAdmin) {
        [Windows.Forms.MessageBox]::Show("You do not have sufficient authorisations to create the backup job. Please run the programme as an administrator.", "Error", [Windows.Forms.MessageBoxButtons]::OK, [Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # Input validation
    if (-not $sourcePath -or -not $destinationPath) {
        [Windows.Forms.MessageBox]::Show("Invalid paths. Please check your entries.", "Error", [Windows.Forms.MessageBoxButtons]::OK, [Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # Check whether Number of Versions and Interval are not 0
    if ($numVersions -eq 0 -or $interval -eq 0) {
        [Windows.Forms.MessageBox]::Show("Invalid Numbers. Please check your entries.", "Error", [Windows.Forms.MessageBoxButtons]::OK, [Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    try {
        # Check if a task with the name "BackupJob" already exists
        $existingTask = Get-ScheduledTask -TaskName "BackupJob" -ErrorAction SilentlyContinue

        # If the task already exists, delete it and wait until it's deleted
        if ($existingTask -ne $null) {
            Unregister-ScheduledTask -TaskName "BackupJob" -Confirm:$false
            # Wait until the task is deleted
            while (Get-ScheduledTask -TaskName "BackupJob" -ErrorAction SilentlyContinue) {
                Start-Sleep -Seconds 1
            }
        }

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

            # Compress files into a zip archive
            `$zipFileName = Join-Path -Path `$destinationPath -ChildPath "`$newFolderName.zip"
            Compress-Archive -Path `$newFolderPath -DestinationPath `$zipFileName -Force

            # Delete folder after creating the zip archive
            Remove-Item -Path `$newFolderPath -Recurse -Force

            # Check and delete older zip archives when the maximum number has been reached
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

        # Open the task scheduling tool after successfully creating the backup job
        Start-Process "taskschd.msc"

        [Windows.Forms.MessageBox]::Show("Backup job successfully created!", "Success", [Windows.Forms.MessageBoxButtons]::OK, [Windows.Forms.MessageBoxIcon]::Information)

    } catch {
        [Windows.Forms.MessageBox]::Show("Error while creating the backup job: $_", "Error", [Windows.Forms.MessageBoxButtons]::OK, [Windows.Forms.MessageBoxIcon]::Error)
    }
})

$form.Controls.Add($button1)

$button2 = New-Object Windows.Forms.Button
$button2.Location = New-Object Drawing.Point(200, 140)
$button2.Size = New-Object Drawing.Size(160, 30)
$button2.Text = "Delete Backup Job"
$button2.Add_Click({
    try {
        # Delete task planner task
        Unregister-ScheduledTask -TaskName "BackupJob" -Confirm:$false

        # Define paths
        $appDataFolder = [System.IO.Path]::Combine($env:APPDATA, "BackupTool")
        $psScriptPath = [System.IO.Path]::Combine($appDataFolder, "BackupScript.ps1")

        # Check whether the file exists before deleting it
        if (Test-Path -Path $psScriptPath -PathType Leaf) {
            Remove-Item -Path $psScriptPath -Force -ErrorAction Stop
        } else {
            [Windows.Forms.MessageBox]::Show("There is no backup job.", "Error", [Windows.Forms.MessageBoxButtons]::OK, [Windows.Forms.MessageBoxIcon]::Error)
            return
        }

        # Check whether the folder exists before deleting it
        if (Test-Path $appDataFolder -PathType Container) {
            Remove-Item -Path $appDataFolder -Force -Recurse -ErrorAction Stop
            [Windows.Forms.MessageBox]::Show("Backup job successfully deleted!", "Success", [Windows.Forms.MessageBoxButtons]::OK, [Windows.Forms.MessageBoxIcon]::Information)
        } else {
            [Windows.Forms.MessageBox]::Show("Backup job not found.", "Error", [Windows.Forms.MessageBoxButtons]::OK, [Windows.Forms.MessageBoxIcon]::Error)
        }
    } catch {
        [Windows.Forms.MessageBox]::Show("Error when deleting the backup job: $_.Exception.Message", "Error", [Windows.Forms.MessageBoxButtons]::OK, [Windows.Forms.MessageBoxIcon]::Error)
    }
})

$form.Controls.Add($button2)

[Windows.Forms.Application]::Run($form)
