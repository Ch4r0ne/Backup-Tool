Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[Windows.Forms.Application]::EnableVisualStyles()

$form = New-Object Windows.Forms.Form
$form.Text = "Backup Tool"
$form.Size = New-Object Drawing.Size(400, 250)
$form.StartPosition = "CenterScreen"

$label1 = New-Object Windows.Forms.Label
$label1.Location = New-Object Drawing.Point(10, 20)
$label1.Size = New-Object Drawing.Size(150, 20)
$label1.Text = "Source Path:"
$form.Controls.Add($label1)

$textBox1 = New-Object Windows.Forms.TextBox
$textBox1.Location = New-Object Drawing.Point(160, 20)
$textBox1.Size = New-Object Drawing.Size(200, 20)
$form.Controls.Add($textBox1)

$label2 = New-Object Windows.Forms.Label
$label2.Location = New-Object Drawing.Point(10, 50)
$label2.Size = New-Object Drawing.Size(150, 20)
$label2.Text = "Destination Path:"
$form.Controls.Add($label2)

$textBox2 = New-Object Windows.Forms.TextBox
$textBox2.Location = New-Object Drawing.Point(160, 50)
$textBox2.Size = New-Object Drawing.Size(200, 20)
$form.Controls.Add($textBox2)

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

$button1 = New-Object Windows.Forms.Button
$button1.Location = New-Object Drawing.Point(10, 140)
$button1.Size = New-Object Drawing.Size(180, 30)
$button1.Text = "Create Backup Job"
$button1.Add_Click({
    # Code für das Erstellen des Backup-Jobs hier einfügen
    $sourcePath = $textBox1.Text.Trim()
    $destinationPath = $textBox2.Text.Trim()
    $numVersions = $numericUpDown2.Value
    $interval = $numericUpDown.Value

    # Eingabevalidierung
    if (-not $sourcePath -or -not $destinationPath) {
        [Windows.Forms.MessageBox]::Show("Invalid paths. Please check your inputs.")
        return
    }

    try {
        # PowerShell-Skript in .ps1-Datei im Backup-Tool-Ordner speichern
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
        # Ordner für das Backup-Tool in %AppData% erstellen, falls nicht vorhanden
        $appDataFolder = [System.IO.Path]::Combine($env:APPDATA, "BackupTool")
        if (-not (Test-Path $appDataFolder)) {
            New-Item -Path $appDataFolder -ItemType Directory | Out-Null
        }

        # PowerShell-Skript in .ps1-Datei im Backup-Tool-Ordner speichern
        $psScriptPath = [System.IO.Path]::Combine($appDataFolder, "BackupScript.ps1")
        Set-Content -Path $psScriptPath -Value $psScriptContent

        # Task Scheduler konfigurieren, um das Skript im Backup-Tool-Ordner auszuführen
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
        # Task Scheduler-Job löschen
        Unregister-ScheduledTask -TaskName "BackupJob" -Confirm:$false

        # PowerShell-Skript-Datei im Backup-Tool-Ordner löschen
        $appDataFolder = [System.IO.Path]::Combine($env:APPDATA, "BackupTool")
        $psScriptPath = [System.IO.Path]::Combine($appDataFolder, "BackupScript.ps1")
        Remove-Item -Path $psScriptPath -Force -ErrorAction SilentlyContinue

        # Backup-Tool-Ordner unter %AppData% löschen, falls vorhanden
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
