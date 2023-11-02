Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[Windows.Forms.Application]::EnableVisualStyles()

# Definiere das Verzeichnis für das Backup-Skript
$scriptDirectory = "C:\BackupTool"

# Erstelle das Verzeichnis, falls es nicht existiert
if (-not (Test-Path $scriptDirectory)) {
    New-Item -Path $scriptDirectory -ItemType Directory
}

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
        # Aktuelles Datum im gewünschten Format erstellen (z.B. "yyyyMMdd")
        $currentDate = Get-Date -Format "yyyyMMdd"
        # Zielpfad mit aktuellem Datum erstellen
        $destinationPathWithDate = Join-Path -Path $destinationPath -ChildPath $currentDate

        # Task Scheduler konfigurieren, um den Task alle $interval Minuten zu wiederholen
        schtasks /create /tn "BackupJob" /tr "robocopy `"$sourcePath`" `"$destinationPathWithDate`" /MIR /V /LOG+:C:\BackupLog.txt" /sc minute /mo $interval /ru SYSTEM

        # Überprüfen, ob mehr als $numVersions Versionen vorhanden sind
        $existingVersions = Get-ChildItem -Path $destinationPath | Where-Object { $_.PSIsContainer } | Sort-Object Name -Descending
        if ($existingVersions.Count -gt $numVersions) {
            # Älteste Version löschen, wenn mehr als $numVersions vorhanden sind
            $versionsToDelete = $existingVersions | Select-Object -Skip $numVersions
            foreach ($version in $versionsToDelete) {
                Remove-Item -Path $version.FullName -Recurse -Force
            }
        }

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
        Unregister-ScheduledTask -TaskName "BackupJob" -Confirm:$false
        [Windows.Forms.MessageBox]::Show("Backup job deleted!")
    } catch {
        [Windows.Forms.MessageBox]::Show("Error deleting backup job: $_")
    }
})

$form.Controls.Add($button2)

[Windows.Forms.Application]::Run($form)