Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Configuration paths
$archiveServer = "\\VMSCORE\ARCHIVE$"
$destinationServer = "\\variancom\VA_DATA$\filedata\Patients"
$tempExtractPath = "$env:TEMP\PatientExtract"

# Check server availability
function Test-Servers {
    $vmScoreAccess = Test-Path $archiveServer
    $variancomAccess = Test-Path $destinationServer
    
    if (-not $vmScoreAccess) {
        [System.Windows.Forms.MessageBox]::Show("Cannot connect to VMSCORE server!", "Error", "OK", "Error")
        return $false
    }
    
    if (-not $variancomAccess) {
        [System.Windows.Forms.MessageBox]::Show("Cannot connect to VARIANCOM server!", "Error", "OK", "Error")
        return $false
    }
    
    # Check for 7z installation
    $7zPath = Get-Command 7z -ErrorAction SilentlyContinue
    if (-not $7zPath) {
        [System.Windows.Forms.MessageBox]::Show("7-Zip not found! Please install 7-Zip and add to PATH.", "Error", "OK", "Error")
        return $false
    }
    
    return $true
}

# Function to search for folder in archives
function Find-FolderInArchives {
    param([string]$folderName)
    
    $foundArchives = @()
    $allArchives = Get-ChildItem -Path $archiveServer -Filter "*.7z" | Select-Object -ExpandProperty FullName
    
    $progressForm = New-Object System.Windows.Forms.Form
    $progressForm.Text = "Searching Archives"
    $progressForm.Size = New-Object System.Drawing.Size(400, 150)
    $progressForm.StartPosition = "CenterScreen"
    $progressForm.FormBorderStyle = "FixedDialog"
    $progressForm.MaximizeBox = $false
    
    $progressLabel = New-Object System.Windows.Forms.Label
    $progressLabel.Location = New-Object System.Drawing.Point(20, 20)
    $progressLabel.Size = New-Object System.Drawing.Size(350, 30)
    $progressLabel.Text = "Searching archives in progress..."
    $progressForm.Controls.Add($progressLabel)
    
    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Location = New-Object System.Drawing.Point(20, 60)
    $progressBar.Size = New-Object System.Drawing.Size(350, 30)
    $progressBar.Style = "Continuous"
    $progressForm.Controls.Add($progressBar)
    
    $progressBar.Maximum = $allArchives.Count
    $progressBar.Value = 0
    
    # Show progress form
    $progressForm.Show()
    $progressForm.Refresh()
    
    # Search in each archive
    for ($i = 0; $i -lt $allArchives.Count; $i++) {
        $archive = $allArchives[$i]
        $archiveName = [System.IO.Path]::GetFileName($archive)
        
        $progressLabel.Text = "Checking: $archiveName"
        $progressBar.Value = $i + 1
        $progressForm.Refresh()
        
        # Use 7z to check archive contents
        try {
            $archiveContent = & 7z l "$archive" | Out-String
            
            # Look for folder in 7z output
            # Account for different formats: "7234/", "2342-23/", "2344-2000/"
            $searchPatterns = @("$folderName/", "$folderName\")
            
            foreach ($pattern in $searchPatterns) {
                if ($archiveContent -match $pattern) {
                    $foundArchives += @{
                        Path = $archive
                        Name = $archiveName
                        Folder = $folderName
                    }
                    break
                }
            }
        }
        catch {
            Write-Warning "Error reading archive: $archiveName"
        }
    }
    
    $progressForm.Close()
    return $foundArchives
}

# Function for extraction and copying
function Extract-And-Copy {
    param(
        [string]$archivePath,
        [string]$folderName,
        [string]$archiveName
    )
    
    try {
        # Create temporary extraction folder
        if (Test-Path $tempExtractPath) {
            Remove-Item $tempExtractPath -Recurse -Force -ErrorAction SilentlyContinue
        }
        New-Item -ItemType Directory -Path $tempExtractPath -Force | Out-Null
        
        # Extract specific folder from archive
        $extractCommand = "7z x `"$archivePath`" `"$folderName/*`" -o`"$tempExtractPath`" -y"
        Write-Host "Executing: $extractCommand"
        
        $extractResult = Invoke-Expression $extractCommand
        
        # Check if folder was extracted
        $extractedFolder = Join-Path $tempExtractPath $folderName
        if (-not (Test-Path $extractedFolder)) {
            throw "Folder was not extracted from archive"
        }
        
        # Copy to target location
        $destinationPath = Join-Path $destinationServer $folderName
        
        # Check if folder already exists at destination
        if (Test-Path $destinationPath) {
            $overwrite = [System.Windows.Forms.MessageBox]::Show(
                "Folder '$folderName' already exists on VARIANCOM server. Overwrite?",
                "Confirmation",
                "YesNo",
                "Question"
            )
            
            if ($overwrite -eq "Yes") {
                Remove-Item $destinationPath -Recurse -Force -ErrorAction SilentlyContinue
            }
            else {
                throw "Operation cancelled by user"
            }
        }
        
        # Copy folder
        Copy-Item -Path $extractedFolder -Destination $destinationServer -Recurse -Force
        
        # Clean up temporary folder
        Remove-Item $tempExtractPath -Recurse -Force -ErrorAction SilentlyContinue
        
        return $true
    }
    catch {
        # Cleanup in case of error
        if (Test-Path $tempExtractPath) {
            Remove-Item $tempExtractPath -Recurse -Force -ErrorAction SilentlyContinue
        }
        throw $_.Exception.Message
    }
}

# Main GUI form
function Show-MainForm {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Patient Folder Search and Extract"
    $form.Size = New-Object System.Drawing.Size(500, 400)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    
    # Application title
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Location = New-Object System.Drawing.Point(20, 20)
    $titleLabel.Size = New-Object System.Drawing.Size(440, 40)
    $titleLabel.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)
    $titleLabel.Text = "Search Patient Folders in Archives"
    $titleLabel.TextAlign = "MiddleCenter"
    $form.Controls.Add($titleLabel)
    
    # Folder name input field
    $folderLabel = New-Object System.Windows.Forms.Label
    $folderLabel.Location = New-Object System.Drawing.Point(20, 80)
    $folderLabel.Size = New-Object System.Drawing.Size(200, 30)
    $folderLabel.Text = "Enter folder name:"
    $folderLabel.Font = New-Object System.Drawing.Font("Arial", 10)
    $form.Controls.Add($folderLabel)
    
    $folderTextBox = New-Object System.Windows.Forms.TextBox
    $folderTextBox.Location = New-Object System.Drawing.Point(20, 110)
    $folderTextBox.Size = New-Object System.Drawing.Size(440, 30)
    $folderTextBox.Font = New-Object System.Drawing.Font("Arial", 10)
    $folderTextBox.Text = ""
    $form.Controls.Add($folderTextBox)
    
    # Folder name examples
    $examplesLabel = New-Object System.Windows.Forms.Label
    $examplesLabel.Location = New-Object System.Drawing.Point(20, 150)
    $examplesLabel.Size = New-Object System.Drawing.Size(440, 30)
    $examplesLabel.Text = "Examples: 7234, 2342-23, 2344-2000"
    $examplesLabel.Font = New-Object System.Drawing.Font("Arial", 9)
    $examplesLabel.ForeColor = "Blue"
    $form.Controls.Add($examplesLabel)
    
    # Search button
    $searchButton = New-Object System.Windows.Forms.Button
    $searchButton.Location = New-Object System.Drawing.Point(20, 190)
    $searchButton.Size = New-Object System.Drawing.Size(440, 40)
    $searchButton.Text = "Start Search"
    $searchButton.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
    $searchButton.BackColor = "LightBlue"
    $form.Controls.Add($searchButton)
    
    # Results display area
    $resultTextBox = New-Object System.Windows.Forms.RichTextBox
    $resultTextBox.Location = New-Object System.Drawing.Point(20, 240)
    $resultTextBox.Size = New-Object System.Drawing.Size(440, 100)
    $resultTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)
    $resultTextBox.ReadOnly = $true
    $resultTextBox.BackColor = "WhiteSmoke"
    $form.Controls.Add($resultTextBox)
    
    # List of found archives (hidden)
    $foundArchivesList = @()
    
    # Search button click handler
    $searchButton.Add_Click({
        $folderName = $folderTextBox.Text.Trim()
        
        if ([string]::IsNullOrEmpty($folderName)) {
            [System.Windows.Forms.MessageBox]::Show("Please enter folder name to search!", "Error", "OK", "Warning")
            return
        }
        
        $resultTextBox.Text = "Searching for folder '$folderName'...`r`n"
        $resultTextBox.Refresh()
        
        # Search for folder in archives
        $foundArchivesList = Find-FolderInArchives -folderName $folderName
        
        if ($foundArchivesList.Count -eq 0) {
            $resultTextBox.AppendText("Folder '$folderName' NOT FOUND in archives!`r`n")
            [System.Windows.Forms.MessageBox]::Show("Folder '$folderName' not found in archives!", "Search Result", "OK", "Information")
        }
        else {
            $resultTextBox.AppendText("`r`nFound in archives:`r`n")
            $resultTextBox.AppendText("="*50 + "`r`n")
            
            foreach ($archive in $foundArchivesList) {
                $resultTextBox.AppendText("OK $($archive.Name)`r`n")
            }
            
            $resultTextBox.AppendText("`r`n")
            
            # Offer to select archive for extraction (if multiple found)
            if ($foundArchivesList.Count -eq 1) {
                $archiveToExtract = $foundArchivesList[0]
                $resultTextBox.AppendText("Extracting from: $($archiveToExtract.Name)`r`n")
                
                $confirm = [System.Windows.Forms.MessageBox]::Show(
                    "Folder found in archive: $($archiveToExtract.Name)`r`nExtract and copy to VARIANCOM server?",
                    "Confirmation",
                    "YesNo",
                    "Question"
                )
                
                if ($confirm -eq "Yes") {
                    try {
                        $resultTextBox.AppendText("Extraction in progress...`r`n")
                        $resultTextBox.Refresh()
                        
                        $success = Extract-And-Copy -archivePath $archiveToExtract.Path -folderName $folderName -archiveName $archiveToExtract.Name
                        
                        if ($success) {
                            $resultTextBox.AppendText("OK SUCCESSFULLY copied to VARIANCOM server!`r`n")
                            [System.Windows.Forms.MessageBox]::Show(
                                "Folder '$folderName' successfully copied to VARIANCOM server!`r`nPath: $destinationServer",
                                "Success",
                                "OK",
                                "Information"
                            )
                        }
                    }
                    catch {
                        $errorMsg = $_.Exception.Message
                        $resultTextBox.AppendText("✗ ERROR: $errorMsg`r`n")
                        [System.Windows.Forms.MessageBox]::Show(
                            "Error during copy: $errorMsg",
                            "Error",
                            "OK",
                            "Error"
                        )
                    }
                }
                else {
                    $resultTextBox.AppendText("Operation cancelled by user.`r`n")
                }
            }
            else {
                # If multiple archives found - offer selection
                $archiveNames = $foundArchivesList | ForEach-Object { $_.Name }
                $selectedArchiveName = [System.Windows.Forms.MessageBox]::Show(
                    "Folder found in multiple archives. Select archive for extraction:`r`n" + 
                    ($archiveNames -join "`r`n") + 
                    "`r`n`r`nEnter archive name (for example: 2014.7z):",
                    "Archive Selection",
                    "OKCancel",
                    "Question"
                )
                
                # Here you can add logic for selecting specific archive
                $resultTextBox.AppendText("Please select specific archive for extraction.`r`n")
            }
        }
    })
    
    # Exit button
    $exitButton = New-Object System.Windows.Forms.Button
    $exitButton.Location = New-Object System.Drawing.Point(20, 350)
    $exitButton.Size = New-Object System.Drawing.Size(440, 30)
    $exitButton.Text = "Exit"
    $exitButton.Font = New-Object System.Drawing.Font("Arial", 9)
    $exitButton.Add_Click({ $form.Close() })
    $form.Controls.Add($exitButton)
    
    # Server information
    $serverInfo = New-Object System.Windows.Forms.Label
    $serverInfo.Location = New-Object System.Drawing.Point(20, 380)
    $serverInfo.Size = New-Object System.Drawing.Size(440, 20)
    $serverInfo.Text = "Archives: \\VMSCORE\ARCHIVE$  |  Destination: \\variancom\VA_DATA$\filedata\Patients"
    $serverInfo.Font = New-Object System.Drawing.Font("Arial", 8)
    $serverInfo.ForeColor = "DarkGray"
    $form.Controls.Add($serverInfo)
    
    # Show form
    return $form.ShowDialog()
}

# Run script
if (Test-Servers) {
    Show-MainForm
}
else {
    [System.Windows.Forms.MessageBox]::Show("Check settings and try again.", "Exit", "OK", "Information")
}