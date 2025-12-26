# Patient Archive Search and Extract Tool
# Version: 1.3.1
# Description: Search for patient folders in 7z archives and extract to destination server

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Configuration
$archiveServer = "\\VMSCORE\ARCHIVE$"
$destinationServer = "\\variancom\VA_DATA$\filedata\Patients"
$tempExtractPath = "$env:TEMP\PatientExtract"

# Check and add 7z to PATH if needed
function Initialize-7z {
    # Check if 7z is already in PATH
    if (Get-Command 7z -ErrorAction SilentlyContinue) {
        return $true
    }
    
    # Common 7z installation paths
    $possiblePaths = @(
        "C:\Program Files\7-Zip",
        "C:\Program Files (x86)\7-Zip",
        "${env:ProgramFiles}\7-Zip",
        "${env:ProgramFiles(x86)}\7-Zip"
    )
    
    foreach ($path in $possiblePaths) {
        if (Test-Path "$path\7z.exe") {
            $env:Path = "$path;$env:Path"
            if (Get-Command 7z -ErrorAction SilentlyContinue) {
                return $true
            }
        }
    }
    
    [System.Windows.Forms.MessageBox]::Show(
        "7-Zip not found! Please install 7-Zip and add to system PATH.`nDownload from: https://www.7-zip.org/",
        "7-Zip Not Found",
        "OK",
        "Error"
    )
    return $false
}

# Check server availability
function Test-Servers {
    $vmScoreAccess = Test-Path $archiveServer
    $variancomAccess = Test-Path $destinationServer
    
    if (-not $vmScoreAccess) {
        [System.Windows.Forms.MessageBox]::Show(
            "Cannot connect to VMSCORE server!`nPath: $archiveServer",
            "Connection Error",
            "OK",
            "Error"
        )
        return $false
    }
    
    if (-not $variancomAccess) {
        [System.Windows.Forms.MessageBox]::Show(
            "Cannot connect to VARIANCOM server!`nPath: $destinationServer",
            "Connection Error",
            "OK",
            "Error"
        )
        return $false
    }
    
    return $true
}

# Enhanced function to read archive content with error handling
function Read-ArchiveContent {
    param([string]$archivePath)
    
    $maxRetries = 2
    $retryCount = 0
    
    while ($retryCount -lt $maxRetries) {
        try {
            # Method 1: Direct 7z call
            $processInfo = New-Object System.Diagnostics.ProcessStartInfo
            $processInfo.FileName = "7z"
            $processInfo.Arguments = "l `"$archivePath`""
            $processInfo.RedirectStandardOutput = $true
            $processInfo.RedirectStandardError = $true
            $processInfo.UseShellExecute = $false
            $processInfo.CreateNoWindow = $true
            
            $process = New-Object System.Diagnostics.Process
            $process.StartInfo = $processInfo
            $process.Start() | Out-Null
            
            $stdout = $process.StandardOutput.ReadToEnd()
            $stderr = $process.StandardError.ReadToEnd()
            $process.WaitForExit()
            
            $exitCode = $process.ExitCode
            
            if ($exitCode -eq 0) {
                return $stdout
            }
            else {
                $fileName = [System.IO.Path]::GetFileName($archivePath)
                Write-Host "Warning: 7z exited with code $exitCode for archive: $fileName" -ForegroundColor Yellow
                Write-Host "Error output: $stderr" -ForegroundColor Red
            }
        }
        catch {
            Write-Host "Exception reading archive: $_" -ForegroundColor Red
        }
        
        $retryCount++
        if ($retryCount -lt $maxRetries) {
            Start-Sleep -Milliseconds 500
        }
    }
    
    return $null
}

# Function to search for EXACT folder name in archives
function Find-FolderInArchives {
    param([string]$folderName)
    
    $foundArchives = @()
    
    # Get all 7z archives
    try {
        $allArchives = Get-ChildItem -Path $archiveServer -Filter "*.7z" -ErrorAction Stop | 
                      Select-Object -ExpandProperty FullName
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Error accessing archive directory!`n$($_.Exception.Message)",
            "Access Error",
            "OK",
            "Error"
        )
        return $foundArchives
    }
    
    if ($allArchives.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "No 7z archives found in directory!`nPath: $archiveServer",
            "No Archives Found",
            "OK",
            "Warning"
        )
        return $foundArchives
    }
    
    # Create progress form
    $progressForm = New-Object System.Windows.Forms.Form
    $progressForm.Text = "Searching Archives"
    $progressForm.Size = New-Object System.Drawing.Size(450, 170)
    $progressForm.StartPosition = "CenterScreen"
    $progressForm.FormBorderStyle = "FixedDialog"
    $progressForm.MaximizeBox = $false
    
    $progressLabel = New-Object System.Windows.Forms.Label
    $progressLabel.Location = New-Object System.Drawing.Point(20, 20)
    $progressLabel.Size = New-Object System.Drawing.Size(410, 30)
    $progressLabel.Text = "Searching archives..."
    $progressForm.Controls.Add($progressLabel)
    
    $archiveLabel = New-Object System.Windows.Forms.Label
    $archiveLabel.Location = New-Object System.Drawing.Point(20, 50)
    $archiveLabel.Size = New-Object System.Drawing.Size(410, 30)
    $archiveLabel.Text = "Current:"
    $progressForm.Controls.Add($archiveLabel)
    
    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Location = New-Object System.Drawing.Point(20, 90)
    $progressBar.Size = New-Object System.Drawing.Size(410, 30)
    $progressBar.Style = "Continuous"
    $progressForm.Controls.Add($progressBar)
    
    $progressBar.Maximum = $allArchives.Count
    $progressBar.Value = 0
    
    $progressForm.Show()
    $progressForm.Refresh()
    
    # Search each archive
    for ($i = 0; $i -lt $allArchives.Count; $i++) {
        $archive = $allArchives[$i]
        $archiveName = [System.IO.Path]::GetFileName($archive)
        
        $archiveLabel.Text = "Current: $archiveName"
        $progressBar.Value = $i + 1
        $progressForm.Refresh()
        
        # Check if archive is accessible
        if (-not (Test-Path $archive)) {
            Write-Host "Warning: Archive not accessible: $archiveName" -ForegroundColor Yellow
            continue
        }
        
        # Read archive content
        $archiveContent = Read-ArchiveContent -archivePath $archive
        
        if ($archiveContent -eq $null) {
            Write-Host "Warning: Failed to read archive: $archiveName" -ForegroundColor Yellow
            continue
        }
        
        # Parse 7z output to find exact folder name
        # 7z output format: Date Time Attr Size Compressed Name
        # Example: "2023-10-15 14:30:00 D....            0            0 1453/"
        
        $lines = $archiveContent -split "`r`n"
        $folderFound = $false
        
        foreach ($line in $lines) {
            # Look for lines that contain folder indicators (D for directory)
            if ($line -match '^\s*\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2}\s+[D.]') {
                # Extract the folder name from the line (last element)
                $parts = $line -split '\s+'
                
                # The last part should be the folder name
                $currentFolderName = $parts[-1]
                
                # Remove trailing slash if present
                $currentFolderName = $currentFolderName.TrimEnd('/')
                
                # Check for EXACT match (no partial matches)
                if ($currentFolderName -eq $folderName) {
                    $folderFound = $true
                    break
                }
            }
        }
        
        if ($folderFound) {
            $foundArchives += @{
                Path = $archive
                Name = $archiveName
                Folder = $folderName
            }
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
        Write-Host "Extracting $folderName from $archiveName..." -ForegroundColor Cyan
        
        # Execute extraction
        $processInfo = New-Object System.Diagnostics.ProcessStartInfo
        $processInfo.FileName = "7z"
        $processInfo.Arguments = "x `"$archivePath`" `"$folderName/*`" -o`"$tempExtractPath`" -y"
        $processInfo.RedirectStandardOutput = $true
        $processInfo.RedirectStandardError = $true
        $processInfo.UseShellExecute = $false
        $processInfo.CreateNoWindow = $true
        
        $process = New-Object System.Diagnostics.Process
        $process.StartInfo = $processInfo
        $process.Start() | Out-Null
        
        $stdout = $process.StandardOutput.ReadToEnd()
        $stderr = $process.StandardError.ReadToEnd()
        $process.WaitForExit()
        
        if ($process.ExitCode -ne 0) {
            throw "Extraction failed with error: $stderr"
        }
        
        # Check if folder was extracted
        $extractedFolder = Join-Path $tempExtractPath $folderName
        if (-not (Test-Path $extractedFolder)) {
            throw "Folder was not extracted from archive"
        }
        
        # Check if destination folder already exists
        $destinationPath = Join-Path $destinationServer $folderName
        
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
        
        # Copy folder to destination
        Write-Host "Copying $folderName to VARIANCOM server..." -ForegroundColor Cyan
        Copy-Item -Path $extractedFolder -Destination $destinationServer -Recurse -Force
        
        # Clean up temporary folder
        Remove-Item $tempExtractPath -Recurse -Force -ErrorAction SilentlyContinue
        
        Write-Host "Folder $folderName successfully copied to VARIANCOM server." -ForegroundColor Green
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
    $form.Size = New-Object System.Drawing.Size(500, 450)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    
    # Application title
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Location = New-Object System.Drawing.Point(20, 20)
    $titleLabel.Size = New-Object System.Drawing.Size(440, 40)
    $titleLabel.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)
    $titleLabel.Text = "Patient Folder Search and Extract"
    $titleLabel.TextAlign = "MiddleCenter"
    $form.Controls.Add($titleLabel)
    
    # Folder name input field
    $folderLabel = New-Object System.Windows.Forms.Label
    $folderLabel.Location = New-Object System.Drawing.Point(20, 80)
    $folderLabel.Size = New-Object System.Drawing.Size(200, 30)
    $folderLabel.Text = "Enter EXACT folder name:"
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
    $examplesLabel.ForeColor = [System.Drawing.Color]::Blue
    $form.Controls.Add($examplesLabel)
    
    # Important note
    $noteLabel = New-Object System.Windows.Forms.Label
    $noteLabel.Location = New-Object System.Drawing.Point(20, 180)
    $noteLabel.Size = New-Object System.Drawing.Size(440, 30)
    $noteLabel.Text = "Note: Search is EXACT match (e.g., 1453 is not equal to 1453-24)"
    $noteLabel.Font = New-Object System.Drawing.Font("Arial", 8, [System.Drawing.FontStyle]::Italic)
    $noteLabel.ForeColor = [System.Drawing.Color]::DarkRed
    $form.Controls.Add($noteLabel)
    
    # Search button
    $searchButton = New-Object System.Windows.Forms.Button
    $searchButton.Location = New-Object System.Drawing.Point(20, 220)
    $searchButton.Size = New-Object System.Drawing.Size(440, 40)
    $searchButton.Text = "Start Search"
    $searchButton.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
    $searchButton.BackColor = [System.Drawing.Color]::LightBlue
    $form.Controls.Add($searchButton)
    
    # Results display area
    $resultTextBox = New-Object System.Windows.Forms.RichTextBox
    $resultTextBox.Location = New-Object System.Drawing.Point(20, 270)
    $resultTextBox.Size = New-Object System.Drawing.Size(440, 120)
    $resultTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)
    $resultTextBox.ReadOnly = $true
    $resultTextBox.BackColor = [System.Drawing.Color]::WhiteSmoke
    $resultTextBox.ScrollBars = "Vertical"
    $form.Controls.Add($resultTextBox)
    
    # Status label
    $statusLabel = New-Object System.Windows.Forms.Label
    $statusLabel.Location = New-Object System.Drawing.Point(20, 400)
    $statusLabel.Size = New-Object System.Drawing.Size(440, 20)
    $statusLabel.Text = "Ready"
    $statusLabel.Font = New-Object System.Drawing.Font("Arial", 8)
    $statusLabel.ForeColor = [System.Drawing.Color]::DarkGray
    $form.Controls.Add($statusLabel)
    
    # Exit button
    $exitButton = New-Object System.Windows.Forms.Button
    $exitButton.Location = New-Object System.Drawing.Point(20, 420)
    $exitButton.Size = New-Object System.Drawing.Size(440, 30)
    $exitButton.Text = "Exit"
    $exitButton.Font = New-Object System.Drawing.Font("Arial", 9)
    $exitButton.Add_Click({ $form.Close() })
    $form.Controls.Add($exitButton)
    
    # Search button click handler
    $searchButton.Add_Click({
        $folderName = $folderTextBox.Text.Trim()
        
        if ([string]::IsNullOrEmpty($folderName)) {
            [System.Windows.Forms.MessageBox]::Show(
                "Please enter folder name to search!",
                "Input Required",
                "OK",
                "Warning"
            )
            return
        }
        
        $statusLabel.Text = "Searching for EXACT folder name: '$folderName'..."
        $statusLabel.ForeColor = [System.Drawing.Color]::Blue
        $resultTextBox.Text = "Searching for EXACT folder name: '$folderName'...`r`n`r`n"
        $resultTextBox.Refresh()
        $form.Refresh()
        
        # Search for folder in archives
        $foundArchives = Find-FolderInArchives -folderName $folderName
        
        if ($foundArchives.Count -eq 0) {
            $resultTextBox.AppendText("Folder '$folderName' NOT FOUND in any archive!`r`n`r`n")
            $resultTextBox.AppendText("Note: Search was for EXACT match only.`r`n")
            $resultTextBox.AppendText("Check if you entered the correct folder name.`r`n")
            $statusLabel.Text = "Folder not found (exact match only)"
            $statusLabel.ForeColor = [System.Drawing.Color]::Red
            
            [System.Windows.Forms.MessageBox]::Show(
                "Folder '$folderName' not found in any archive!`r`n(Search was for EXACT match only)",
                "Search Result",
                "OK",
                "Information"
            )
        }
        else {
            $resultTextBox.AppendText("Found EXACT match in archive(s):`r`n")
            $resultTextBox.AppendText(("-" * 50) + "`r`n")
            
            foreach ($archive in $foundArchives) {
                $resultTextBox.AppendText("[FOUND] $($archive.Name)`r`n")
            }
            
            $resultTextBox.AppendText("`r`n")
            
            # If found in only one archive, ask for extraction
            if ($foundArchives.Count -eq 1) {
                $archiveToExtract = $foundArchives[0]
                $resultTextBox.AppendText("Selected archive: $($archiveToExtract.Name)`r`n")
                
                $confirm = [System.Windows.Forms.MessageBox]::Show(
                    "EXACT folder match found in archive: $($archiveToExtract.Name)`r`n`r`nExtract and copy to VARIANCOM server?`r`nDestination: $destinationServer",
                    "Confirmation",
                    "YesNo",
                    "Question"
                )
                
                if ($confirm -eq "Yes") {
                    try {
                        $resultTextBox.AppendText("Extracting...`r`n")
                        $resultTextBox.Refresh()
                        $statusLabel.Text = "Extracting folder..."
                        $statusLabel.ForeColor = [System.Drawing.Color]::Blue
                        $form.Refresh()
                        
                        $success = Extract-And-Copy -archivePath $archiveToExtract.Path `
                                                    -folderName $folderName `
                                                    -archiveName $archiveToExtract.Name
                        
                        if ($success) {
                            $resultTextBox.AppendText("[SUCCESS] Folder copied to VARIANCOM server!`r`n")
                            $statusLabel.Text = "Folder copied successfully"
                            $statusLabel.ForeColor = [System.Drawing.Color]::Green
                            
                            [System.Windows.Forms.MessageBox]::Show(
                                "Folder '$folderName' successfully copied to VARIANCOM server!`r`n`r`nPath: $destinationServer\$folderName",
                                "Success",
                                "OK",
                                "Information"
                            )
                        }
                    }
                    catch {
                        $errorMsg = $_.Exception.Message
                        $resultTextBox.AppendText("[ERROR] $errorMsg`r`n")
                        $statusLabel.Text = "Error during extraction"
                        $statusLabel.ForeColor = [System.Drawing.Color]::Red
                        
                        [System.Windows.Forms.MessageBox]::Show(
                            "Error during copy operation:`r`n$errorMsg",
                            "Error",
                            "OK",
                            "Error"
                        )
                    }
                }
                else {
                    $resultTextBox.AppendText("Operation cancelled by user.`r`n")
                    $statusLabel.Text = "Operation cancelled"
                    $statusLabel.ForeColor = [System.Drawing.Color]::Orange
                }
            }
            else {
                # If multiple archives found
                $resultTextBox.AppendText("Folder found in multiple archives.`r`n")
                $resultTextBox.AppendText("Please select specific archive for extraction.`r`n")
                $statusLabel.Text = "Multiple archives found"
                $statusLabel.ForeColor = [System.Drawing.Color]::Orange
                
                # Create simple selection dialog
                $selectionForm = New-Object System.Windows.Forms.Form
                $selectionForm.Text = "Select Archive"
                $selectionForm.Size = New-Object System.Drawing.Size(400, 300)
                $selectionForm.StartPosition = "CenterScreen"
                
                $selectionLabel = New-Object System.Windows.Forms.Label
                $selectionLabel.Location = New-Object System.Drawing.Point(20, 20)
                $selectionLabel.Size = New-Object System.Drawing.Size(350, 40)
                $selectionLabel.Text = "Select archive to extract from:"
                $selectionForm.Controls.Add($selectionLabel)
                
                $listBox = New-Object System.Windows.Forms.ListBox
                $listBox.Location = New-Object System.Drawing.Point(20, 70)
                $listBox.Size = New-Object System.Drawing.Size(350, 150)
                $listBox.SelectionMode = "One"
                
                foreach ($archive in $foundArchives) {
                    $listBox.Items.Add($archive.Name) | Out-Null
                }
                $listBox.SelectedIndex = 0
                $selectionForm.Controls.Add($listBox)
                
                $okButton = New-Object System.Windows.Forms.Button
                $okButton.Location = New-Object System.Drawing.Point(120, 230)
                $okButton.Size = New-Object System.Drawing.Size(80, 30)
                $okButton.Text = "OK"
                $okButton.DialogResult = "OK"
                $selectionForm.Controls.Add($okButton)
                
                $cancelButton = New-Object System.Windows.Forms.Button
                $cancelButton.Location = New-Object System.Drawing.Point(210, 230)
                $cancelButton.Size = New-Object System.Drawing.Size(80, 30)
                $cancelButton.Text = "Cancel"
                $cancelButton.DialogResult = "Cancel"
                $selectionForm.Controls.Add($cancelButton)
                
                $result = $selectionForm.ShowDialog()
                
                if ($result -eq "OK" -and $listBox.SelectedItem) {
                    $selectedArchiveName = $listBox.SelectedItem
                    $selectedArchive = $foundArchives | Where-Object { $_.Name -eq $selectedArchiveName } | Select-Object -First 1
                    
                    if ($selectedArchive) {
                        try {
                            $resultTextBox.AppendText("Extracting from: $($selectedArchive.Name)`r`n")
                            $resultTextBox.Refresh()
                            $statusLabel.Text = "Extracting folder..."
                            $statusLabel.ForeColor = [System.Drawing.Color]::Blue
                            
                            $success = Extract-And-Copy -archivePath $selectedArchive.Path `
                                                        -folderName $folderName `
                                                        -archiveName $selectedArchive.Name
                            
                            if ($success) {
                                $resultTextBox.AppendText("[SUCCESS] Folder copied to VARIANCOM server!`r`n")
                                $statusLabel.Text = "Folder copied successfully"
                                $statusLabel.ForeColor = [System.Drawing.Color]::Green
                                
                                [System.Windows.Forms.MessageBox]::Show(
                                    "Folder '$folderName' successfully copied to VARIANCOM server!",
                                    "Success",
                                    "OK",
                                    "Information"
                                )
                            }
                        }
                        catch {
                            $errorMsg = $_.Exception.Message
                            $resultTextBox.AppendText("[ERROR] $errorMsg`r`n")
                            $statusLabel.Text = "Error during extraction"
                            $statusLabel.ForeColor = [System.Drawing.Color]::Red
                            
                            [System.Windows.Forms.MessageBox]::Show(
                                "Error during copy operation:`r`n$errorMsg",
                                "Error",
                                "OK",
                                "Error"
                            )
                        }
                    }
                }
                else {
                    $resultTextBox.AppendText("Archive selection cancelled.`r`n")
                    $statusLabel.Text = "Selection cancelled"
                    $statusLabel.ForeColor = [System.Drawing.Color]::Orange
                }
            }
        }
        
        $form.Refresh()
    })
    
    # Show form
    return $form.ShowDialog()
}

# Main execution
function Main {
    Write-Host "Patient Archive Search and Extract Tool" -ForegroundColor Cyan
    Write-Host "Version 1.3.1" -ForegroundColor Gray
    Write-Host "Feature: EXACT folder name matching" -ForegroundColor Yellow
    Write-Host ""
    
    # Initialize 7z
    if (-not (Initialize-7z)) {
        return
    }
    
    # Test server connections
    if (-not (Test-Servers)) {
        return
    }
    
    # Show main form
    Show-MainForm | Out-Null
    
    Write-Host "`nApplication closed." -ForegroundColor Gray
}

# Start the application
Main