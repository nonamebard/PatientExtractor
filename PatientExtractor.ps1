Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Конфигурация путей
$archiveServer = "\\VMSCORE\ARCHIVE$"
$destinationServer = "\\variancom\VA_DATA$\filedata\Patients"
$tempExtractPath = "$env:TEMP\PatientExtract"

# Проверка доступности серверов
function Test-Servers {
    $vmScoreAccess = Test-Path $archiveServer
    $variancomAccess = Test-Path $destinationServer
    
    if (-not $vmScoreAccess) {
        [System.Windows.Forms.MessageBox]::Show("Не удается подключиться к серверу VMSCORE!", "Ошибка", "OK", "Error")
        return $false
    }
    
    if (-not $variancomAccess) {
        [System.Windows.Forms.MessageBox]::Show("Не удается подключиться к серверу VARIANCOM!", "Ошибка", "OK", "Error")
        return $false
    }
    
    # Проверка наличия 7z
    $7zPath = Get-Command 7z -ErrorAction SilentlyContinue
    if (-not $7zPath) {
        [System.Windows.Forms.MessageBox]::Show("7-Zip не найден! Установите 7-Zip и добавьте в PATH.", "Ошибка", "OK", "Error")
        return $false
    }
    
    return $true
}

# Функция поиска папки в архивах
function Find-FolderInArchives {
    param([string]$folderName)
    
    $foundArchives = @()
    $allArchives = Get-ChildItem -Path $archiveServer -Filter "*.7z" | Select-Object -ExpandProperty FullName
    
    $progressForm = New-Object System.Windows.Forms.Form
    $progressForm.Text = "Поиск в архивах"
    $progressForm.Size = New-Object System.Drawing.Size(400, 150)
    $progressForm.StartPosition = "CenterScreen"
    $progressForm.FormBorderStyle = "FixedDialog"
    $progressForm.MaximizeBox = $false
    
    $progressLabel = New-Object System.Windows.Forms.Label
    $progressLabel.Location = New-Object System.Drawing.Point(20, 20)
    $progressLabel.Size = New-Object System.Drawing.Size(350, 30)
    $progressLabel.Text = "Идет поиск в архивах..."
    $progressForm.Controls.Add($progressLabel)
    
    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Location = New-Object System.Drawing.Point(20, 60)
    $progressBar.Size = New-Object System.Drawing.Size(350, 30)
    $progressBar.Style = "Continuous"
    $progressForm.Controls.Add($progressBar)
    
    $progressBar.Maximum = $allArchives.Count
    $progressBar.Value = 0
    
    # Показать форму прогресса
    $progressForm.Show()
    $progressForm.Refresh()
    
    # Поиск в каждом архиве
    for ($i = 0; $i -lt $allArchives.Count; $i++) {
        $archive = $allArchives[$i]
        $archiveName = [System.IO.Path]::GetFileName($archive)
        
        $progressLabel.Text = "Проверка: $archiveName"
        $progressBar.Value = $i + 1
        $progressForm.Refresh()
        
        # Используем 7z для проверки содержимого
        try {
            $archiveContent = & 7z l "$archive" | Out-String
            
            # Ищем папку в выводе 7z
            # Учитываем разные форматы: "7234/", "2342-23/", "2344-2000/"
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
            Write-Warning "Ошибка при чтении архива: $archiveName"
        }
    }
    
    $progressForm.Close()
    return $foundArchives
}

# Функция извлечения и копирования
function Extract-And-Copy {
    param(
        [string]$archivePath,
        [string]$folderName,
        [string]$archiveName
    )
    
    try {
        # Создаем временную папку для извлечения
        if (Test-Path $tempExtractPath) {
            Remove-Item $tempExtractPath -Recurse -Force -ErrorAction SilentlyContinue
        }
        New-Item -ItemType Directory -Path $tempExtractPath -Force | Out-Null
        
        # Извлекаем конкретную папку из архива
        $extractCommand = "7z x `"$archivePath`" `"$folderName/*`" -o`"$tempExtractPath`" -y"
        Write-Host "Выполняется: $extractCommand"
        
        $extractResult = Invoke-Expression $extractCommand
        
        # Проверяем, извлеклась ли папка
        $extractedFolder = Join-Path $tempExtractPath $folderName
        if (-not (Test-Path $extractedFolder)) {
            throw "Папка не была извлечена из архива"
        }
        
        # Копируем в целевое расположение
        $destinationPath = Join-Path $destinationServer $folderName
        
        # Проверяем, существует ли уже папка в назначении
        if (Test-Path $destinationPath) {
            $overwrite = [System.Windows.Forms.MessageBox]::Show(
                "Папка '$folderName' уже существует на сервере VARIANCOM. Перезаписать?",
                "Подтверждение",
                "YesNo",
                "Question"
            )
            
            if ($overwrite -eq "Yes") {
                Remove-Item $destinationPath -Recurse -Force -ErrorAction SilentlyContinue
            }
            else {
                throw "Операция отменена пользователем"
            }
        }
        
        # Копируем папку
        Copy-Item -Path $extractedFolder -Destination $destinationServer -Recurse -Force
        
        # Очищаем временную папку
        Remove-Item $tempExtractPath -Recurse -Force -ErrorAction SilentlyContinue
        
        return $true
    }
    catch {
        # Очистка в случае ошибки
        if (Test-Path $tempExtractPath) {
            Remove-Item $tempExtractPath -Recurse -Force -ErrorAction SilentlyContinue
        }
        throw $_.Exception.Message
    }
}

# Основной графический интерфейс
function Show-MainForm {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Поиск и извлечение папок пациентов"
    $form.Size = New-Object System.Drawing.Size(500, 400)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    
    # Название приложения
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Location = New-Object System.Drawing.Point(20, 20)
    $titleLabel.Size = New-Object System.Drawing.Size(440, 40)
    $titleLabel.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)
    $titleLabel.Text = "Поиск папок пациентов в архивах"
    $titleLabel.TextAlign = "MiddleCenter"
    $form.Controls.Add($titleLabel)
    
    # Поле ввода имени папки
    $folderLabel = New-Object System.Windows.Forms.Label
    $folderLabel.Location = New-Object System.Drawing.Point(20, 80)
    $folderLabel.Size = New-Object System.Drawing.Size(200, 30)
    $folderLabel.Text = "Введите имя папки:"
    $folderLabel.Font = New-Object System.Drawing.Font("Arial", 10)
    $form.Controls.Add($folderLabel)
    
    $folderTextBox = New-Object System.Windows.Forms.TextBox
    $folderTextBox.Location = New-Object System.Drawing.Point(20, 110)
    $folderTextBox.Size = New-Object System.Drawing.Size(440, 30)
    $folderTextBox.Font = New-Object System.Drawing.Font("Arial", 10)
    $folderTextBox.Text = ""
    $form.Controls.Add($folderTextBox)
    
    # Примеры имен папок
    $examplesLabel = New-Object System.Windows.Forms.Label
    $examplesLabel.Location = New-Object System.Drawing.Point(20, 150)
    $examplesLabel.Size = New-Object System.Drawing.Size(440, 30)
    $examplesLabel.Text = "Примеры: 7234, 2342-23, 2344-2000"
    $examplesLabel.Font = New-Object System.Drawing.Font("Arial", 9)
    $examplesLabel.ForeColor = "Blue"
    $form.Controls.Add($examplesLabel)
    
    # Кнопка поиска
    $searchButton = New-Object System.Windows.Forms.Button
    $searchButton.Location = New-Object System.Drawing.Point(20, 190)
    $searchButton.Size = New-Object System.Drawing.Size(440, 40)
    $searchButton.Text = "Начать поиск"
    $searchButton.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
    $searchButton.BackColor = "LightBlue"
    $form.Controls.Add($searchButton)
    
    # Область для вывода результатов
    $resultTextBox = New-Object System.Windows.Forms.RichTextBox
    $resultTextBox.Location = New-Object System.Drawing.Point(20, 240)
    $resultTextBox.Size = New-Object System.Drawing.Size(440, 100)
    $resultTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)
    $resultTextBox.ReadOnly = $true
    $resultTextBox.BackColor = "WhiteSmoke"
    $form.Controls.Add($resultTextBox)
    
    # Список найденных архивов (скрытый)
    $foundArchivesList = @()
    
    # Обработчик кнопки поиска
    $searchButton.Add_Click({
        $folderName = $folderTextBox.Text.Trim()
        
        if ([string]::IsNullOrEmpty($folderName)) {
            [System.Windows.Forms.MessageBox]::Show("Введите имя папки для поиска!", "Ошибка", "OK", "Warning")
            return
        }
        
        $resultTextBox.Text = "Идет поиск папки '$folderName'...`r`n"
        $resultTextBox.Refresh()
        
        # Ищем папку в архивах
        $foundArchivesList = Find-FolderInArchives -folderName $folderName
        
        if ($foundArchivesList.Count -eq 0) {
            $resultTextBox.AppendText("Папка '$folderName' НЕ НАЙДЕНА в архивах!`r`n")
            [System.Windows.Forms.MessageBox]::Show("Папка '$folderName' не найдена в архивах!", "Результат поиска", "OK", "Information")
        }
        else {
            $resultTextBox.AppendText("`r`nНайдено в архивах:`r`n")
            $resultTextBox.AppendText("="*50 + "`r`n")
            
            foreach ($archive in $foundArchivesList) {
                $resultTextBox.AppendText("✓ $($archive.Name)`r`n")
            }
            
            $resultTextBox.AppendText("`r`n")
            
            # Предлагаем выбрать архив для извлечения (если их несколько)
            if ($foundArchivesList.Count -eq 1) {
                $archiveToExtract = $foundArchivesList[0]
                $resultTextBox.AppendText("Извлекаем из: $($archiveToExtract.Name)`r`n")
                
                $confirm = [System.Windows.Forms.MessageBox]::Show(
                    "Найдена папка в архиве: $($archiveToExtract.Name)`r`nИзвлечь и скопировать на сервер VARIANCOM?",
                    "Подтверждение",
                    "YesNo",
                    "Question"
                )
                
                if ($confirm -eq "Yes") {
                    try {
                        $resultTextBox.AppendText("Идет извлечение...`r`n")
                        $resultTextBox.Refresh()
                        
                        $success = Extract-And-Copy -archivePath $archiveToExtract.Path -folderName $folderName -archiveName $archiveToExtract.Name
                        
                        if ($success) {
                            $resultTextBox.AppendText("✓ УСПЕШНО скопировано на сервер VARIANCOM!`r`n")
                            [System.Windows.Forms.MessageBox]::Show(
                                "Папка '$folderName' успешно скопирована на сервер VARIANCOM!`r`nПуть: $destinationServer",
                                "Успех",
                                "OK",
                                "Information"
                            )
                        }
                    }
                    catch {
                        $errorMsg = $_.Exception.Message
                        $resultTextBox.AppendText("✗ ОШИБКА: $errorMsg`r`n")
                        [System.Windows.Forms.MessageBox]::Show(
                            "Ошибка при копировании: $errorMsg",
                            "Ошибка",
                            "OK",
                            "Error"
                        )
                    }
                }
                else {
                    $resultTextBox.AppendText("Операция отменена пользователем.`r`n")
                }
            }
            else {
                # Если несколько архивов - предлагаем выбрать
                $archiveNames = $foundArchivesList | ForEach-Object { $_.Name }
                $selectedArchiveName = [System.Windows.Forms.MessageBox]::Show(
                    "Папка найдена в нескольких архивах. Выберите архив для извлечения:`r`n" + 
                    ($archiveNames -join "`r`n") + 
                    "`r`n`r`nВведите номер архива (например: 2014.7z):",
                    "Выбор архива",
                    "OKCancel",
                    "Question"
                )
                
                # Здесь можно добавить логику для выбора конкретного архива
                $resultTextBox.AppendText("Необходимо выбрать конкретный архив для извлечения.`r`n")
            }
        }
    })
    
    # Кнопка выхода
    $exitButton = New-Object System.Windows.Forms.Button
    $exitButton.Location = New-Object System.Drawing.Point(20, 350)
    $exitButton.Size = New-Object System.Drawing.Size(440, 30)
    $exitButton.Text = "Выход"
    $exitButton.Font = New-Object System.Drawing.Font("Arial", 9)
    $exitButton.Add_Click({ $form.Close() })
    $form.Controls.Add($exitButton)
    
    # Информация о серверах
    $serverInfo = New-Object System.Windows.Forms.Label
    $serverInfo.Location = New-Object System.Drawing.Point(20, 380)
    $serverInfo.Size = New-Object System.Drawing.Size(440, 20)
    $serverInfo.Text = "Архивы: \\VMSCORE\ARCHIVE$  |  Назначение: \\variancom\VA_DATA$\filedata\Patients"
    $serverInfo.Font = New-Object System.Drawing.Font("Arial", 8)
    $serverInfo.ForeColor = "DarkGray"
    $form.Controls.Add($serverInfo)
    
    # Показать форму
    return $form.ShowDialog()
}

# Запуск скрипта
if (Test-Servers) {
    Show-MainForm
}
else {
    [System.Windows.Forms.MessageBox]::Show("Проверьте настройки и попробуйте снова.", "Завершение", "OK", "Information")
}