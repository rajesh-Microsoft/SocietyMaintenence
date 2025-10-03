# ===============================================
# Excel to JSON → Azure Blob Updater (Real-time)
# Folder: C:\Users\rrathore\OneDrive - Microsoft\Personal\Society maintenance app
# Using Storage Account Key
# ===============================================

# --- Parameters ---
$excelFolder        = "C:\Users\rrathore\OneDrive - Microsoft\Personal\Society maintenance app"
$excelFilePattern   = "*.xls*"   # watches .xls, .xlsx, .xlsb (and temp variants)
$localJsonFile      = "C:\Users\rrathore\OneDrive - Microsoft\Personal\Society maintenance app\maintenance_data.json"
$logFile            = "C:\Users\rrathore\OneDrive - Microsoft\Personal\Society maintenance app\UpdateExcelJson.log"
$blobStorageAccount = "societymantaintracker"
$blobContainerName  = "files"
$blobName           = "maintenance_data.json"
$storageAccountKey  = "5MWApAVEjQ18nYM8O9OJNsN8PtjAkwq0S0mtidWjkimvii+rcMuUGM5QTfLA8iuHPwZp9J6MIRq/+AStiva3Fg=="   # <-- Provide your key here

# --- Functions ---
function Ensure-Module {
    param(
        [Parameter(Mandatory=$true)][string]$Name
    )
    try {
        if (-not (Get-Module -ListAvailable -Name $Name)) {
            if (-not (Get-PSRepository -Name 'PSGallery' -ErrorAction SilentlyContinue)) {
                Register-PSRepository -Name 'PSGallery' -SourceLocation 'https://www.powershellgallery.com/api/v2' -InstallationPolicy Trusted
            } elseif ((Get-PSRepository -Name 'PSGallery').InstallationPolicy -ne 'Trusted') {
                Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted
            }
            Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        }
        Import-Module -Name $Name -ErrorAction Stop
        return $true
    } catch {
        Log-Message "Failed to install/import module '$Name': $_"
        return $false
    }
}
function Log-Message {
    param([string]$message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "$timestamp - $message"
    Add-Content -Path $logFile -Value $logEntry
    Write-Host $logEntry
}

function Update-JsonBlob {
    param ([string]$excelPath)

    try {
        # Ensure required modules are available
        $okExcel = Ensure-Module -Name 'ImportExcel'
        $okAzAcc = Ensure-Module -Name 'Az.Accounts'
        $okAzSto = Ensure-Module -Name 'Az.Storage'
        if(-not ($okExcel -and $okAzAcc -and $okAzSto)){
            throw "Required PowerShell modules missing (ImportExcel/Az.Accounts/Az.Storage)"
        }

        # Convert Excel to JSON
        if(-not (Test-Path -LiteralPath $excelPath)){
            Log-Message "Skipped: File not found '$excelPath'"
            return
        }
        if([IO.Path]::GetFileName($excelPath) -like "~$*"){
            Log-Message "Skipped temp file '$excelPath'"
            return
        }
        # Read ALL worksheets and build a sheet-name -> rows map
        $sheetInfo = Get-ExcelSheetInfo -Path $excelPath
        if(-not $sheetInfo){ throw "No worksheets found in '$excelPath'" }
        $sheetsObject = @{}
        foreach($si in $sheetInfo){
            $wsName = $si.WorksheetName
            if([string]::IsNullOrWhiteSpace($wsName)) { $wsName = $si.Name }
            if([string]::IsNullOrWhiteSpace($wsName)) { continue }
            try {
                $rows = Import-Excel -Path $excelPath -WorksheetName $wsName
                $sheetsObject[$wsName] = $rows
            } catch {
                Log-Message "Failed to import worksheet '$wsName': $_"
            }
        }
        if($sheetsObject.Keys.Count -eq 0){ throw "No data imported from any worksheet in '$excelPath'" }
        $jsonData = $sheetsObject | ConvertTo-Json -Depth 10

        # Generate hash of current JSON
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($jsonData)
        $currentHash = [System.BitConverter]::ToString((New-Object System.Security.Cryptography.SHA256Managed).ComputeHash($bytes)) -replace "-",""

        # Create Azure storage context using storage account key
        $ctx = New-AzStorageContext -StorageAccountName $blobStorageAccount -StorageAccountKey $storageAccountKey

        # Check if blob exists
        $existingBlob = Get-AzStorageBlob -Container $blobContainerName -Blob $blobName -Context $ctx
        $updateNeeded = $true

        if ($existingBlob) {
            # Download existing JSON temporarily
            $tempFile = New-TemporaryFile
            Get-AzStorageBlobContent -Container $blobContainerName -Blob $blobName -Destination $tempFile.FullName -Context $ctx -Force

            # Compute hash of existing JSON
            $existingJson = Get-Content -Path $tempFile.FullName -Raw
            $bytesOld = [System.Text.Encoding]::UTF8.GetBytes($existingJson)
            $existingHash = [System.BitConverter]::ToString((New-Object System.Security.Cryptography.SHA256Managed).ComputeHash($bytesOld)) -replace "-",""

            if ($existingHash -eq $currentHash) {
                Log-Message "No changes detected. JSON upload skipped."
                $updateNeeded = $false
            }
        }

        if ($updateNeeded) {
            # Save locally
            $jsonData | Out-File -FilePath $localJsonFile -Encoding UTF8

            # Upload to Azure Blob
            Set-AzStorageBlobContent -File $localJsonFile -Container $blobContainerName -Blob $blobName -Context $ctx -Force
            Log-Message "Excel converted to JSON and uploaded/updated in Azure Blob Storage successfully!"
        }
    }
    catch {
        Log-Message "Error: $_"
    }
}

# --- Real-time Monitoring ---
$watcher = New-Object System.IO.FileSystemWatcher
$watcher.Path = $excelFolder
$watcher.Filter = $excelFilePattern
$watcher.IncludeSubdirectories = $false
$watcher.NotifyFilter = [System.IO.NotifyFilters]::LastWrite -bor [System.IO.NotifyFilters]::FileName

# Debounced handler
$script:lastHandled = @{}
$action = {
    try {
        $full = $Event.SourceEventArgs.FullPath
        if([IO.Path]::GetFileName($full) -like "~$*") { return }
        Start-Sleep -Seconds 2  # allow writes to complete
        $now = Get-Date
        $key = $full.ToLowerInvariant()
        if($script:lastHandled.ContainsKey($key)){
            $prev = $script:lastHandled[$key]
            if(($now - $prev).TotalSeconds -lt 2){ return }
        }
        $script:lastHandled[$key] = $now
        Update-JsonBlob -excelPath $full
    } catch {
        Log-Message "Watcher Error: $_"
    }
}

# Register events for various file changes
Register-ObjectEvent $watcher Changed -Action $action | Out-Null
Register-ObjectEvent $watcher Created -Action $action | Out-Null
Register-ObjectEvent $watcher Renamed -Action $action | Out-Null

# Initial upload: pick the most recently modified Excel file
$initialFile = Get-ChildItem -Path $excelFolder -Filter $excelFilePattern -File | Sort-Object LastWriteTime -Descending | Select-Object -First 1
if($initialFile){
    Update-JsonBlob -excelPath $initialFile.FullName
} else {
    Log-Message "No Excel files found for initial upload in '$excelFolder' matching '$excelFilePattern'"
}

# Keep script alive
while ($true) { Start-Sleep -Seconds 60 }
