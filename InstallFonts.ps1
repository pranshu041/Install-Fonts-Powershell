# By: Pranshu Bahadur
# Get Fonts Folder from the same folder as this script.
$fontsFolderDirectory = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
$sourcePath = Join-Path $fontsFolderDirectory "\FontFiles\"
Invoke-Expression ".\$sourcePath"

# Creates a log file for this script.
New-Item -ItemType "file" -Path "C:\Windows\Logs\fontInstall.log" -Force
# Log file path.
$logFile = "C:\Windows\Logs\fontInstall.log"

# Path to the font files folder.
$destinationPath = "C:\Windows\Fonts\"

# Code for setting up Copy Here.
function InstallFont {
    param( $Source, $DstFolder, $CopyType = 0 )

    # Convert the decimal to hex
    $copyFlag = [String]::Format("{0:x}", $CopyType)

    $objShell = New-Object -ComObject "Shell.Application"
    $objFolder = $objShell.NameSpace($destinationPath) 
    $objFolder.CopyHere($Source, $copyFlag)
}

# Function to get the hash code of a font.
function Get-Hash {
    param ($filepath, $Algorithm)
    $algo = New-Object -TypeName ("System.Security.Cryptography.$Algorithm" + "CryptoServiceProvider")
    return [System.BitConverter]::ToString($algo.ComputeHash([System.IO.File]::ReadAllBytes($filePath.fullname)))
}

# Function to get a list of all fonts installed on this pc.
function InstalledFonts {
    [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
    $installedFonts = (New-Object System.Drawing.Text.InstalledFontCollection).Families

    $trimmedExistingFonts = @()
    foreach ($key in $installedFonts) {
        $key = $key.Name
        $key = $key.Replace(' ', '').ToLower()
        $key = $key.Replace('-', '')
        $key = $key.Replace('_', '')
        if ($trimmedExistingFonts.Contains($key)) {
            continue
        }
        $trimmedExistingFonts += $key
    }
    return $trimmedExistingFonts
}


# Code for getting time stamps for the log files.
$time = (Get-Date -f g)

# Map of a Map of information needed for a source font.
$sourceFonts = @{}

# Map of a Map of information needed for a source font.
$destinationFonts = @{}

# Function to map fonts and calculates their different hash keys.
function FontMapper {
    param ($Path, $Map)
    $hashes = @()
    $index = 0
    add-content -Path $logFile -Value "==============================================================================="
    add-content -Path $logFile -Value "[$time] Creating Hashtable for files in $Path"
    foreach ($file in $(Get-ChildItem -Path $path)) {
        $index++
        add-content -Path $logFile -Value "================================== [$index] ============================================="
        $fontValue = @{
            md5HashKey = Get-Hash $file -Algorithm MD5
            filePath   = $file.fullname
            name       = $file.name
            title      = $file.BaseName
        } 
        $key = $file.BaseName
        $key = $key.Replace('-', '').ToLower()
        $key = $key.Replace(' ', '').ToLower()
        $key = $key.Replace('_', '')

        $hashKey = $fontvalue.md5HashKey

        $hashExists = $hashes.Contains($hashKey)
        
        $match = $Map.ContainsKey($key)

        if ($match -or $hashExists) {
            add-content -Path $logFile -Value "[$time] $key already exists in this hashtable" -Force
            add-content -Path $logFile -Value "==============================================================================="
            continue
        }


        $Map.add($key, $fontValue)
        $hashes += $fontValue.md5HashKey
        $title = $Map[$key].title
        add-content -Path $logFile -Value "[$time] $title element created in hash table." -Force
        add-content -Path $logFile -Value "[$time] Values in Element:" -Force
        $hash = $Map[$key].md5HashKey
        $filePath = $Map[$key].filePath
        $name = $Map[$key].name
        
        add-content -Path $logFile -Value "[$time] File Name: $name" -Force
        add-content -Path $logFile -Value "[$time] MD5 HASH: $hash" -Force
        add-content -Path $logFile -Value "[$time] Font name string without special characters: $key" -Force
        add-content -Path $logFile -Value "[$time] File Path: $filePath" -Force
    }
}

# Function to find if a source font hash key and title is equal to a destination font hash key.
function FontExists {
    param ($fontKey)
    $exists = $installedFonts.Contains($fontKey)
    return $exists -or $sourceFonts[$fontKey].md5HashKey -like $destinationFonts[$fontKey].md5HashKey -or (Test-Path -Path ($destinationPath + $sourceFonts[$fontKey].name))
}

# Tests if script ran successfully. Using gdi32 library, temporary add resource font method.
function TestScript {
    add-content -Path $logFile -Value "==============================================================================="
    add-content -Path $logFile -Value "[$time] TESTING PHASE" -Force
    add-content -Path $logFile -Value "==============================================================================="
    add-type -name Session -namespace "" -member @"
[DllImport("gdi32.dll")]
public static extern int AddFontResource(string filePath);
"@
    $failures = @()

    foreach ($font in (Get-ChildItem $sourcePath)) {
        if ([Session]::AddFontResource($font.FullName) -lt 1) {
            $failures.add($font)   
        }
    }

    if ($failures.length -eq 0) {
        Add-Content -Path $logFile -Value "[$time] Test was successful."
    }
    else {
        Add-Content -Path $logFile -Value "[$time] Test failed."
        Add-Content -Path $logFile -Value "[$time] Failures:"
        $index = 0
        foreach ($file in $failures) {
            $index++
            Add-Content -Path $logFile -Value "$index. [$time] $file"

        }
    }
}

$installedFonts = InstalledFonts

# Creates a map of fonts in the source directory.
FontMapper -Path $sourcePath -Map $sourceFonts

# Creates a map of fonts in the destination directory.
FontMapper -Path $destinationPath -Map $destinationFonts -Force

# Index for log file.
$index = 0
add-content -Path $logFile -Value "==============================================================================="
add-content -Path $logFile -Value "[$time] INSTALL PHASE" -Force
add-content -Path $logFile -Value "==============================================================================="

# Loop that compares hash values and installs fonts.
foreach($fontKey in $sourceFonts.Keys) {
    $index++
    $font = $sourceFonts[$fontKey].name
    if (FontExists $fontKey) {
        add-content -Path $logFile -Value "$index. [$time] $font font already exists" -Force
        continue
    }
    InstallFont $sourceFonts[$fontKey].filePath
    add-content -Path $logFile -Value "$index. [$time] |||||||||||||||||||||||| INSTALLED $font |||||||||||||||||||||||||" -Force 
}

TestScript


