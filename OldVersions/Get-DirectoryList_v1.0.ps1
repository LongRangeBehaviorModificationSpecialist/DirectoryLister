<#
.SYNOPSIS
    Recursively lists files and directories with their attributes to a CSV,
    with a minimal footprint on the host, and documents the run in a manifest.

.DESCRIPTION
    Walks a directory tree using read-only .NET directory enumeration and
    streams one CSV row per file or directory straight to the output file.
    File contents are never opened, no temp files are created, no compiled
    code is loaded, and the output is never written inside the folder being
    examined.

    Files produced (all in the output directory):
        `<RunDate>_<ComputerName>_<VolumeSerialNumber>_output.csv`
                CSV listing
        <RunDate>_<ComputerName>_<VolumeSerialNumber>_manifest.txt
                Run manifest: host/user/time/offset, parameters, script
                hash, volume details, counts, SHA-256 of the CSV, plus any
                read errors and deliberately skipped directories.

    -Include :  Only directories matching these names/paths are documented
                recursively).
                Non-matching directories are still walked so that matching
                directories nested deeper can be found, but their own files
                are not recorded.
    -Exclude :  Directories matching these names/paths are pruned. They are
                never entered, and nothing beneath them is read or recorded.

    Exclude wins over Include.

    Patterns are case-insensitive and may use wildcards (* ? [ab]). A pattern
    is compared to both the directory's name (e.g. "Windows") and its full
    path (e.g. "C:\Users\Bob\Temp").

    Reparse points (junctions/symlinks) are listed with their target in the
    ReparseTarget column but are never followed. Reading a link's target
    opens the link itself (not what it points to); cloud placeholder files
    are not hydrated.

.PARAMETER Folder
    Root folder to examine.

.PARAMETER CsvDir
    Output directory to save CSV path. The run timestamp is prefixed to the
    file name (which is automatically generated).

    The output directory must already exist and should be on separate
    (external) media, NOT the machine being examined.

.PARAMETER IncludeVolumeSerial
    Also records the volume serial number in the manifest.

    This uses a CIM/WMI query, which leaves more host-side traces than the
    rest of the script, so it is off by default.

.EXAMPLE
    .\Get-DirectoryList.ps1 `
        -Folder C:\ `
        -CsvDir E:\case1 `
        -Exclude Windows,'Program Files*'

.EXAMPLE
    .\Get-DirectoryList.ps1 `
        -Folder C:\Users `
        -CsvDir E:\case1 `
        -Include Documents,Desktop

.EXAMPLE
    .\Get-DirectoryList.ps1 `
        -DIR C:\ `
        -CSV E:\Case1\docs.csv

.NOTES
    Suggested launch (no profile, no prompts, execution policy bypass for
    this process only):

    powershell.exe `
        -NoProfile `
        -NonInteractive `
        -ExecutionPolicy Bypass `
        -File .\Get-DirectoryList.ps1
#>
[CmdletBinding()]
param(
    [string]$RunDate = [DateTime]::UtcNow.ToString("yyyy.MM.dd_HH.mm.ss"),

    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [Alias("DIR")]
    [string]$Folder,

    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [Alias("CSV")]
    [string]$CsvDir,

    [string[]]$Include,

    [string[]]$Exclude,

    [Alias("VSN")]
    [switch]$IncludeVolumeSerial
)

$ScriptVersion         = "1.0"
$StartTime             = [DateTime]::UtcNow
$Stopwatch             = [System.Diagnostics.Stopwatch]::StartNew()
$HostName              = [System.Environment]::MachineName
$Sep                   = [System.IO.Path]::DirectorySeparatorChar
$Inv                   = [System.Globalization.CultureInfo]::InvariantCulture
$ErrorActionPreference = [System.Management.Automation.ActionPreference]::Continue


function Resolve-FullPath {
    # Resolves a path against the PowerShell location (not the .NET
    # process directory) without touching the disk and without
    # expanding wildcards.
    param(
        [string]$Path
    )
    [System.IO.Path]::GetFullPath($ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path))
}

function New-DirPatterns {
    param(
        [string[]]$Names
    )
    $Opts = [System.Management.Automation.WildcardOptions]::IgnoreCase -bor
            [System.Management.Automation.WildcardOptions]::CultureInvariant
    $List = [System.Collections.Generic.List[System.Management.Automation.WildcardPattern]]::New()
    foreach ($Name in $Names) {
        if ([string]::IsNullOrWhiteSpace($Name)) {
            continue
        }
        $List.Add([System.Management.Automation.WildcardPattern]::New($Name.Trim().TrimEnd("\", "/"), $Opts))
    }
    # The leading comma stops PowerShell from unrolling the list
    , $List
}

function Test-DirMatch {
    param(
        [System.IO.DirectoryInfo]$Dir,
        $Patterns
    )
    foreach ($Pattern in $Patterns) {
        if ($Pattern.IsMatch($Dir.Name) -or $Pattern.IsMatch($Dir.FullName)) {
            return $True
        }
    }
    return $False
}

function Get-FormattedFileSize {
    <#
    .SYNOPSIS
        Returns a size in bytes as readable text, showing the rounded size
        and the exact byte count.
    .DESCRIPTION

        Chooses the largest unit (TB, GB, MB or KB) that fits, using 1024-
        based units like Windows File Explorer. The rounded value is shown
        with two decimal places, followed by the exact byte count with
        thousands separators. Sizes under 1 KB are shown as plain bytes.

        The size can be supplied directly with -Bytes or looked up from a
        file with -FilePath.  The invariant culture is always used so the
        text is the same on every computer.

        Example result: 1.50 MB (1,572,864 bytes)

        A value that would round up to 1,024.00 of a unit (for example
        1,048,575 bytes) is shown in the next larger unit instead.

    .PARAMETER FilePath
        Path to the file whose size should be formatted. The path is used
        literally (wildcard characters such as [ ] are not expanded) and
        must point to an existing file, otherwise a terminating error is
        thrown.

    .PARAMETER Bytes
        The size in bytes.

        Use this when the size is already known (for example $File.Length)
        to avoid a second file lookup.

    .EXAMPLE
        Get-FormattedFileSize -FilePath "C:\Evidence\memory.raw"

        Returns a string such as "8.00 GB (8,589,934,592 bytes)".

    .EXAMPLE
        Get-FormattedFileSize -Bytes $File.Length

        Formats a size that is already available.

    .EXAMPLE
        Get-FormattedFileSize -Bytes 512

        Returns "512 bytes".

    .INPUTS
        None.  This function does not accept pipeline input.

    .OUTPUTS
        System.String.  The formatted size.

    .NOTES
        Units are 1024-based (KB = 1,024 bytes), matching Windows File Explorer.
    #>
    [CmdletBinding(DefaultParameterSetName = "ByPath")]
    [OutputType([string])]
    param (
        [Parameter(Mandatory, Position = 0, ParameterSetName = "ByPath")]
        [ValidateNotNullOrEmpty()]
        [string]$FilePath,

        [Parameter(Mandatory, ParameterSetName = "ByBytes")]
        [ValidateRange(0, [long]::MaxValue)]
        [long]$Bytes
    )

    if ($PSCmdlet.ParameterSetName -eq "ByPath") {
        if (-not (Test-Path -LiteralPath $FilePath -PathType Leaf)) {
            throw "File not found => $FilePath."
        }
        $Bytes = (Get-Item -LiteralPath $FilePath).Length
    }

    $Inv = [System.Globalization.CultureInfo]::InvariantCulture

    if ($Bytes -lt 1KB) {
        return [string]::Format($Inv, "{0:N0} bytes", $Bytes)
    }

    # Divide by 1024 until the rounded value is below 1,024 (or the largest unit is reached)
    $Units = "KB", "MB", "GB", "TB"
    $Size  = [double]$Bytes
    $Index = -1
    do {
        $Size /= 1024
        $Index++
    } while ([math]::Round($Size, 2) -ge 1024 -and $Index -lt ($Units.Count - 1))

    return [string]::Format($Inv, "{0:N2} {1} ({2:N0} bytes)", $Size, $Units[$Index], $Bytes)
}

function Get-ReparseTarget {
    # Returns the link target for symlinks/junctions, or "" for anything else.
    param(
        [System.IO.FileSystemInfo]$Item
    )
    if (-not $Item.Attributes.HasFlag([System.IO.FileAttributes]::ReparsePoint)) {
        return ""
    }
    try {
        return (@($Item.Target) -join "; ")
    }
    catch {
        return ""
    }
}

function Format-Duration {
    <#
    .SYNOPSIS
        Converts a TimeSpan into short, readable text for console
        and log messages.

    .DESCRIPTION
        Output depends on the length of the span (invariant culture
        is always used):

            under 1 minute   "12.34 seconds"
            under 1 hour     "5 min 07 sec"
            1 hour or more   "2 hr 05 min 09 sec"

    .PARAMETER Span
        The elapsed time to format, for example the Elapsed value
        of a Stopwatch.

    .EXAMPLE
        Format-TriageDuration -Span (New-TimeSpan -Seconds 75)

        Returns "1 min 15 sec".

    .INPUTS
        None.

    .OUTPUTS
        System.String.
    #>
    [CmdletBinding()]
    [OutputType([System.String])]
    param(
        [Parameter(Mandatory)]
        [TimeSpan]$Span
    )

    if ($Span.TotalSeconds -lt 60) {
        return [string]::Format($Inv, "{0:N6} seconds", $Span.TotalSeconds)
    }

    if ($Span.TotalHours -lt 1) {
        return [string]::Format($Inv, "{0} min {1:00} sec", $Span.Minutes, $Span.Seconds)
    }

    [string]::Format(
        $Inv, "{0} hr {1:00} min {2:00} sec", `
        [int][math]::Floor($Span.TotalHours), $Span.Minutes, $Span.Seconds
    )
}


function Get-FileSha256 {
    param(
        [string]$Path
    )
    $Fs = [System.IO.FileStream]::New(
        $Path,
        [System.IO.FileMode]::Open,
        [System.IO.FileAccess]::Read,
        [System.IO.FileShare]::ReadWrite
    )
    try {
        $Sha = [System.Security.Cryptography.SHA256]::Create()
        try {
            [System.BitConverter]::ToString($Sha.ComputeHash($Fs)).Replace("-", "")
        }
        finally {
            $Sha.Dispose()
        }
    }
    finally {
        $Fs.Dispose()
    }
}

function Write-Row {
    # Positional parameters (faster than named ones on very large trees):
    # Type ("File"|"Directory"), containing-directory path, the item,
    # reparse target ("" if none)
    param(
        [string]$Type,
        [string]$DirName,
        [System.IO.FileSystemInfo]$Item,
        [string]$Target
    )

    if ($Type -eq "File") {
        $Len   = $Item.Length
        $Ext   = $Item.Extension
        $Bytes = $Len.ToString($Inv)
        $Kb    = ([math]::Round($Len / 1KB, 2)).ToString($Inv)
    }
    else {
        $Ext = ""
        $Bytes = ""
        $Kb = ""
    }

    $Writer.WriteLine(('"{0}","{1}","{2}","{3}","{4}","{5}","{6}","{7}","{8}","{9}","{10}","{11}"' -f
        $DirName.Replace('"', '""'),
        $Item.Name.Replace('"', '""'),
        $Type,
        $Ext.Replace('"', '""'),
        $Bytes,
        $Kb,
        $Item.Mode,
        $Item.Attributes.ToString(),
        $Item.CreationTimeUtc.ToString("o"),
        $Item.LastAccessTimeUtc.ToString("o"),
        $Item.LastWriteTimeUtc.ToString("o"),
        $Target.Replace('"', '""')))
}


# ----------
# Validate input and output locations
# ----------

$RootPath = Resolve-FullPath $Folder

if (-not [System.IO.Directory]::Exists($RootPath)) {
    Write-Warning "Folder not found => $Folder"
    exit 1
}

# Strip a trailing separator unless the path is a drive/share root.
if ($RootPath.Length -gt [System.IO.Path]::GetPathRoot($RootPath).Length) {
    $RootPath = $RootPath.TrimEnd($Sep)
}

$CsvDir = Resolve-FullPath $CsvDir

if ([string]::IsNullOrEmpty($CsvDir)) {

    Write-Error "OutFile must include a directory path, e.g. E:\case1"
    exit 1
}
if (-not [System.IO.Directory]::Exists($CsvDir)) {
    # Deliberately not created
    # Creating directories is an extra change to the file system
    Write-Error "Output directory does not exist => $CsvDir"
    exit 1
}

# Volume serial number (opt-in: needs a CIM/WMI query).  It is looked up here, before the output
# path is built, because it becomes part of the output file name.
$VolRoot      = [System.IO.Path]::GetPathRoot($RootPath)  # Drive name (i.e., `F:\`)
$Device       = $VolRoot.TrimEnd($Sep)  # (i.e., `F:`)
$DriveLetter  = $VolRoot.TrimEnd(":\")  # (i.e., `F`)
$VolumeSerial = $null
$SerialNote   = $null

# if ($IncludeVolumeSerial) {
if ($Device -match "^[A-Za-z]:$") {
    try {
        $Ld = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='$Device'" `
            -Property VolumeSerialNumber -ErrorAction Stop
        # Keep only letters/digits so the value is always safe in a file name.
        if ($Ld) {
            $VolumeSerial = ([string]$Ld.VolumeSerialNumber) -replace "[^0-9A-Za-z]", ""
        }
    }
    catch {
        $SerialNote = "Unavailable => $( $_.Exception.Message )"
    }
}
else {
    $SerialNote = "Unavailable (not a drive-letter volume)"
}

if (-not $VolumeSerial) {
    if (-not $SerialNote) {
        $SerialNote = "Unavailable"
    }
    Write-Warning ("Volume serial number lookup failed ($SerialNote). " +
        "`NOSERIAL` will be used in the file name.")
    $VolumeSerial = "NOSERIAL"
}
# }

$SerialPart  = if ($VolumeSerial) {
    "$( $VolumeSerial )_"
}
else {
    ""
}

# <RunDate>_<HostName>_<Serial>_results.csv
$OutFilePath = [System.IO.Path]::Combine(
    $CsvDir,
    "$( $RunDate )_$( $HostName )_$( $SerialPart )results.csv"
)


# Never write into the tree being examined
# Would alter the target and pollute the listing
$RootPrefix = $RootPath.TrimEnd($Sep) + $Sep
if ($OutFilePath.StartsWith($RootPrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
    Write-Error ("Output file would be written inside the folder being examined. " +
        "Choose an output location outside $RootPath (ideally external media).")
    exit 1
}

if ([string]::Equals(
        [System.IO.Path]::GetPathRoot($RootPath),
        [System.IO.Path]::GetPathRoot($OutFilePath),
        [System.StringComparison]::OrdinalIgnoreCase
        )
    ) {
    Write-Warning ("Output is on the same volume as the folder being examined. " +
        "For forensic soundness, write to separate media.")
}

$IncludePatterns = New-DirPatterns $Include
$ExcludePatterns = New-DirPatterns $Exclude
$UseInclude      = $IncludePatterns.Count -gt 0
$UseExclude      = $ExcludePatterns.Count -gt 0


# ----------
# Volume details
# Metadata only -- snapshot taken before the walk starts
# ----------

$Volume  = [ordered]@{
    "VolumeRoot" = $VolRoot
}

try {
    $Drive = [System.IO.DriveInfo]::New($VolRoot)
    $Volume["DriveType"] = $Drive.DriveType.ToString()

    if ($Drive.IsReady) {
        $Volume["DriveFormat"]         = $Drive.DriveFormat
        $Volume["DrivePhysicalSN"]     = (Get-Partition -DriveLetter $DriveLetter |
            Get-Disk).SerialNumber.Trim(" ")
        $Volume["VolumeLabel"]         = $Drive.VolumeLabel
        $Volume["Volume SN"]           = $VolumeSerial
        $Volume["DriveTotalSizeBytes"] = [string]::Format($Inv, "{0:N0}", $Drive.TotalSize)
        $Volume["DriveFreeBytes"]      = [string]::Format($Inv, "{0:N0}", $Drive.TotalFreeSpace)
        $Volume["DriveStatus"]         = $Drive.IsReady
    }
    else {
        $Volume["DriveStatus"] = "Drive not ready"
    }
}
catch {
    $Volume["DriveStatus"] = "DriveInfo unavailable (UNC or unsupported path) => $( $_.Exception.Message )"
}

if ($IncludeVolumeSerial) {
    $Volume["SerialNumber"] = if ($SerialNote) {
        $SerialNote
    }
    else {
        $VolumeSerial
    }
}


# ----------
# Walk the tree and stream rows to the CSV
# ----------

$Errors    = [System.Collections.Generic.List[string]]::New()
$Skipped   = [System.Collections.Generic.List[string]]::New()
$FileCount = [int]0
$DirRows   = [int]0
$DirCount  = [int]0
# Sum of the logical size of every recorded file
$TotalBytes = [long]0

# Each stack item = (directory, "this directory and its files should be recorded")
$Stack    = [System.Collections.Generic.Stack[System.Tuple[System.IO.DirectoryInfo, bool]]]::New()
$RootDir  = [System.IO.DirectoryInfo]::New($RootPath)
$RootCapt = (-not $UseInclude) -or (Test-DirMatch $RootDir $IncludePatterns)
$Stack.Push([System.Tuple[System.IO.DirectoryInfo, bool]]::New($RootDir, [bool]$RootCapt))

$Stream = $null
$Writer = $null
try {
    # CreateNew: refuses to overwrite an existing file. FileShare.Read: others may read it.
    $Stream = [System.IO.FileStream]::New(
        $OutFilePath,
        [System.IO.FileMode]::CreateNew,
        [System.IO.FileAccess]::Write,
        [System.IO.FileShare]::Read
    )
    # Unicode (UTF-16 LE with BOM) matches the original Export-Csv -Encoding Unicode output.
    $Writer = [System.IO.StreamWriter]::New(
        $Stream,
        [System.Text.Encoding]::Unicode
    )

    $Writer.WriteLine('"DirectoryName","Name","Type","Extension","SizeInBytes","SizeInKB","Mode","Attributes","CreationTimeUTC","LastAccessTimeUTC","LastWriteTimeUTC","ReparseTarget"')

    while ($Stack.Count -gt 0) {
        $Item      = $Stack.Pop()
        $Dir       = $Item.Item1
        $Capturing = $Item.Item2
        $DirCount++

        # --- the directory's own row (before its contents) ---
        if ($Capturing) {
            try {
                $ParentPath = if ($Dir.Parent) {
                    $Dir.Parent.FullName
                }
                else {
                    ""
                }
                Write-Row "Directory" $ParentPath $Dir (Get-ReparseTarget $Dir)
                $DirRows++
            }
            catch {
                $Errors.Add("$( $Dir.FullName )`t$( $_.Exception.Message )")
                Write-Warning "Could not read $( $Dir.FullName ) => $( $_.Exception.Message )"
            }
        }

        $Files   = [System.Collections.Generic.List[System.IO.FileInfo]]::New()
        $SubDirs = [System.Collections.Generic.List[System.IO.DirectoryInfo]]::New()

        try {
            foreach ($Entry in $Dir.EnumerateFileSystemInfos()) {
                if ($Entry -is [System.IO.DirectoryInfo]) {
                    $SubDirs.Add($Entry)
                }
                else {
                    $Files.Add($Entry)
                }
            }
        }
        catch {
            $Errors.Add("$( $Dir.FullName )`t$( $_.Exception.Message )")
            Write-Warning "Could not fully read $( $Dir.FullName ) => $( $_.Exception.Message )"
        }

        # --- files in this directory (only if this branch is being captured) ---
        if ($Capturing -and $Files.Count -gt 0) {
            foreach ($File in ($Files | Sort-Object -Property Name)) {
                try {
                    Write-Row "File" $Dir.FullName $File (Get-ReparseTarget $File)
                    $FileCount++
                    $TotalBytes +=$File.Length
                }
                catch {
                    $Errors.Add("$( $File.FullName )`t$( $_.Exception.Message )")
                    Write-Warning "Could not read $( $File.FullName ) => $( $_.Exception.Message )"
                }
            }
        }

        # --- queue subdirectories (pushed in reverse so they pop in alphabetical order) ---
        foreach ($Sd in ($SubDirs | Sort-Object -Property Name -Descending)) {
            if ($UseExclude -and (Test-DirMatch $Sd $ExcludePatterns)) {
                $Skipped.Add("EXCLUDED`t$( $Sd.FullName )")
                continue
            }

            $ChildCapt = $Capturing -or ($UseInclude -and (Test-DirMatch $Sd $IncludePatterns))

            if ($Sd.Attributes.HasFlag([System.IO.FileAttributes]::ReparsePoint)) {
                # Never followed. Recorded as a row (with its target) if it is in capture scope.
                $Target = Get-ReparseTarget $Sd
                $Skipped.Add("REPARSEPOINT (not followed)`t$( $Sd.FullName )`t$Target")
                if ($ChildCapt) {
                    try {
                        Write-Row "Directory" $Dir.FullName $Sd $Target
                        $DirRows++
                    }
                    catch {
                        $Errors.Add("$( $Sd.FullName )`t$( $_.Exception.Message )")
                        Write-Warning "Could not read $( $Sd.FullName ) => $( $_.Exception.Message )"
                    }
                }
                continue
            }

            $Stack.Push([System.Tuple[System.IO.DirectoryInfo, bool]]::New($Sd, [bool]$ChildCapt))
        }
    }
}
finally {
    if ($Writer) {
        # Also closes the underlying stream
        $Writer.Dispose()
    }
    elseif ($Stream) {
        $Stream.Dispose()
    }
}


# ----------
# Hash the finished CSV and write the manifest
# ----------

$CsvHash = try {
    Get-FileSha256 $OutFilePath
}
catch {
    "Unavailable => $( $_.Exception.Message )"
}

$CsvSize = try {
    [System.IO.FileInfo]::New($OutFilePath).Length
}
catch {
    "Unavailable"
}

$ScriptHash = "N/A (not run from a script file)"
if ($PSCommandPath) {
    $ScriptHash = try {
        Get-FileSha256 $PSCommandPath
    }
    catch {
        "Unavailable => $( $_.Exception.Message )"
    }
}

$Elevated = try {
    (
        [System.Security.Principal.WindowsPrincipal]`
        [System.Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
        [System.Security.Principal.WindowsBuiltInRole]::Administrator
    )
}
catch {
    "Unknown"
}


$EndTime = [DateTime]::UtcNow
$Stopwatch.Stop()
$ElapsedTime = Format-Duration -Span $stopwatch.Elapsed

$Offset = [System.TimeZoneInfo]::Local.GetUtcOffset([DateTime]::UtcNow)

$OffsetText = [string]::Format("UTC{0}{1}",
    $(if ($Offset -lt [TimeSpan]::Zero) {
        "-"
    }
    else {
        "+"
    }),
    $Offset.Duration().ToString("hh\:mm")
)

$M = [System.Collections.Generic.List[string]]::New()

$M.Add("===== RUN =====")

$M.Add(
    (
        [string]::Format(
            "{0,-21}: {1} v{2}",
            "Script",
            $MyInvocation.MyCommand.Name,
            $ScriptVersion
        )
    )
)
$M.Add(
    (
        [string]::Format(
            "{0,-21}: {1}",
            "Script SHA-256",
            $ScriptHash
        )
    )
)
$M.Add(
    (
        [string]::Format(
            "{0,-21}: {1}",
            "Start Time (Utc)",
            $StartTime.ToString("o")
        )
    )
)
$M.Add(
    (
        [string]::Format(
            "{0,-21}: {1}",
            "End Time (Utc)",
            $EndTime.ToString("o")
        )
    )
)
$M.Add(
    (
        [string]::Format(
            "{0,-21}: {1}",
            "Elapsed Time",
            $ElapsedTime
        )
    )
)
$M.Add(
    (
        [string]::Format(
            "{0,-21}: {1} ({2})",
            "Host Utc Offset",
            $OffsetText,
            [System.TimeZoneInfo]::Local.Id
        )
    )
)
$M.Add(
    (
        [string]::Format(
            "{0,-21}: {1}",
            "Host Name",
            $HostName
        )
    )
)
$M.Add(
    (
        [string]::Format(
            "{0,-21}: {1}\{2}",
            "Run As User",
            [System.Environment]::UserDomainName,
            [System.Environment]::UserName
        )
    )
)
$M.Add(
    (
        [string]::Format(
            "{0,-21}: {1}",
            "Elevated",
            $Elevated
        )
    )
)
$M.Add(
    (
        [string]::Format(
            "{0,-21}: {1} ({2})",
            "PowerShell",
            $PSVersionTable.PSVersion,
            $PSVersionTable.PSEdition
        )
    )
)
$M.Add(
    (
        [string]::Format(
            "{0,-21}: {1}",
            "OS",
            [System.Environment]::OSVersion.VersionString
        )
    )
)
$M.Add("")

$M.Add("===== PARAMETERS =====")

$M.Add(
    (
        [string]::Format(
            "{0,-21}: {1}",
            "Folder (given)",
            $Folder
        )
    )
)
$M.Add(
    (
        [string]::Format(
            "{0,-21}: {1}",
            "Folder (full)",
            $RootPath
        )
    )
)
$M.Add(
    (
        [string]::Format(
            "{0,-21}: {1}",
            "Include",
            $(if ($UseInclude) {
                $Include -join ", "
            }
            else {
                "(none)"
            })
        )
    )
)
$M.Add(
    (
        [string]::Format(
            "{0,-21}: {1}",
            "Exclude",
            $(if ($UseExclude) {
                $Exclude -join ", "
            }
            else {
                "(none)"
            })
        )
    )
)
$M.Add(
    (
        [string]::Format(
            "{0,-21}: {1}",
            "RunDate (Utc)",
            $RunDate
        )
    )
)

$M.Add("")

$M.Add("===== VOLUME (snapshot at start) =====")

foreach ($Key in $Volume.Keys) {
    $M.Add(
        (
            [string]::Format(
                "{0,-21}: {1}",
                $Key,
                $Volume[$Key]
            )
        )
    )
}

$M.Add("")

$M.Add("===== OUTPUT =====")

$M.Add(
    (
        [string]::Format(
            "{0,-21}: {1}",
            "Csv File",
            $OutFilePath
            )
        )
    )
$M.Add(
    (
        [string]::Format(
            "{0,-21}: {1}",
            "Csv size (bytes)",
            $CsvSize
        )
    )
)
$M.Add(
    (
        [string]::Format(
            "{0,-21}: {1}",
            "Csv SHA-256",
            $CsvHash
        )
    )
)
$M.Add(
    (
        [string]::Format(
            "{0,-21}: {1}",
            "File rows",
            $FileCount
        )
    )
)
$M.Add(
    (
        [string]::Format(
            "{0,-21}: {1}",
            "Directory rows",
            $DirRows
        )
    )
)
$M.Add(
    (
        [string]::Format(
            "{0,-21}: {1}",
            "Dirs walked",
            $DirCount
        )
    )
)

$M.Add("")

$M.Add("===== ERRORS ($( $Errors.Count )) =====")

if ($Errors.Count -gt 0) {
    $M.AddRange($Errors)
}
else {
    $M.Add("(none)")
}

$M.Add("")

$M.Add("===== SKIPPED DIRECTORIES ($( $Skipped.Count )) =====")

if ($Skipped.Count -gt 0) {
    $M.AddRange($Skipped)
}
else {
    $M.Add("(none)")
}

$ManifestPath = [System.IO.Path]::Combine(
    $CsvDir,
    "$( $RunDate )_$( $HostName )_$( $SerialPart )manifest.txt"
)

$MStream = $null
$MWriter = $null

try {
    $MStream = [System.IO.FileStream]::New(
        $ManifestPath,
        [System.IO.FileMode]::CreateNew,
        [System.IO.FileAccess]::Write,
        [System.IO.FileShare]::Read
    )
    $MWriter = [System.IO.StreamWriter]::New(
        $MStream,
        [System.Text.UTF8Encoding]::New($True)
    )
    foreach ($Line in $M) {
        $MWriter.WriteLine($Line)
    }
}
catch {
    Write-Warning "Could not write manifest => $( $_.Exception.Message )"
    $ManifestPath = $null
}
finally {
    if ($MWriter) {
        $MWriter.Dispose()
    }
    elseif ($MStream) {
        $MStream.Dispose()
    }
}

Write-Host ("`n" + "-" * 50 + "`n") `
    -ForegroundColor Green
Write-Host "[-] Script complete in " `
    -ForegroundColor Gray -NoNewline
Write-Host "$( $ElapsedTime )`n`n" `
    -ForegroundColor Green
Write-Host "[-] Enumerated`n" `
    -ForegroundColor Gray
Write-Host "    File$(if ($FileCount -eq 1) { " " } else { "s" })          :  " `
    -ForegroundColor Gray -NoNewline
Write-Host "$( [string]::Format($Inv, "{0:N0}", $FileCount) ) " `
    -ForegroundColor Green
Write-Host "    Director$(if ($DirRows -eq 1) { "y" } else { "ies" })    :  " `
    -ForegroundColor Gray -NoNewline
Write-Host "$( [string]::Format($Inv, "{0:N0}", $DirRows) ) " `
    -ForegroundColor Green
Write-Host "    Total Size     :  " `
    -ForegroundColor Gray -NoNewline
Write-Host "$(Get-FormattedFileSize -Bytes $TotalBytes)`n" `
    -ForegroundColor Green
Write-Host "[-] Output Path    :  " `
    -ForegroundColor Gray -NoNewline
Write-Host "$OutFilePath" `
    -ForegroundColor Green
Write-Host "[-] CSV SHA-256    :  " `
    -ForegroundColor Gray -NoNewline
Write-Host "$CsvHash" `
    -ForegroundColor Green

if ($ManifestPath) {
    Write-Host "`n[-] Manifest File  :  " `
        -ForegroundColor Gray -NoNewline
    Write-Host "$ManifestPath" -ForegroundColor Green
    Write-Host ("-" * 50) -ForegroundColor Green
}
else {
    Write-Host "-" * 50
}

if ($Errors.Count -gt 0) {
    Write-Host "`n [-] $( $Errors.Count ) error(s) => see the manifest." `
        -ForegroundColor Red
}

if ($Stopwatch.IsRunning) {
    $Stopwatch.Stop()
}
