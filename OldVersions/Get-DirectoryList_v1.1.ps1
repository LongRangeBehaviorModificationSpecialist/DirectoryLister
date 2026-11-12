<#
.SYNOPSIS
Recursively lists files and directories with their attributes to a CSV,
with a minimal footprint on the host, and documents the run in a manifest.

.DESCRIPTION
Walks a directory tree using read-only .NET directory enumeration and
streams one CSV row per file or directory straight to the output file.

The contents of the files being listed are never opened, no temporary
files are created, and nothing is written inside the folder being
examined.

Two lookups go beyond plain directory enumeration and leave more
host-side traces than the rest of the script: a CIM/WMI query for the
volume serial number (Win32_LogicalDisk), and the Storage cmdlets
(Get-Partition and Get-Disk) for the drive's physical serial number.

Files produced (both are written to '-CsvDir' and are never overwritten):

    CSV listing:
        "<RunDate>_<ComputerName>_<VolumeSerialNumber>_results.csv"

    Run manifest:
        "<RunDate>_<ComputerName>_<VolumeSerialNumber>_manifest.txt"

        Host, user, time and UTC offset, parameters, script SHA-256,
        volume details, row and directory counts, the SHA-256 of the
        CSV, plus any read errors and deliberately skipped directories.

    <RunDate> is the UTC start time formatted "yyyy.MM.dd_HH.mm.ss".

    If the volume serial number cannot be read, "NOSERIAL" is used in its
    place and a warning is displayed.

Include and Exclude

    -Include limits what is recorded.

        A directory that matches, and everything beneath it, is recorded.

        Directories that do not match are still walked so that matching
        directories nested deeper can be found, but their own files and
        directory rows are not recorded.  The files directly inside
        '-Folder' are recorded only if '-Folder' itself matches.

    -Exclude limits what is walked.

        A matching directory is pruned: it is never entered, nothing
        beneath it is read or recorded, and it is listed only in the
        SKIPPED DIRECTORIES section of the manifest.  Exclude is tested
        against subdirectories only, never against '-Folder' itself.

    Exclude wins over Include.

    Patterns are case-insensitive and may use wildcards (* ? [ab]).  A
    pattern must match the whole directory name or the whole full path,
    and is compared to both, so a plain name matches a directory with
    that name at any depth.

    A trailing "\" or "/" on a pattern is ignored.

        "Windows"              every directory named Windows
        "Program Files*"       Program Files and Program Files (x86)
        "C:\Users\Bob\Temp"    that one directory only
        "*\Cache"              any directory whose path ends in "\Cache"

Reparse points

    Reparse points (junctions and symbolic links) are never followed.
    A directory reparse point is not entered.  It is recorded as a Directory
    row with its target in the ReparseTarget column (when it is within the
    recorded scope) and is listed in the SKIPPED DIRECTORIES section of the
    manifest.  File symbolic links are recorded as File rows with their target.
    Reading a link's target opens the link itself (not what it points to);
    cloud placeholder files are not hydrated.

.PARAMETER Folder
Root folder to examine.  It must exist; otherwise the script exits with
code "1".

A relative path is resolved against the current PowerShell location.
Wildcard characters in the path are not expanded.  A trailing path separator
is ignored, except on a drive root.

.PARAMETER CsvDir
Directory in which the CSV and manifest files are saved.

The file names are generated automatically (see DESCRIPTION).  The directory
must already exist; the script does not create it.  A relative path is
resolved against the current PowerShell location.  It should be on separate
(external) media, NOT the machine being examined.  The script exits with
code "1" if the CSV would be created inside '-Folder', and warns if the
output is on the same volume as '-Folder'.

.PARAMETER Include
One or more directory names or paths to record.

When specified, only matching directories, and everything beneath them,
are recorded; other directories are walked but not recorded.  When omitted,
everything is recorded.  See DESCRIPTION for the matching rules.

.PARAMETER Exclude
One or more directory names or paths to prune.

A matching directory is never entered and nothing beneath it is recorded;
it is listed in the manifest as EXCLUDED.  Exclude wins over Include and is
never applied to '-Folder' itself.  See DESCRIPTION for the matching rules.

.INPUTS
None.  The script does not accept pipeline input.

.OUTPUTS
None.  The script writes no objects to the pipeline.

It creates the CSV and manifest files (see DESCRIPTION) and displays
progress, warnings and a summary in the console.

.EXAMPLE
.\Get-DirectoryList.ps1 -Folder C:\ -CsvDir E:\case1

Lists the whole C: drive and saves the CSV and manifest to E:\case1.

.EXAMPLE
.\Get-DirectoryList.ps1 `
-Folder C:\ `
-CsvDir E:\case1 `
-Exclude Windows,"Program Files*"

Lists C: but excludes every directory named Windows, and every directory
whose name starts with "Program Files", at any depth.

The excluded directories are recorded in the manifest.

.EXAMPLE
.\Get-DirectoryList.ps1 `
-Folder C:\Users `
-CsvDir E:\case1 `
-Include Documents,Desktop

Records only directories named Documents or Desktop (at any depth below
C:\Users) together with everything beneath them.  All other directories
are walked but not recorded.

.EXAMPLE
.\Get-DirectoryList.ps1 `
-Folder C:\Users\Bob `
-CsvDir E:\case1 `
-Include Documents `
-Exclude "*\node_modules"

Records Bob's Documents directories, but excludes any node_modules
directory beneath them (Exclude wins over Include).

.EXAMPLE
.\Get-DirectoryList.ps1 -Dir C:\ -Csv .

Uses the parameter aliases. "." is the current directory, so run this
from a location on external media.  If the current directory is inside
'-Dir', the script exits with an error instead of writing into the examined
tree.

.NOTES
Version: 1.1

Requirements

    Windows with PowerShell 5.0 or later.  The script uses CIM/WMI and the
    Storage module cmdlets, so it is not portable to other platforms.

Permissions

    Run elevated to read protected directories and, generally, to read
    the physical drive serial number.  The manifest records whether the
    run was elevated.  A directory that cannot be read is recorded in the
    ERRORS section of the manifest and the run continues.

Exit codes

    "1" if '-Folder' does not exist, '-CsvDir' does not exist, or the CSV
    would be created inside '-Folder'.  Otherwise the script ends normally,
    even if some items could not be read (see the ERRORS section of the
    manifest).

Files the script itself reads

    Besides the output files it writes, the script reads two files for
    hashing: the finished CSV (its SHA-256 is recorded in the manifest)
    and the script file (Script SHA-256).

CSV columns (UTF-16 LE with BOM; every field quoted; timestamps are UTC
in ISO 8601 round-trip format)

    DirectoryName       Containing directory of a file; parent directory
                        of a directory row (empty for a drive root)
    Name                File or directory name
    Type                "File" or "Directory"
    Extension           For files only
    SizeInBytes         Logical size in bytes (files only)
    SizeInKB            SizeInBytes / 1024, two decimals (files only)
    Mode                PowerShell mode string, for example "d-----"
    Attributes          .NET FileAttributes, for example "Hidden",
                        "Archive"
    CreationTimeUTC
    LastAccessTimeUTC
    LastWriteTimeUTC
    ReparseTarget       Target of a symbolic link or junction, otherwise
                        empty

    Rows are written depth-first: a directory's own row, then its files
    sorted by name, then its subdirectories in alphabetical order.

Manifest sections (UTF-8 with BOM)

    RUN, PARAMETERS, VOLUME (snapshot at start), OUTPUT, ERRORS and
    SKIPPED DIRECTORIES.  Entries in the last two are tab-separated.

Suggested launch (no profile, no prompts, execution policy bypass for
this process only).  Because -NonInteractive turns a missing mandatory
parameter into an error instead of a prompt, supply '-Folder' and '-CsvDir':

powershell.exe `
-NoProfile `
-NonInteractive `
-ExecutionPolicy Bypass `
-File .\Get-DirectoryList.ps1 `
-Folder C:\ `
-CsvDir E:\case1
#>
[CmdletBinding()]
[OutputType([void])]
param(
    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [Alias("Dir")]
    [string]$Folder,

    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [Alias("Csv")]
    [string]$CsvDir,

    [string[]]$Include,

    [string[]]$Exclude
)

$ScriptVersion         = "1.1"
[string]$RunDate       = [DateTime]::UtcNow.ToString("yyyy.MM.dd_HH.mm.ss")
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
    based units like Windows File Explorer.

    The rounded value is shown with two decimal places, followed by the
    exact byte count with thousands separators.

    Sizes under 1 KB are shown as plain bytes.

    The size can be supplied directly with -Bytes or looked up from a
    file with '-FilePath'.

    The invariant culture is always used so the text is the same on every
    computer.

    Example result: 1.50 MB (1,572,864 bytes)

    A value that would round up to 1,024.00 of a unit (for example
    1,048,575 bytes) is shown in the next larger unit instead.

    .PARAMETER FilePath
    Path to the file whose size should be formatted.

    The path is used literally (wildcard characters such as [ ] are not
    expanded) and must point to an existing file, otherwise a terminating
    error is thrown.

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
    Converts a TimeSpan into short, readable text for console and log
    messages.

    .DESCRIPTION
    Output depends on the length of the span (invariant culture is always
    used):

        under 1 minute   "12.34 seconds"
        under 1 hour     "5 min 07 sec"
        1 hour or more   "2 hr 05 min 09 sec"

    .PARAMETER Span
    The elapsed time to format, for example the Elapsed value of a
    Stopwatch.

    .EXAMPLE
    Format-Duration -Span (New-TimeSpan -Seconds 75)

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
        $Target.Replace('"', '""')
    ))
}


# ----------
# Validate input and output locations
# ----------

$RootPath = Resolve-FullPath $Folder

Write-Host "`nEnumerating directory: " -ForegroundColor Gray -NoNewline
Write-Host "$RootPath" -ForegroundColor Green -NoNewline
Write-Host ". Please wait, this may take a hot minute..." -ForegroundColor Gray

if (-not [System.IO.Directory]::Exists($RootPath)) {
    Write-Warning "Target directory was not found => $Folder"
    exit 1
}

# Strip a trailing separator unless the path is a drive/share root.
if ($RootPath.Length -gt [System.IO.Path]::GetPathRoot($RootPath).Length) {
    $RootPath = $RootPath.TrimEnd($Sep)
}

$CsvDir = Resolve-FullPath $CsvDir

if ([string]::IsNullOrEmpty($CsvDir)) {

    Write-Error "The '-Csv' parameter must include a directory path, e.g. E:\case1"
    exit 1
}
if (-not [System.IO.Directory]::Exists($CsvDir)) {
    # Deliberately not created
    # Creating directories is an extra change to the file system
    Write-Error "The output directory provided does not exist => $CsvDir"
    exit 1
}

# Volume serial number (opt-in: needs a CIM/WMI query).  It is looked up
# here, before the output path is built, because it becomes part of the
# output file name.

# Drive name (i.e., "F:\")
$VolRoot      = [System.IO.Path]::GetPathRoot($RootPath)
$Device       = $VolRoot.TrimEnd($Sep)  # (i.e., "F:")
$DriveLetter  = $VolRoot.TrimEnd(":\")  # (i.e., "F")
$VolumeSerial = $null
$SerialNote   = $null


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
    Write-Warning ("Volume Serial Number lookup failed ($SerialNote). " +
        "'NOSERIAL' will be used in the file name.")
    $VolumeSerial = "NOSERIAL"
}

$SerialPart  = if ($VolumeSerial) {
    "$( $VolumeSerial )_"
}
else {
    ""
}

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

$Drive = [System.IO.DriveInfo]::New($VolRoot)

$DriveType = $Drive.DriveType.ToString()
$DriveFormat = $Drive.DriveFormat

if ($DriveType -eq "Network") {
    $DriveSerialNumber = "N/A"
}
else {
    $DriveSerialNumber = (Get-Partition -DriveLetter $DriveLetter -ErrorAction Continue |
        Get-Disk).SerialNumber.Trim(" ")
}

$DriveVolLabel = $Drive.VolumeLabel
$DriveTotalSize = [string]::Format($Inv, "{0:N0}", $Drive.TotalSize)
$DriveFreeSpace = [string]::Format($Inv, "{0:N0}", $Drive.TotalFreeSpace)
$DriveIsReady = $Drive.IsReady


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

        # --- The directory's own row (before its contents) ---
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

        # --- Files in this directory (only if this branch is being captured) ---
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

        # --- Queue subdirectories (pushed in reverse so they pop in alphabetical order) ---
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

$Manifest = [System.Collections.Generic.List[string]]::New()

$Manifest.Add("===== RUN =====")

$Manifest.Add(
    (
        [string]::Format(
            "{0,-21}: {1}",
            "Script",
            $MyInvocation.MyCommand.Name
        )
    )
)
$Manifest.Add(
    (
        [string]::Format(
            "{0,-21}: {1}",
            "Version",
            $ScriptVersion
        )
    )
)
$Manifest.Add(
    (
        [string]::Format(
            "{0,-21}: {1}",
            "Script SHA-256",
            $ScriptHash
        )
    )
)
$Manifest.Add(
    (
        [string]::Format(
            "{0,-21}: {1}",
            "Start Time (Utc)",
            $StartTime.ToString("o")
        )
    )
)
$Manifest.Add(
    (
        [string]::Format(
            "{0,-21}: {1}",
            "End Time (Utc)",
            $EndTime.ToString("o")
        )
    )
)
$Manifest.Add(
    (
        [string]::Format(
            "{0,-21}: {1}",
            "Elapsed Time",
            $ElapsedTime
        )
    )
)
$Manifest.Add(
    (
        [string]::Format(
            "{0,-21}: {1} ({2})",
            "Host Utc Offset",
            $OffsetText,
            [System.TimeZoneInfo]::Local.Id
        )
    )
)
$Manifest.Add(
    (
        [string]::Format(
            "{0,-21}: {1}",
            "Host Name",
            $HostName
        )
    )
)
$Manifest.Add(
    (
        [string]::Format(
            "{0,-21}: {1}\{2}",
            "Run As User",
            [System.Environment]::UserDomainName,
            [System.Environment]::UserName
        )
    )
)
$Manifest.Add(
    (
        [string]::Format(
            "{0,-21}: {1}",
            "Elevated",
            $Elevated
        )
    )
)
$Manifest.Add(
    (
        [string]::Format(
            "{0,-21}: {1} ({2})",
            "PowerShell",
            $PSVersionTable.PSVersion,
            $PSVersionTable.PSEdition
        )
    )
)
$Manifest.Add(
    (
        [string]::Format(
            "{0,-21}: {1}",
            "OS",
            [System.Environment]::OSVersion.VersionString
        )
    )
)
$Manifest.Add("")
$Manifest.Add("")

$Manifest.Add("===== PARAMETERS =====")

$Manifest.Add(
    (
        [string]::Format(
            "{0,-21}: {1}",
            "Folder (given)",
            $Folder
        )
    )
)
$Manifest.Add(
    (
        [string]::Format(
            "{0,-21}: {1}",
            "Folder (full)",
            $RootPath
        )
    )
)
$Manifest.Add(
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
$Manifest.Add(
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
$Manifest.Add(
    (
        [string]::Format(
            "{0,-21}: {1}",
            "RunDate (Utc)",
            $RunDate
        )
    )
)

$Manifest.Add("")
$Manifest.Add("")

$Manifest.Add("===== VOLUME (snapshot at start) =====")

$Manifest.Add(
    (
        [string]::Format(
            "{0,-21}: {1}",
            "VolumeRoot",
            $VolRoot
        )
    )
)
$Manifest.Add(
    (
        [string]::Format(
            "{0,-21}: {1}",
            "DriveType",
            $DriveType
        )
    )
)
$Manifest.Add(
    (
        [string]::Format(
            "{0,-21}: {1}",
            "DriveFormat",
            $DriveFormat
        )
    )
)
$Manifest.Add(
    (
        [string]::Format(
            "{0,-21}: {1}",
            "Drive Physical SN",
            $DriveSerialNumber
        )
    )
)
$Manifest.Add(
    (
        [string]::Format(
            "{0,-21}: {1}",
            "VolumeLabel",
            $DriveVolLabel
        )
    )
)
$Manifest.Add(
    (
        [string]::Format(
            "{0,-21}: {1}",
            "Volume SN",
            $VolumeSerial
        )
    )
)
$Manifest.Add(
    (
        [string]::Format(
            "{0,-21}: {1} bytes",
            "DriveTotalSize",
            $DriveTotalSize
        )
    )
)
$Manifest.Add(
    (
        [string]::Format(
            "{0,-21}: {1} bytes",
            "DriveFreeSpace",
            $DriveFreeSpace
        )
    )
)
$Manifest.Add(
    (
        [string]::Format(
            "{0,-21}: {1}",
            "DriveIsReady",
            $DriveIsReady
        )
    )
)

$Manifest.Add("")
$Manifest.Add("")

$Manifest.Add("===== OUTPUT =====")

$Manifest.Add(
    (
        [string]::Format(
            "{0,-21}: {1}",
            "Csv File Name",
            $OutFilePath
            )
        )
)
$Manifest.Add(
    (
        [string]::Format(
            "{0,-21}: {1} {2}",
            "Csv File Size",
            [string]::Format(
                $Inv,
                "{0:N0}",
                $CsvSize
            ),
            "(bytes)"
        )
    )
)
$Manifest.Add(
    (
        [string]::Format(
            "{0,-21}: {1}",
            "Csv SHA-256",
            $CsvHash
        )
    )
)
$Manifest.Add(
    (
        [string]::Format(
            "{0,-21}: {1}",
            "File rows",
            [string]::Format(
                $Inv,
                "{0:N0}",
                $FileCount
            )
        )
    )
)
$Manifest.Add(
    (
        [string]::Format(
            "{0,-21}: {1}",
            "Directory rows",
            [string]::Format(
                $Inv,
                "{0:N0}",
                $DirRows
            )
        )
    )
)
$Manifest.Add(
    (
        [string]::Format(
            "{0,-21}: {1}",
            "Dirs walked",
            $DirCount
        )
    )
)

$Manifest.Add("")
$Manifest.Add("")

$Manifest.Add("===== ERRORS ($( $Errors.Count )) =====")

if ($Errors.Count -gt 0) {
    $Manifest.AddRange($Errors)
}
else {
    $Manifest.Add("(none)")
}

$Manifest.Add("")
$Manifest.Add("")

$Manifest.Add("===== SKIPPED DIRECTORIES ($( $Skipped.Count )) =====")

if ($Skipped.Count -gt 0) {
    $Manifest.AddRange($Skipped)
}
else {
    $Manifest.Add("(none)")
}

$ManifestPath = [System.IO.Path]::Combine(
    $CsvDir,
    "$( $RunDate )_$( $HostName )_$( $SerialPart )manifest.txt"
)

$MStream = $null
$MWriter = $null

# Create the manifest file and stream the data to the file.
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
    foreach ($Line in $Manifest) {
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


Write-Host ("`n" + "-" * 50 + "`n") -ForegroundColor Green
Write-Host "Script complete in " -ForegroundColor Gray -NoNewline
Write-Host "$( $ElapsedTime )`n" -ForegroundColor Green
Write-Host "Enumerated`n" -ForegroundColor Gray
Write-Host "  File$(if ($FileCount -eq 1) { " " } else { "s" })       :  " `
    -ForegroundColor Gray -NoNewline
Write-Host "$( [string]::Format($Inv, "{0:N0}", $FileCount) ) " `
    -ForegroundColor Green
Write-Host "  Director$(if ($DirRows -eq 1) { "y" } else { "ies" }) :  " `
    -ForegroundColor Gray -NoNewline
Write-Host "$( [string]::Format($Inv, "{0:N0}", $DirRows) ) " `
    -ForegroundColor Green
Write-Host "  Total Size  :  " -ForegroundColor Gray -NoNewline
Write-Host "$(Get-FormattedFileSize -Bytes $TotalBytes)`n" `
    -ForegroundColor Green
Write-Host "Output Path   :  " -ForegroundColor Gray -NoNewline
Write-Host "$OutFilePath" -ForegroundColor Green
Write-Host "CSV SHA-256   :  " -ForegroundColor Gray -NoNewline
Write-Host "$CsvHash" -ForegroundColor Green


if ($ManifestPath) {
    Write-Host "`nManifest File :  " `
        -ForegroundColor Gray -NoNewline
    Write-Host "$ManifestPath" -ForegroundColor Green
    Write-Host ("`n" + "-" * 50) -ForegroundColor Green
}
else {
    Write-Host "`n" + "-" * 50
}

if ($Errors.Count -gt 0) {
    Write-Host "`n$( $Errors.Count ) error(s) => see the manifest for details.`n" `
        -ForegroundColor Red
}

if ($Stopwatch.IsRunning) {
    $Stopwatch.Stop()
}
