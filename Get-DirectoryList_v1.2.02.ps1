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
    ('Get-Partition' and 'Get-Disk') for the drive's physical serial number.

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

    Examiner notice

        Before anything is read or written, the script shows a notice about
        cloud-synchronized folders (OneDrive, Dropbox, Google Drive, Proton
        Drive and similar) and about preserving the integrity of the machine
        being examined, then waits for the examiner to type Y.  Any other
        answer cancels the run and nothing is written.  In a non-interactive
        session, or with '-AcceptNotice', the notice is displayed and the script
        carries on.  The start time and the elapsed time are counted from the
        moment the notice is confirmed.  The outcome is recorded in the
        manifest (ExaminerNotice).

    Progress

        While the tree is walked, one status line is redrawn in place on the
        console about four times a second: elapsed time, directories walked,
        files recorded and the directory currently being read.  The total is
        not known in advance, so no percentage is shown.  When the console
        output is redirected, a plain line is written every 10 seconds instead.
        The progress display only writes to the console; it adds nothing to the
        CSV, the manifest or the examined volume.

    Include and Exclude

        '-Include' limits what is recorded.

            A directory that matches, and everything beneath it, is recorded.

            Directories that do not match are still walked so that matching
            directories nested deeper can be found, but their own files and
            directory rows are not recorded.  The files directly inside
            '-Folder' are recorded only if '-Folder' itself matches.

        '-Exclude' limits what is walked.

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

            "Windows"              Every directory named Windows
            "Program Files*"       Program Files and Program Files (x86)
            "C:\Users\Bob\Temp"    That one directory only
            "*\Cache"              Any directory whose path ends in "\Cache"

    Reparse points

        By default, reparse points (junctions and symbolic links) are never
        followed.  A directory reparse point is not entered.  It is recorded as a
        Directory row with its target in the ReparseTarget column (when it is
        within the recorded scope) and is listed in the SKIPPED DIRECTORIES
        section of the manifest.  File symbolic links are recorded as File rows
        with their target.  Reading a link's target opens the link itself (not
        what it points to); cloud placeholder files are not hydrated.

        Only junctions and symbolic links count as links.  Other kinds of
        reparse point, such as the folders of a cloud-storage client (Proton
        Drive, OneDrive and similar), have no target and are not links: they are
        walked like ordinary directories, with or without '-FollowReparsePoints'.
        Their Attributes column shows "ReparsePoint".  Listing names does not
        download file content, but browsing a partly synced folder can make the
        sync client fetch that folder's listing.

        With '-FollowReparsePoints', directory reparse points are entered and the
        files and directories inside them are recorded.  They appear under the
        path of the link (for example "C:\Users\All Users\..."), not under the
        path of the real target, so the same file can be listed twice: once under
        the real path and once under the link.  Totals count both.

        Following is made safe against endless loops:

            - A link whose target contains the link itself is not followed.
            - Each target is followed through only one link; any further link to
              the same target is recorded as a row but not entered.
            - A link whose target cannot be resolved is not followed.

        Followed links are listed in the FOLLOWED REPARSE POINTS section of the
        manifest.  Links that were not followed, with the reason, stay in the
        SKIPPED DIRECTORIES section.  Some legacy junctions (for example
        "Application Data" inside a user profile) deny access by design; those
        appear in the ERRORS section and the run continues.

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

.PARAMETER FollowReparsePoints
    Enter directory reparse points (junctions and directory symbolic links) and
    record what is inside them.  Off by default.  Loop protection applies; see
    DESCRIPTION.  The setting is recorded in the manifest.

.PARAMETER AcceptNotice
    Shows the examiner notice (see DESCRIPTION) but does not stop to ask the
    examiner to confirm it.  Intended for unattended runs.  Without this switch
    the script waits until the examiner types "Y".  The manifest records which of
    the two happened.

.INPUTS
    None.  The script does not accept pipeline input.

.OUTPUTS
    None.  The script writes no objects to the pipeline.

    It creates the CSV and manifest files (see DESCRIPTION) and displays
    progress, warnings and a summary in the console.

.EXAMPLE
    .\Get-DirectoryList.ps1 -Folder C:\ -CsvDir E:\cases

    Lists the whole C: drive and saves the CSV and manifest to E:\cases.

.EXAMPLE
    .\Get-DirectoryList.ps1 -Folder C:\ -CsvDir E:\cases -Exclude Windows,"Program Files*"

    Lists "C:" but excludes every directory named Windows, and every directory
    whose name starts with "Program Files", at any depth.

    The excluded directories are recorded in the manifest.

.EXAMPLE
    .\Get-DirectoryList.ps1 -Folder C:\Users -CsvDir E:\cases -Include Documents,Desktop

    Records only directories named Documents or Desktop (at any depth below
    C:\Users) together with everything beneath them.  All other directories
    are walked but not recorded.

.EXAMPLE
    .\Get-DirectoryList.ps1 -Folder C:\Users\Bob -CsvDir E:\cases -Include Documents -Exclude "*\node_modules"

    Records Bob's Documents directories, but excludes any node_modules
    directory beneath them (Exclude wins over Include).

.EXAMPLE
    .\Get-DirectoryList.ps1 -Folder C:\Users -CsvDir E:\cases -FollowReparsePoints

    Lists C:\Users and also enters junctions and directory symbolic links,
    recording the files inside them under the link's path.

.EXAMPLE
    .\Get-DirectoryList.ps1 -Dir C:\ -Csv .

    Uses the parameter aliases.  "." is the current directory, so run this
    from a location on external media.  If the current directory is inside
    '-Dir', the script exits with an error instead of writing into the examined
    tree.

.NOTES
    Version: 1.2.02

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

        RUN, PARAMETERS, VOLUME (snapshot at start), OUTPUT, ERRORS,
        FOLLOWED REPARSE POINTS and SKIPPED DIRECTORIES.  Entries in the last
        three are tab-separated.

    Suggested launch (no profile, no prompts, execution policy bypass for
    this process only).  Because '-NonInteractive' turns a missing mandatory
    parameter into an error instead of a prompt, supply '-Folder' and '-CsvDir':

    powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File .\Get-DirectoryList.ps1 -Folder C:\ -CsvDir E:\cases

.LINK
    https://github.com/LongRangeBehaviorModificationSpecialist/DirectoryLister
#>
[CmdletBinding()]
[OutputType([void])]
param
(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [Alias("Dir")]
    [string]$Folder,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [Alias("Csv")]
    [string]$CsvDir,

    [string[]]$Include,

    [string[]]$Exclude,

    [Alias("Follow")]
    [switch]$FollowReparsePoints,

    [switch]$AcceptNotice
)


function Show-Message
{
    param
    (
        [Parameter(Position = 0)]
        [string[]]$Text,

        [switch]$NoTime,
        [switch]$NoNewline,
        [System.ConsoleColor[]]$Color = @(),
        [Nullable[System.ConsoleColor]]$BackgroundColor,
        [int]$LinesAbove = 0,
        [int]$Indent
    )

    if ($LinesAbove -gt 0)
    {
        Write-Host ("`n" * $LinesAbove) -NoNewline
    }

    if (-not $NoTime)
    {
        $Timestamp = (Get-Date).ToUniversalTime().ToString("yyyy.MM.ddTHH.mm.ss.fffZ")
        Write-Host "[$Timestamp] " -ForegroundColor DarkCyan -NoNewline
    }

    # Apply dynamic indentation if the parameter was provided
    if ($PSBoundParameters.ContainsKey("Indent") -and $Indent -gt 0)
    {
        Write-Host (" " * $Indent) -NoNewline
    }

    for ($i = 0; $i -lt $Text.Count; $i++)
    {
        $hostArgs = @{ Object = $Text[$i]; NoNewline = $true }

        if ($i -lt $Color.Count)
        {
            $hostArgs.ForegroundColor = $Color[$i]
        }

        if ($null -ne $BackgroundColor)
        {
            $hostArgs.BackgroundColor = $BackgroundColor
        }
        Write-Host @hostArgs
    }
    Write-Host
}


# Temporarily disable PowerShell history logging for this process
# Minimize local artifact footprint by disabling PSReadLine history logging
# for this session
if (Get-Module PSReadLine)
{
    Set-PSReadLineOption -HistorySaveStyle SaveNothing -ErrorAction SilentlyContinue
    Show-Message -Text "'PSReadLineOption' has been disabled during execution this script." `
        -Color Green -LinesAbove 1
}

$ScriptVersion         = "1.2.02"
$HostName              = [System.Environment]::MachineName
$Sep                   = [System.IO.Path]::DirectorySeparatorChar
$Inv                   = [System.Globalization.CultureInfo]::InvariantCulture
$ErrorActionPreference = [System.Management.Automation.ActionPreference]::Continue
[int]$LargeIndent = 27
[int]$ShortIndent = 4


function Resolve-FullPath
{
    # Resolves a path against the PowerShell location (not the .NET
    # process directory) without touching the disk and without
    # expanding wildcards.
    param
    (
        [string]$Path
    )
    [System.IO.Path]::GetFullPath(
        $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
    )
}


function New-DirPatterns
{
    param
    (
        [string[]]$Names
    )
    $Opts = [System.Management.Automation.WildcardOptions]::IgnoreCase -bor
            [System.Management.Automation.WildcardOptions]::CultureInvariant

    $List = [System.Collections.Generic.List[System.Management.Automation.WildcardPattern]]::New()

    foreach ($Name in $Names)
    {
        if ([string]::IsNullOrWhiteSpace($Name))
        {
            continue
        }

        $List.Add([System.Management.Automation.WildcardPattern]::New(
            $Name.Trim().TrimEnd("\", "/"),
            $Opts)
        )
    }
    # The leading comma stops PowerShell from unrolling the list
    , $List
}


function Test-DirMatch
{
    param
    (
        [System.IO.DirectoryInfo]$Dir,
        $Patterns
    )
    foreach ($Pattern in $Patterns)
    {
        if ($Pattern.IsMatch($Dir.Name) -or $Pattern.IsMatch($Dir.FullName))
        {
            return $true
        }
    }
    return $false
}


function Get-FormattedFileSize
{
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
    param
    (
        [Parameter(Mandatory, Position = 0, ParameterSetName = "ByPath")]
        [ValidateNotNullOrEmpty()]
        [string]$FilePath,

        [Parameter(Mandatory, ParameterSetName = "ByBytes")]
        [ValidateRange(0, [long]::MaxValue)]
        [long]$Bytes
    )

    if ($PSCmdlet.ParameterSetName -eq "ByPath")
    {
        if (-not (Test-Path -LiteralPath $FilePath -PathType Leaf))
        {
            throw "File not found => $FilePath."
        }
        $Bytes = (Get-Item -LiteralPath $FilePath).Length
    }

    $Inv = [System.Globalization.CultureInfo]::InvariantCulture

    if ($Bytes -lt 1KB)
    {
        return [string]::Format($Inv, "{0:N0} bytes", $Bytes)
    }

    # Divide by 1024 until the rounded value is below 1,024
    # (or the largest unit is reached)
    $Units = "KB", "MB", "GB", "TB"
    $Size  = [double]$Bytes
    $Index = -1
    do
    {
        $Size /= 1024
        $Index++
    } while ([math]::Round($Size, 2) -ge 1024 -and $Index -lt ($Units.Count - 1))

    return [string]::Format(
        $Inv,
        "{0:N2} {1} ({2:N0} bytes)",
        $Size,
        $Units[$Index],
        $Bytes
    )
}


function Get-ReparseTarget
{
    # Returns the link target for symlinks/junctions, or "" for anything else.
    param
    (
        [System.IO.FileSystemInfo]$Item
    )

    if (-not $Item.Attributes.HasFlag([System.IO.FileAttributes]::ReparsePoint))
    {
        return ""
    }

    try
    {
        return (@($Item.Target) -join "; ")
    }
    catch
    {
        return ""
    }
}

function Test-DirIsLink
{
    # True for junctions and symbolic links.  Other reparse points (cloud
    # placeholders, deduplication and so on) have no link target and are
    # ordinary directories as far as listing goes.
    param
    (
        [System.IO.FileSystemInfo]$Item,
        [string]$Target
    )

    if ($Target)
    {
        return $true
    }

    try
    {
        $LinkType = [string]$Item.LinkType
    }
    catch
    {
        $LinkType = ""
    }

    return ($LinkType -eq "Junction" -or $LinkType -eq "SymbolicLink")
}


function Resolve-LinkTargetPath
{
    # Returns the full path of an existing directory that a reparse point
    # points to, without trailing separator, or "$null" if it cannot be
    # resolved.  Relative targets are resolved against the link's parent.
    param
    (
        [string]$LinkPath,
        [string]$Target
    )

    if ([string]::IsNullOrWhiteSpace($Target))
    {
        return $null
    }

    try
    {
        $First = ($Target -split "; ")[0]

        if (-not [System.IO.Path]::IsPathRooted($First))
        {
            $First = [System.IO.Path]::Combine(
                [System.IO.Path]::GetDirectoryName($LinkPath),
                $First
            )
        }
        $Full = [System.IO.Path]::GetFullPath($First)

        if ([System.IO.Directory]::Exists($Full))
        {
            return $Full.TrimEnd($Sep)
        }
    }
    catch
    {
        # Fall through: treated as unresolved
    }

    return $null
}


function Write-DirNotFoundMessage
{
    param
    (
        [ValidateSet("TARGET","OUTPUT")]
        [string]$DirType,
        [string]$Folder
    )

    Show-Message -Text "The $DirType DIRECTORY provided was not found => `"$Folder`"" `
        -Color White -BackgroundColor Red -LinesAbove 1
    Show-Message -Text "Check the directory path and try again." -Color Red -NoTime -Indent $LargeIndent
    Show-Message -Text "EXITING..." -Color Red -NoTime -Indent $LargeIndent

}


function Format-Duration
{
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
    param
    (
        [Parameter(Mandatory)]
        [TimeSpan]$Span
    )

    if ($Span.TotalSeconds -lt 60)
    {
        return [string]::Format(
            $Inv,
            "{0:N6} seconds",
            $Span.TotalSeconds
        )
    }

    if ($Span.TotalHours -lt 1)
    {
        return [string]::Format(
            $Inv,
            "{0} min {1:00} sec",
            $Span.Minutes,
            $Span.Seconds
        )
    }

    [string]::Format(
        $Inv,
        "{0} hr {1:00} min {2:00} sec",
        [int][math]::Floor($Span.TotalHours),
        $Span.Minutes,
        $Span.Seconds
    )
}


function Get-FileSha256
{
    param
    (
        [string]$Path
    )

    $Fs = [System.IO.FileStream]::New(
        $Path,
        [System.IO.FileMode]::Open,
        [System.IO.FileAccess]::Read,
        [System.IO.FileShare]::ReadWrite
    )
    try
    {
        $Sha = [System.Security.Cryptography.SHA256]::Create()
        try
        {
            [System.BitConverter]::ToString($Sha.ComputeHash($Fs)).Replace("-", "")
        }
        finally
        {
            $Sha.Dispose()
        }
    }
    finally
    {
        $Fs.Dispose()
    }
}


function Write-Row
{
    # Positional parameters (faster than named ones on very large trees):
    # Type ("File"|"Directory"), containing-directory path, the item,
    # reparse target ("" if none)
    param
    (
        [string]$Type,
        [string]$DirName,
        [System.IO.FileSystemInfo]$Item,
        [string]$Target
    )

    if ($Type -eq "File")
    {
        $Len   = $Item.Length
        $Ext   = $Item.Extension
        $Bytes = $Len.ToString($Inv)
        $Kb    = ([math]::Round($Len / 1KB, 2)).ToString($Inv)
    }
    else
    {
        $Ext = ""
        $Bytes = ""
        $Kb = ""
    }

    $Writer.WriteLine(
        (
            '"{0}","{1}","{2}","{3}","{4}","{5}","{6}","{7}","{8}","{9}","{10}","{11}"' -f
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
        )
    )
}


# ------------------------------
#
# Progress display
#
# ------------------------------

# One status line is redrawn in place, at most once every $ProgressEveryMs
# milliseconds, so reporting costs next to nothing on very large trees.  The
# walk streams to the CSV and never builds a full list first, so the total is
# unknown: counters are shown instead of a percentage.  If console output is
# redirected, a plain line is written every 2.5 seconds instead.

$ProgressRedirected = [Console]::IsOutputRedirected
$ProgressEveryMs    = if ($ProgressRedirected) { 2500 }  else { 250 }
$ProgressTimer      = [System.Diagnostics.Stopwatch]::StartNew()
$ProgressActive     = $false


function Get-ProgressWidth
{
    try
    {
        [math]::Max(20, [Console]::WindowWidth - 1)
    }
    catch
    {
        119
    }
}

function Write-ScanProgress
{
    param
    (
        [string]$CurrentDir
    )

    $ProgressTimer.Restart()

    $Span   = $Stopwatch.Elapsed

    # Use escaped comma '000\,000' to display a literal comma with leading zero padding
    $FormattedDirs  = $DirCount.ToString("000\,000", $Inv)
    $FormattedFiles = $FileCount.ToString("000\,000", $Inv)

    $Prefix = [string]::Format(
        $Inv,
        "[{0:00}:{1:00}:{2:00}] Dirs: {3} | Files: {4} | ",
        [int][math]::Floor($Span.TotalHours),
        $Span.Minutes,
        $Span.Seconds,
        $FormattedDirs,
        $FormattedFiles
    )

    $MaxWidth      = 120
    $WinWidth      = Get-ProgressWidth
    $MaxTotalWidth = [math]::Min($MaxWidth, $WinWidth)
    # Subtract 4 for the requested screen indent
    $MaxLineLength = [math]::Max(20, $MaxTotalWidth - 4)

    if ($ProgressRedirected)
    {
        Show-Message -Text "$Prefix$CurrentDir" -Color Gray -NoTime -Indent $ShortIndent
        return
    }

    $RemainingForDir = $MaxLineLength - $Prefix.Length

    # If space is too tight, guarantee a minimum window for the path instead of hiding it
    if ($RemainingForDir -lt 15)
    {
        $RemainingForDir = 15
    }

    # Truncate $CurrentDir in the middle if it exceeds the remaining budget
    if ($CurrentDir.Length -gt $RemainingForDir)
    {
        $HeadLength = 3
        $TailLength = [math]::Max(10, $RemainingForDir - 6)  # 3 chars for '...' + head + tail budget

        if ($CurrentDir.Length -gt ($HeadLength + $TailLength))
        {
            $Head       = $CurrentDir.Substring(0, $HeadLength)
            $Tail       = $CurrentDir.Substring($CurrentDir.Length - $TailLength)
            $CurrentDir = "$Head...$Tail"
        }
    }

    $Line = $Prefix + $CurrentDir

    # Ensures it fits within the max boundary and pads out old artifacts
    if ($Line.Length -gt $MaxLineLength)
    {
        $Line = $Line.Substring(0, $MaxLineLength)
    }

    Show-Message -Text ($Line.PadRight($MaxLineLength)) -Color Gray -NoTime -Indent 4
    $script:ProgressActive = $true
}


function Clear-ScanProgress
{
    # Blanks the status line and leaves the cursor at the start of it.
    if ($ProgressActive)
    {
        Show-Message -Text ("`r" + (" " * (Get-ProgressWidth)) + "`r") -Color Gray -NoNewline
        $script:ProgressActive = $false
    }
}


function Write-ScanWarning
{
    # Clears the status line first so the warning is not glued to it.
    param
    (
        [string]$Message
    )

    Clear-ScanProgress
    Show-Message -Text $Message -Color Yellow
}


# ------------------------------
#
# Examiner notice
#
# ------------------------------

# Shown before anything is read or written.

function Show-ExaminerNotice
{
    $Notice = @'
+------------------------------------------------+
|                                                |
|                  !! NOTICE !!                  |
|                                                |
+------------------------------------------------+

CLOUD-SYNCED FOLDERS

Folders managed by a cloud-sync client can hold "online-only"
placeholder files: the name, size and dates are stored on this machine,
but the content is not.

A running sync client can also upload, rename, delete or otherwise change
files and folders, including changes made from other devices.  Any of
this alters the machine being examined.

This script reads names, sizes, dates and attributes only.  It does
not open file contents.

RECOMMENDED BEFORE RUNNING

1.  Exit every cloud-sync client, including its system-tray icon
    and background processes.  Do not sign out of or unlink any account.
2.  If possible, disconnect the machine from the network so placeholders
    cannot be fetched.  Folders that are not already populated on this
    machine may then list incompletely.
3.  Do not open, preview, copy, hash or scan cloud-only files.
4.  Record in your notes which sync clients were running, what you
    stopped, and when.

PRESERVING THE MACHINE BEING EXAMINED

1.  Document that this script was run, and when.
2.  Save the output to separate (external) media, NOT to the machine being
    examined.
3.  In the CSV, the Attributes column shows which items are online-only
    The size listed for those items is the size in the cloud, not space used
    on this machine.
4.  The listing shows this machine's state at the time of the run.  It may
    not match the full contents of the cloud account.

+------------------------------------------------+
'@

    Show-Message -Text "" -NoTime
    foreach ($Line in ($Notice -split "\r?\n"))
    {
        if ($Line -match '^\+-|\|' -or $Line -match '^[A-Z][A-Z \-]+$')
        {
            Show-Message -Text $Line -Color Yellow -NoTime
        }
        else
        {
            Show-Message -Text $Line -Color Gray -NoTime
        }
    }
}

Show-ExaminerNotice

$NoticeCancelled = $false

if ($AcceptNotice)
{
    $NoticeStatus = "Displayed; confirmation skipped (-AcceptNotice)"
}
else
{
    try
    {
        $Answer = Read-Host "[-] Type Y to confirm you have read this notice and wish to continue" `
            -ErrorAction Stop
        if ($Answer -match '^\s*(y|yes)\s*$')
        {
            $NoticeStatus = "Acknowledged by examiner"
        }
        else
        {
            $NoticeCancelled = $true
        }
    }
    catch
    {
        # -NonInteractive (or no console): nobody can be asked
        $NoticeStatus = "Displayed; confirmation not possible (non-interactive session)"
    }
}

if ($NoticeCancelled)
{
    Show-Message -Text "Cancelled by the examiner. Nothing was read or written." `
        -Color Yellow -LinesAbove 1
    exit 1
}

$NoticeStatus += " at $( [DateTime]::UtcNow.ToString("yyyy-MM-dd HH:mm:ss", $Inv) ) UTC"

# Start the clock when the examination actually begins, not while the
# notice was being read.
[string]$RunDate = [DateTime]::UtcNow.ToString("yyyy.MM.dd_HH.mm.ss")
$StartTime       = [DateTime]::UtcNow
$Stopwatch       = [System.Diagnostics.Stopwatch]::StartNew()


# ------------------------------
#
# Validate input &  output locations
#
# ------------------------------

$RootPath = Resolve-FullPath -Path $Folder

Show-Message -Text ("Enumerating target directory: ",
    "$RootPath ",
    "and all sub-directories...") -Color Gray, Green, Gray -LinesAbove 1

Show-Message -Text "Please wait, this may take a hot minute..." `
    -Color Green -NoTime -LinesAbove 1 -Indent $ShortIndent
Show-Message -Text "" -NoTime

if (-not [System.IO.Directory]::Exists($RootPath))
{
    Write-DirNotFoundMessage -DirType TARGET -Folder $Folder
    exit 1
}

# Strip a trailing separator unless the path is a drive/share root.
if ($RootPath.Length -gt [System.IO.Path]::GetPathRoot($RootPath).Length)
{
    $RootPath = $RootPath.TrimEnd($Sep)
}

$CsvDir = Resolve-FullPath -Path $CsvDir

if ([string]::IsNullOrEmpty($CsvDir))
{

    Show-Message -Text ("The '-Csv' parameter must include a directory " +
        "path, e.g. E:\case1") -Color Red
    exit 1
}

if (-not [System.IO.Directory]::Exists($CsvDir))
{
    # Deliberately not created
    # Creating directories is an extra change to the file system
    Write-DirNotFoundMessage -DirType OUTPUT -Folder $CsvDir
    exit 1
}

# Volume serial number (opt-in: needs a CIM/WMI query).  It is looked up
# here, before the output path is built, because it becomes part of the
# output file name.


$VolRoot      = [System.IO.Path]::GetPathRoot($RootPath)  # returns "F:\"
$Device       = $VolRoot.TrimEnd($Sep)  # returns "F:"
$DriveLetter  = $VolRoot.TrimEnd(":\")  # returns "F"
$VolumeSerial = $null
$SerialNote   = $null


if ($Device -match "^[A-Za-z]:$")
{
    try
    {
        $Ld = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='$Device'" `
            -Property VolumeSerialNumber -ErrorAction Stop
        # Keep only letters/digits so the value is always safe in a file name.
        if ($Ld)
        {
            $VolumeSerial = ([string]$Ld.VolumeSerialNumber) -replace "[^0-9A-Za-z]", ""
        }
    }
    catch
    {
        $SerialNote = "Unavailable => $( $_.Exception.Message )"
    }
}
else
{
    $SerialNote = "Unavailable (not a drive-letter volume)"
}

if (-not $VolumeSerial)
{
    if (-not $SerialNote)
    {
        $SerialNote = "Unavailable"
    }
    Show-Message -Text ("Volume Serial Number lookup failed ($SerialNote). " +
        "'NOSERIAL' will be used in the file name.") -Color Yellow
    $VolumeSerial = "NOSERIAL"
}

$SerialPart  = if ($VolumeSerial)
{
    "$( $VolumeSerial )_"
}
else
{
    ""
}

$OutFilePath = [System.IO.Path]::Combine(
    $CsvDir,
    "$( $RunDate )_$( $HostName )_$( $SerialPart )results.csv"
)


# Never write into the tree being examined
# Would alter the target and pollute the listing
$RootPrefix = $RootPath.TrimEnd($Sep) + $Sep
if ($OutFilePath.StartsWith(
    $RootPrefix,
    [System.StringComparison]::OrdinalIgnoreCase))
{
    Show-Message -Text ("Output file would be written inside the folder " +
        "being examined. Choose an output location outside $RootPath " +
        "(ideally external media).") -Color Red
    exit 1
}

if ([string]::Equals(
        [System.IO.Path]::GetPathRoot($RootPath),
        [System.IO.Path]::GetPathRoot($OutFilePath),
        [System.StringComparison]::OrdinalIgnoreCase
        )
) {
    Show-Message -Text ("Output is on the same volume as the folder being examined. " +
        "For forensic soundness, write to separate media.") -Color Yellow
}

$IncludePatterns = New-DirPatterns -Names $Include
$ExcludePatterns = New-DirPatterns -Names $Exclude
$UseInclude      = $IncludePatterns.Count -gt 0
$UseExclude      = $ExcludePatterns.Count -gt 0


# ------------------------------
#
# Get Volume details
#
# ------------------------------

# Metadata only -- snapshot taken before the walk starts

$Drive       = [System.IO.DriveInfo]::New($VolRoot)
$DriveType   = $Drive.DriveType.ToString()
$DriveFormat = $Drive.DriveFormat

if ($DriveType -eq "Network")
{
    $DriveSerialNumber = "N/A"
}
else
{
    $DriveSerialNumber = (Get-Partition -DriveLetter $DriveLetter `
        -ErrorAction Continue |
        Get-Disk).SerialNumber.Trim(" ")
}

$DriveVolLabel  = $Drive.VolumeLabel
$DriveTotalSize = [string]::Format($Inv, "{0:N0}", $Drive.TotalSize)
$DriveFreeSpace = [string]::Format($Inv, "{0:N0}", $Drive.TotalFreeSpace)
$DriveIsReady   = $Drive.IsReady


# ------------------------------
#
# Walk the tree & stream rows to the CSV
#
# ------------------------------

$Errors    = [System.Collections.Generic.List[string]]::New()
$Skipped   = [System.Collections.Generic.List[string]]::New()
$Followed  = [System.Collections.Generic.List[string]]::New()

# Real paths of reparse-point targets already followed (loop protection)
$FollowedTargets = [System.Collections.Generic.HashSet[string]]::New(
    [System.StringComparer]::OrdinalIgnoreCase
)
$FileCount = [int]0
$DirRows   = [int]0
$DirCount  = [int]0
# Sum of the logical size of every recorded file
$TotalBytes = [long]0

# Each stack item = (directory, "this directory and its files should be recorded")
$Stack    = [System.Collections.Generic.Stack[System.Tuple[System.IO.DirectoryInfo, bool]]]::New()
$RootDir  = [System.IO.DirectoryInfo]::New($RootPath)
$RootCapt = (-not $UseInclude) -or (Test-DirMatch -Dir $RootDir -Patterns $IncludePatterns)
$Stack.Push([System.Tuple[System.IO.DirectoryInfo, bool]]::New($RootDir, [bool]$RootCapt))

$Stream = $null
$Writer = $null
try
{
    # CreateNew: refuses to overwrite an existing file.
    # FileShare.Read: others may read it.
    $Stream = [System.IO.FileStream]::New(
        $OutFilePath,
        [System.IO.FileMode]::CreateNew,
        [System.IO.FileAccess]::Write,
        [System.IO.FileShare]::Read
    )
    # Unicode (UTF-16 LE with BOM) matches the original
    # 'Export-Csv -Encoding Unicode' output.
    $Writer = [System.IO.StreamWriter]::New(
        $Stream,
        [System.Text.Encoding]::Unicode
    )

    $Writer.WriteLine('"DirectoryName","Name","Type","Extension","SizeInBytes","SizeInKB","Mode","Attributes","CreationTimeUTC","LastAccessTimeUTC","LastWriteTimeUTC","ReparseTarget"')

    while ($Stack.Count -gt 0)
    {
        $Item      = $Stack.Pop()
        $Dir       = $Item.Item1
        $Capturing = $Item.Item2
        $DirCount++

        if ($ProgressTimer.ElapsedMilliseconds -ge $ProgressEveryMs)
        {
            Write-ScanProgress -CurrentDir $Dir.FullName
        }

        # The directory's own row (before its contents)
        if ($Capturing)
        {
            try
            {
                $ParentPath = if ($Dir.Parent)
                {
                    $Dir.Parent.FullName
                }
                else
                {
                    ""
                }
                Write-Row `
                    -Type "Directory" `
                    -DirName $ParentPath `
                    -Item $Dir `
                    -Target (Get-ReparseTarget -Item $Dir)
                $DirRows++
            }
            catch
            {
                $Errors.Add("$( $Dir.FullName ) => $( $_.Exception.Message )")
                Write-ScanWarning -Message ("Could not read $( $Dir.FullName ) => " +
                    "$( $_.Exception.Message )")
            }
        }

        $Files   = [System.Collections.Generic.List[System.IO.FileInfo]]::New()
        $SubDirs = [System.Collections.Generic.List[System.IO.DirectoryInfo]]::New()

        try
        {
            foreach ($Entry in $Dir.EnumerateFileSystemInfos())
            {
                if ($Entry -is [System.IO.DirectoryInfo])
                {
                    $SubDirs.Add($Entry)
                }
                else
                {
                    $Files.Add($Entry)
                }
            }
        }
        catch
        {
            $Errors.Add("$( $Dir.FullName ) => $( $_.Exception.Message )")
            Write-ScanWarning -Message ("Could not fully read $( $Dir.FullName ) => " +
                "$( $_.Exception.Message )")
        }

        # Files in this directory (only if this branch is being captured)
        if ($Capturing -and $Files.Count -gt 0)
        {
            foreach ($File in ($Files | Sort-Object -Property Name))
            {
                try
                {
                    Write-Row `
                        -Type "File" `
                        -DirName $Dir.FullName `
                        -Item $File `
                        -Target (Get-ReparseTarget -Item $File)
                    $FileCount++
                    $TotalBytes +=$File.Length

                    if ($ProgressTimer.ElapsedMilliseconds -ge $ProgressEveryMs)
                    {
                        Write-ScanProgress -CurrentDir $Dir.FullName
                    }
                }
                catch
                {
                    $Errors.Add("$( $File.FullName ) => $( $_.Exception.Message )")
                    Write-ScanWarning -Message ("Could not read $( $File.FullName ) => " +
                        "$( $_.Exception.Message )")
                }
            }
        }

        # Queue subdirectories (pushed in reverse so they pop in alphabetical order)
        foreach ($Sd in ($SubDirs | Sort-Object -Property Name -Descending))
        {
            if ($UseExclude -and (Test-DirMatch -Dir $Sd -Patterns $ExcludePatterns))
            {
                $Skipped.Add("EXCLUDED => $( $Sd.FullName )")
                continue
            }

            $ChildCapt = $Capturing -or ($UseInclude -and (
                Test-DirMatch -Dir $Sd -Patterns $IncludePatterns
                )
            )

            $IsLink = $false
            if ($Sd.Attributes.HasFlag([System.IO.FileAttributes]::ReparsePoint))
            {
                $Target = Get-ReparseTarget -Item $Sd
                # Only junctions and symbolic links are links.  Any other reparse
                # point (for example a cloud-sync placeholder folder) is walked
                # like a normal directory.
                $IsLink = Test-DirIsLink -Item $Sd -Target $Target
            }

            if ($IsLink)
            {
                $SkipReason = $null

                if (-not $FollowReparsePoints)
                {
                    $SkipReason = "Not followed"
                }
                else
                {
                    $Real = Resolve-LinkTargetPath -LinkPath $Sd.FullName -Target $Target
                    if (-not $Real)
                    {
                        $SkipReason = "Not followed => target could not be resolved"
                    }
                    elseif (($Sd.FullName + $Sep).StartsWith(
                            $Real + $Sep,
                            [System.StringComparison]::OrdinalIgnoreCase))
                    {
                        $SkipReason = "Not followed => loop, target contains the link"
                    }
                    elseif (-not $FollowedTargets.Add($Real))
                    {
                        $SkipReason = "Not followed => target already listed through another link"
                    }
                }

                if ($SkipReason)
                {
                    # Recorded as a row (with its target) if it is in capture scope.
                    $Skipped.Add("REPARSEPOINT ($SkipReason) => $( $Sd.FullName ) => $Target")

                    if ($ChildCapt)
                    {
                        try
                        {
                            Write-Row `
                                -Type "Directory" `
                                -DirName $Dir.FullName `
                                -Item $Sd `
                                -Target $Target
                            $DirRows++
                        }
                        catch
                        {
                            $Errors.Add("$( $Sd.FullName ) => $( $_.Exception.Message )")
                            Write-ScanWarning -Message ("Could not read $( $Sd.FullName ) => " +
                                "$( $_.Exception.Message )")
                        }
                    }
                    continue
                }

                # Followed: queued below like any other directory.  Its own
                # row (with the target) is written when it is popped.
                $Followed.Add("$( $Sd.FullName ) => $Target")
            }

            $Stack.Push([System.Tuple[System.IO.DirectoryInfo, bool]]::New(
                $Sd,
                [bool]$ChildCapt
                )
            )
        }
    }
}
finally
{
    Clear-ScanProgress
    if ($Writer)
    {
        # Also closes the underlying stream
        $Writer.Dispose()
    }
    elseif ($Stream)
    {
        $Stream.Dispose()
    }
}


# ------------------------------
#
# Hash the finished CSV and write the manifest
#
# ------------------------------

Show-Message -Text "Directory walk finished." -Color Gray
Show-Message -Text "Hashing the CSV and writing the manifest..." -Color Gray

$CsvHash = try
{
    Get-FileSha256 -Path $OutFilePath
}
catch
{
    "Unavailable => $( $_.Exception.Message )"
}

$CsvSize = try
{
    [System.IO.FileInfo]::New($OutFilePath).Length
}
catch
{
    "Unavailable"
}

$ScriptHash = "N/A (program not run from a script file)"

if ($PSCommandPath)
{
    $ScriptHash = try
    {
        Get-FileSha256 -Path $PSCommandPath
    }
    catch
    {
        "Unavailable => $( $_.Exception.Message )"
    }
}

$Elevated = try
{
    (
        [System.Security.Principal.WindowsPrincipal]`
        [System.Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
        [System.Security.Principal.WindowsBuiltInRole]::Administrator
    )
}
catch
{
    "Unknown"
}


$EndTime = [DateTime]::UtcNow
$Stopwatch.Stop()
$ElapsedTime = Format-Duration -Span $Stopwatch.Elapsed

$Offset = [System.TimeZoneInfo]::Local.GetUtcOffset([DateTime]::UtcNow)

$OffsetText = [string]::Format(
    "UTC{0}{1}",
    $(if ($Offset -lt [TimeSpan]::Zero)
    {
        "-"
    }
    else
    {
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
            "Start Time",
            $StartTime.ToString("o")
        )
    )
)
$Manifest.Add(
    (
        [string]::Format(
            "{0,-21}: {1}",
            "End Time",
            $EndTime.ToString("o")
        )
    )
)
$Manifest.Add(
    (
        [string]::Format(
            "{0,-21}: {1}",
            "Total Scan Time",
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
$Manifest.Add(
    (
        [string]::Format(
            "{0,-21}: {1}",
            "Notice Accepted",
            $NoticeStatus
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
            $(if ($UseInclude)
            {
                $Include -join ", "
            }
            else
            {
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
            $(if ($UseExclude)
            {
                $Exclude -join ", "
            }
            else
            {
                "(none)"
            })
        )
    )
)
$Manifest.Add(
    (
        [string]::Format(
            "{0,-21}: {1}",
            "FollowReparsePoints",
            [bool]$FollowReparsePoints
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
            "Output Folder",
            [System.IO.Path]::GetDirectoryName($OutFilePath)
        )
    )
)

$Manifest.Add(
    (
        [string]::Format(
            "{0,-21}: {1}",
            "Csv File Name",
            [System.IO.Path]::GetFileName($OutFilePath)
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
            [string]::Format(
                $Inv,
                "{0:N0}",
                $DirCount
            )
        )
    )
)

$Manifest.Add("")
$Manifest.Add("")

$Manifest.Add("===== ERRORS ($( $Errors.Count )) =====")

if ($Errors.Count -gt 0)
{
    $Manifest.AddRange($Errors)
}
else
{
    $Manifest.Add("(none)")
}

$Manifest.Add("")
$Manifest.Add("")

$Manifest.Add("===== FOLLOWED REPARSE POINTS ($( $Followed.Count )) =====")

if ($Followed.Count -gt 0)
{
    $Manifest.AddRange($Followed)
}
else
{
    $Manifest.Add("(none)")
}

$Manifest.Add("")
$Manifest.Add("")

$Manifest.Add("===== SKIPPED DIRECTORIES ($( $Skipped.Count )) =====")

if ($Skipped.Count -gt 0)
{
    $Manifest.AddRange($Skipped)
}
else
{
    $Manifest.Add("(none)")
}

$ManifestPath = [System.IO.Path]::Combine(
    $CsvDir,
    "$( $RunDate )_$( $HostName )_$( $SerialPart )manifest.txt"
)

$MStream = $null
$MWriter = $null

# Create the manifest file and stream the data to the file.
try
{
    $MStream = [System.IO.FileStream]::New(
        $ManifestPath,
        [System.IO.FileMode]::CreateNew,
        [System.IO.FileAccess]::Write,
        [System.IO.FileShare]::Read
    )
    $MWriter = [System.IO.StreamWriter]::New(
        $MStream,
        [System.Text.UTF8Encoding]::New($true)
    )
    foreach ($Line in $Manifest)
    {
        $MWriter.WriteLine($Line)
    }
    Show-Message -Text "Writing manifest file complete" -Color Gray
}
catch
{
    Write-Warning "Could not write the manifest file => $( $_.Exception.Message )"
    $ManifestPath = $null
}
finally
{
    if ($MWriter)
    {
        $MWriter.Dispose()
    }
    elseif ($MStream)
    {
        $MStream.Dispose()
    }
}

try
{
    Show-Message -Text "Hashing manifest file..." -Color Gray
    $ManifestHash = Get-FileSha256 -Path $ManifestPath
    Show-Message -Text "Hashing manifest file complete" -Color Gray
}
catch
{
    "Unable to hash the manifest file => $( $_.Exception.Message )"
}


# Write-Host ("`n" + "-" * 50 + "`n") -ForegroundColor Green
Show-Message -Text "Script completed in ", "$( $ElapsedTime )" -Color Gray, Green

Show-Message -Text "Enumerated:" -Color Gray

Show-Message -Text ("File$(if ($FileCount -eq 1) { " " } else { "s" })       :  ",
    "$( [string]::Format($Inv, "{0:N0}", $FileCount) ) ") `
    -Color Gray, Green -NoTime -LinesAbove 1 -Indent $ShortIndent

Show-Message -Text ("Director$(if ($DirRows -eq 1) { "y" } else { "ies" }) :  ",
    "$( [string]::Format($Inv, "{0:N0}", $DirRows) ) ") `
    -Color Gray, Green -NoTime -Indent $ShortIndent

Show-Message -Text ("Total Size  :  ", "$(Get-FormattedFileSize -Bytes $TotalBytes)") `
    -Color Gray, Green -NoTime -Indent $ShortIndent


Show-Message -Text "Output Folder:" -Color Gray -LinesAbove 1
Show-Message -Text "$( [System.IO.Path]::GetDirectoryName($OutFilePath) ) " `
    -Color Green -NoTime -Indent $LargeIndent

Show-Message -Text "Csv File Name:" -Color Gray
Show-Message -Text "$( [System.IO.Path]::GetFileName($OutFilePath) ) " `
    -Color Green -NoTime -Indent $LargeIndent

Show-Message -Text "Csv SHA-256:" -Color Gray
Show-Message -Text "$CsvHash" -Color Green -NoTime -Indent $LargeIndent


if ($ManifestPath)
{
    Show-Message -Text "Manifest File Name:" -Color Gray
    Show-Message -Text "$( [System.IO.Path]::GetFileName($ManifestPath) )" `
        -Color Green -NoTime -Indent $LargeIndent

    Show-Message -Text "Manifest SHA-256:" -Color Gray
    Show-Message -Text "$ManifestHash" -Color Green -NoTime -Indent $LargeIndent
}

Show-Message -Text ("+" + "-" * 48 + "+") -Color Green -NoTime -LinesAbove 1

if ($Errors.Count -gt 0)
{
    Show-Message -Text ("$( $Errors.Count ) " +
        "error$(if ($Errors.Count -eq 1) { " " } else { "s " })=> " +
        "see the manifest for details.`n") -Color Red -LinesAbove 1
}

if ($Stopwatch.IsRunning)
{
    $Stopwatch.Stop()
}
