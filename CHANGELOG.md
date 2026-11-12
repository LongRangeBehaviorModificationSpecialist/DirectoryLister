### Change Log for "Get-DirectoryList_vX.X.ps1"

1. All notable changes to this application will be documented in this file.

2. The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

3. This project adheres to [Semantic Versioning](https://semver.org).

#### [1.2.01] - 2026-11-12

##### Changed

- Move the "UTC" marker to the right of the timestamps in the manifest file.

##### Added

- Add warning for examiner regarding cloud-synced folders.
- Add dedicated function to handle the formatting of the text printed to the console.
- Add 'Examiner Notice' section to the main comment-based help.
- Edited format of a few lines in the manifest file (split the long output path to two seperate lines - one for the directory and one for the file name).
- Add SHA256 hash value for the manifest file to be displayed to screen (but not recorded anywhere).
- Add more output lines to show progress update of the script.

##### Removed

- Remove duplicate variables.

#### [1.2] - 2026-11-11

##### Changed

- N/A

##### Added

- Line to the output screen to show the script's progress. (The progress display only writes to the console. It doesn’t change the CSV contents or touch the examined volume.)
- Add "Progress" section to the help text.
- Add ability for script to follow ReparsePoints and list files and sub-directories.

##### Removed

- N/A

##### Fixed

- Format for numbers in the manifest file (add comma seperator).

##### Deprecated

- N/A

##### Security

- N/A

#### [1.1] - 2026-11-10

##### Added

- Add comment-based help for the overall file.
