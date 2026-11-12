### Change Log for "Get-DirectoryList_vX.X.ps1"

1. All notable changes to this application will be documented in this file.

2. The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

3. This project adheres to [Semantic Versioning](https://semver.org).

#### [1.2.01] - 2026-11-12

##### Added

- Add warning for examiner regarding cloud-synced folders.

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

- Comment-based help for the overall file.
