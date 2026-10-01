# Changelog project: ArcGIS_Online_Member_Tool
---

## Features Heading
- `Added` for new features.
- `Changed` for changes in existing functionality.
- `Fixed` for any bug fixes.
- `Removed` for now removed features.
- `Security` in case of vulnerabilities.
- `Deprecated` for soon-to-be removed features.

[//]: # (Copy paste pallette)
[//]: # (#### Added)
[//]: # (#### Changed)
[//]: # (#### Fixed)
[//]: # (#### Removed)
[//]: # (#### Security)
[//]: # (#### Deprecated)

---

### Version 1.0.0 (2026-10-01)
#### Added
- `Fix-MissingName.ps1` PowerShell script to automatically fill in First Name and Last Name from the student database for users missing name
- call to `Fix-MissingName.ps1` after the student database search
- console feedback confirming the csv import file was successfully updated with user names
- version number display in the Main Menu banner

#### Changed
- cache directory moved from `%TEMP%\cache` to a `cache` folder local to the script directory

#### Fixed
- `FOR /F` loops reading cache files now use `usebackq` with quoted paths, fixing failures when the script path contains spaces

---

### Version 0.10.0 (2025-01-23)
#### Added
- search student database for name information when first name is missing
- variable for student db keyword
- working directory variable
- cache directory in temp

#### Changed
- CD to working directory
  
---


### Version 0.9.0 (2024-09-19)
#### Added
- updatd user types
#### Changed
- default user type to Professional Plus


### Version 0.8.0 (2022-01-05)
#### Changed
- Default ROLE to Publisher


### Version 0.7.0 (2020-09-09)
#### Changed
- Default USER_TYPE to GIS Professional Advanced
- Default log path

#### Added
- Open folder at end
- Delete search file


### Version 0.6.1 (2020-02-14)
#### Changed
- formatting style

### Version 0.6.0 (2020-01-27)
#### Changed
- Search parameters not case sensitive


### Version 0.5.0 (2020-01-24)
#### Added
- data validation for "Role"
- data validation for "USER_TYPE"


### Version 0.4.0 (2020-01-24)
#### Added
- Debugger
- Year autofill
- Search my offerings
#### Changed
- Feedback from advanced search
- Intermediary file cleanup


### Version 0.3.0 (2020-01-23)
#### Added
- Data validation
- OU base search
- Cleanup interim files


### Version 0.2.0 (2020-01-23)
#### Added
- beta release


### Version 0.1.0 (2020-01-21)
#### Added
 - Initial commit