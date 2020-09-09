::#############################################################################
:: Name: ArcGIS Online Member Bulk Load Tool
:: Author:		David Geeraerts
:: Location:	Olympia, Washington USA
:: E-Mail:		geeraerd@evergreen.edu
::#############################################################################
:: Copyleft License(s)
:: GNU GPL (General Public License)
:: https://www.gnu.org/licenses/gpl-3.0.en.html
::#############################################################################

::###########################################################################::
::							#DESCRIPTION#
::	ArcGIS Online Member Bulk Load Tool, bulk loads users (students)
::	into ArcGIS Online with a properly formatted csv file
::	created from Offerings, Classes, Student Group, or any other directory 
::	group.
::
::###########################################################################::


@Echo Off
setlocal enableextensions
:: /**************************************************************************/


::###########################################################################::
::	VERSIONING INFORMATION
::	Semantic Versioning used
::		http://semver.org/
::###########################################################################::
::	Major.Minor.Revision
::	Added BUILD number which is used during development and testing.
::###########################################################################::

:: Commandlet
::	Set Console.
	SET SCRIPT_NAME=ArcGIS_Online_Member_Tool
	SET SCRIPT_VERSION=0.7.0
	SET SCRIPT_BUILD=20200909-0745
	Title %SCRIPT_NAME% %SCRIPT_VERSION%
	Prompt AOBU$G
::	Set mode
:mode
	cls
	mode con:cols=72
	mode con:lines=40
	color 0B
:: /**************************************************************************/

:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:: Declare Global variables
::	All User variables are set within here.
::		(configure variables)
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:: CSV File Output
SET "FILE_OUTPUT=%USERPROFILE%\Documents\%SCRIPT_NAME%" 
SET FILE_NAME=ArcGIS_Online_Bulk_Uploader.csv

:: CSV Header format
::	this is hard coded into the ESRI ArcGIS Online bulk member upload
SET "CSV_HEADERS=First Name,Last Name,Email,Username,Role,User Type"

:: DEFAULT Role
::	{Data Editor, Publisher, Student_Publisher, User, Viewer}
SET ROLE=Student_Publisher

:: DEFAULT USER Type
::	{Creator, GIS Professional Advanced}
SET "USER_TYPE=GIS Professional Advanced"


:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
::	ADVANCED SETTINGS
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:: DEFAULT OU for group search
SET OU_DN=OU=offerings,OU=groups,OU=managed,DC=evergreen,DC=edu

:: DEGUGGER
:: Impersonate a User
SET DEBUG_USER=
:: /**************************************************************************/


::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
::
::##### Everything below here is 'hard-coded' [DO NOT MODIFY] #####
::
::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
::	Defaults, not recommended to change in script, rather
::		change during commandlet run time, using menu system.
SET DC=%LOGONSERVER:~2%
IF NOT DEFINED DEBUG_USER SET $USERNAME=%USERNAME%
IF DEFINED DEBUG_USER SET $USERNAME=%DEBUG_USER%
:: /**************************************************************************/

:: Autofills
:: Year
::	contains a trailing space
FOR /F "tokens=4 delims=/ " %%P IN ('date /T') DO SET YEAR=%%P
:: /**************************************************************************/

:: Check log output
IF NOT EXIST "%FILE_OUTPUT%" MD "%FILE_OUTPUT%"
:: /**************************************************************************/	


:Main
	cls
	Echo   ******************************************************************
	Echo.
	Echo                     %SCRIPT_NAME%
	echo.
	Echo   ******************************************************************
	Echo.
	Echo   ******************************************************************
	echo.
	Echo Location: Main Menu
	Echo.
	Echo Current settings
	Echo _____________________
	echo.
	Echo  Log File Path: %FILE_OUTPUT%
	Echo  Log File Name: ^<groupName^>_%FILE_NAME%
	Echo.
	Echo  Running Account: %USERNAME%
	Echo  Domain: %USERDNSDOMAIN%
	Echo  Domain Controller: %DC%
	echo.
	Echo   ******************************************************************
	echo.
	echo Testing connection to domain:
	DSQUERY * -q -limit 1 2> nul || GoTo error00
	echo Success!
	echo.
	Echo.
	Echo Choose an action to perform from the list:
	Echo.
	Echo [1] Search my offerings
	Echo [2] Advanced search for a group
	Echo [3] Set group name, create csv file
	Echo [4] Exit
	Echo.
	Choice /c 1234
	Echo.
	If ERRORLevel 4 GoTo exit
	If ERRORLevel 3 GoTo sGroup
	If ERRORLevel 2 GoTo asearch
	If ERRORLevel 1 GoTo search
:: /**************************************************************************/

:: Selective Search by User
:search
	mode con:cols=72
	mode con:lines=40
	color 0B
echo   ******************************************************************
echo.
echo     Location: Selective offering search         
echo.
echo		User:%$USERNAME%
echo.
echo   ******************************************************************
echo.
:year
	SET CHECKER=0
	IF DEFINED YEAR Echo Year currently set to: %YEAR%
	echo [REQUIRED] Set the academic year as yyyy:
	SET /P YEAR=Academic Year:
	echo Academic Year set to: %YEAR%
	IF NOT DEFINED YEAR GoTo year
	FOR %%P IN (a b c d e f g h i j k l m n o p q r s t u v w x y z) DO (
		echo %YEAR% | FIND /I "%%P" && SET /A CHECKER=CHECKER+1
	)
	IF %CHECKER% NEQ 0 ECHO Year contains alpha character!
	IF %CHECKER% NEQ 0 GoTo year
	echo.
:: /**************************************************************************/
:runUS
	IF EXIST "%FILE_OUTPUT%\%$USERNAME%_%YEAR%_Offering.txt" (
		DEL /F /Q "%FILE_OUTPUT%\%$USERNAME%_%YEAR%_Offering.txt"
	)
	DSQUERY USER -o dn -name %$USERNAME% | DSGET USER -memberof 2> nul | DSGET group -samid 2> nul > "%FILE_OUTPUT%\int_%$USERNAME%_Offering.txt"
	FINDSTR /R /C:"%YEAR%*" "%FILE_OUTPUT%\int_%$USERNAME%_Offering.txt" > "%FILE_OUTPUT%\int_%$USERNAME%_%YEAR%_Offering.txt"
	FOR /F "usebackq tokens=1 delims=_" %%P IN ("%FILE_OUTPUT%\int_%$USERNAME%_%YEAR%_Offering.txt") DO dsquery * %OU_DN% -limit 0 -filter "(&(objectCategory=group)(cn=%%P_STU))" -attr name description displayName >> "%FILE_OUTPUT%\%$USERNAME%_%YEAR%_Offering.txt"
	IF EXIST "%FILE_OUTPUT%\%$USERNAME%_%YEAR%_Offering.txt" echo These are your offerings for %YEAR%:
	IF EXIST "%FILE_OUTPUT%\%$USERNAME%_%YEAR%_Offering.txt" FINDSTR /R /C:"%YEAR%*" "%FILE_OUTPUT%\%$USERNAME%_%YEAR%_Offering.txt" 2> nul
	IF NOT EXIST "%FILE_OUTPUT%\%$USERNAME%_%YEAR%_Offering.txt" ECHO No offerings found!
	echo.
	IF NOT EXIST "%FILE_OUTPUT%\%$USERNAME%_%YEAR%_Offering.txt" PAUSE
	IF NOT EXIST "%FILE_OUTPUT%\%$USERNAME%_%YEAR%_Offering.txt" GoTo Main
:skipUS
:: /**************************************************************************/

:: What to do now?
echo What to do next?
Echo Choose an action to perform from the list:
Echo.
Echo [1] Set the group, create csv file
Echo [2] Advanced search
echo [3] Main menu
echo [4] Exit
Echo.
Choice /c 1234
Echo.
If ERRORLevel 4 GoTo exit
If ERRORLevel 3 GoTo Main
If ERRORLevel 2 GoTo aGroup
If ERRORLevel 1 GoTo sGroup
:: ////////////////////////////////////////////////////////////////////////////


:: Advanced Search
:asearch
	cls
	mode con:cols=72
	mode con:lines=40
	color 0B
	Echo   ******************************************************************
	echo.
	Echo     Location: Group Search           
	echo.
	Echo   ******************************************************************
	echo.
	echo  Setting search parameters:
	echo.

:year
	SET CHECKER=0
	IF DEFINED YEAR Echo Year currently set to: %YEAR%
	echo [REQUIRED] Set the academic year as yyyy:
	SET /P YEAR=Academic Year:
	IF NOT DEFINED YEAR GoTo year
	FOR %%P IN (a b c d e f g h i j k l m n o p q r s t u v w x y z) DO echo %YEAR% | FIND /I "%%P" && SET /A CHECKER=CHECKER+1
	IF %CHECKER% NEQ 0 ECHO Year contains alpha character!
	IF %CHECKER% NEQ 0 SET YEAR=
	IF %CHECKER% NEQ 0 GoTo year
	echo Academic Year set to: %YEAR%
	echo.

:quarter
	SET CHECKER=0
	IF DEFINED QUARTER Echo Quarter currently set to: %QUARTER%
	echo [REQUIRED] Set the academic quarter {*, fall, winter, spring, summer}:
	SET /P QUARTER=Academic Quarter:
	IF NOT DEFINED QUARTER GoTo quarter
	echo %QUARTER% | FIND "*"
	IF %ERRORLEVEL% EQU 0 GoTo skipTQ
	FOR %%P IN (fall winter spring summer Fall Winter Spring Summer) DO IF /I "%%P"=="%QUARTER%" SET /A CHECKER=CHECKER+1
	IF %CHECKER% LEQ 0 ECHO quarter is invalid!
	IF %CHECKER% LEQ 0 SET QUARTER=
	IF %CHECKER% LEQ 0 GoTo quarter
	echo Academic quarter set to: %QUARTER%
:skipTQ

	echo.

:keyterm
	IF DEFINED KEY_TERM Echo Key term currently set to: %KEY_TERM%
	echo [REQUIRED] Set a key search term, i.e. GIS:
	SET /P KEY_TERM=Key term:
	IF NOT DEFINED KEY_TERM GoTo keyterm
	echo %KEY_TERM% | FIND "*"
	IF %ERRORLEVEL% EQU 0 SET KEY_TERM=
	IF %ERRORLEVEL% EQU 0 echo Cannot be "*" wildcard!
	IF %ERRORLEVEL% EQU 0 GoTo keyterm
	echo Search term set to: %KEY_TERM%
	echo.
	:: Convert term to numeric
	IF /I "%QUARTER%"=="*" SET $QUARTER=*
	IF /I "%QUARTER%"=="fall" SET $QUARTER=10
	IF /I "%QUARTER%"=="winter" SET $QUARTER=20
	IF /I "%QUARTER%"=="spring" SET $QUARTER=30
	IF /I "%QUARTER%"=="summer" SET $QUARTER=40
	
	::###########################################
	:: temp DEBUG
	:: ECHO OU_DN: %OU_DN%
	:: echo YEAR$QUARTER: %YEAR%%$QUARTER%
	:: echo KEY_TERM: %KEY_TERM%
	::###########################################
	
:: Search for groups
	dsquery * %OU_DN% -limit 0 -filter "(&(objectCategory=group)(cn=%YEAR%%$QUARTER%*_STU)(displayName=*%KEY_TERM%*))" -attr name description displayName
	dsquery * %OU_DN% -limit 0 -filter "(&(objectCategory=group)(cn=%YEAR%%$QUARTER%*_STU)(displayName=*%KEY_TERM%*))" -attr name description displayName > %FILE_OUTPUT%\AD_Group_Search_Results.txt
	echo.
	:: What to do now?
	echo Search again?
	Echo Choose an action to perform from the list:
	Echo.
	Echo [1] Yes, search again
	Echo [2] Set the group, create csv file
	echo [3] Main menu
	Echo.
	Choice /c 123
	Echo.
	If ERRORLevel 3 GoTo Main
	If ERRORLevel 2 GoTo sGroup
	If ERRORLevel 1 GoTo asearch
GoTo sGroup
:: /**************************************************************************/

:sGroup
	cls
	mode con:cols=72
	mode con:lines=40
	color 0B
	Echo   ******************************************************************
	echo.
	Echo     Location: Set Group name            
	Echo.   
	Echo   ******************************************************************
	echo.
	IF EXIST "%FILE_OUTPUT%\%$USERNAME%_%YEAR%_Offering.txt" echo These are your offerings for %YEAR%:
	IF EXIST "%FILE_OUTPUT%\%$USERNAME%_%YEAR%_Offering.txt" FINDSTR /R /C:"%YEAR%*" "%FILE_OUTPUT%\%$USERNAME%_%YEAR%_Offering.txt" 2> nul
	IF DEFINED YEAR IF DEFINED $QUARTER IF DEFINED KEY_TERM dsquery * %OU_DN% -limit 0 -filter "(&(objectCategory=group)(cn=%YEAR%%$QUARTER%*_STU)(displayName=*%KEY_TERM%*))" -attr name description displayName
	IF EXIST "%FILE_OUTPUT%\AD_Group_Search_Results.txt" FOR /F "skip=1 delims= " %%P IN (%FILE_OUTPUT%\AD_Group_Search_Results.txt) DO SET GROUP_NAME=%%P
	IF EXIST "%FILE_OUTPUT%\%$USERNAME%_%YEAR%_Offering.txt" FOR /F "tokens=1 delims= " %%P IN ('FINDSTR /R /C:"....%YEAR%" "%FILE_OUTPUT%\%$USERNAME%_%YEAR%_Offering.txt"') DO SET GROUP_NAME=%%P
	echo.
	Echo What is the name of the group to use for reference?
	echo This should be an *_STU group with students!
	echo.
	echo current Group name: %GROUP_NAME%
	SET /P GROUP_NAME=Group Name:
	IF NOT DEFINED GROUP_NAME GoTo sGroup
	echo checking group...
	DSQUERY GROUP -o dn %OU_DN% -name %GROUP_NAME% -limit 1 | DSGET GROUP -samid -desc 2> nul
	IF %ERRORLEVEL% NEQ 0 GoTo error10
	echo Group is set to: %GROUP_NAME%
	echo.
:sRole
	IF DEFINED ROLE ECHO Current role: %ROLE%
	SET CHECKER=0
	echo Define Role: ^{Data Editor, Publisher, Student_Publisher, User, Viewer^}
	SET /P ROLE=Role:
	FOR %%P IN ("Data Editor" Publisher Student_Publisher User Viewer) DO IF /I "%ROLE%"=="%%~P" SET /A CHECKER=CHECKER+1
	IF %CHECKER% EQU 0 (ECHO Not a valid entry!) & (SET ROLE=) & (GoTo sRole)
	ECHO Role is set to: %ROLE%
	echo.
:sUT
	IF DEFINED USER_TYPE ECHO Current User_Type: %USER_TYPE%
	SET CHECKER=0
	echo Define User_Type: ^{Creator, GIS Professional Advanced^}
	SET /P User_Type=User_Type:
	FOR %%P IN (Creator "GIS Professional Advanced") DO IF /I "%User_Type%"=="%%~P" SET /A CHECKER=CHECKER+1
	IF %CHECKER% EQU 0 (ECHO Not a valid entry!) & (SET User_Type=) & (GoTo sUT)
	ECHO User_Type is set to: %USER_TYPE%
	echo.
	IF EXIST "%FILE_OUTPUT%\%GROUP_NAME%_%FILE_NAME%" copy /Y "%FILE_OUTPUT%\%GROUP_NAME%_%FILE_NAME%" "%FILE_OUTPUT%\%GROUP_NAME%_%FILE_NAME%.old"
	IF EXIST "%FILE_OUTPUT%\int_%GROUP_NAME%_%FILE_NAME%" DEL /F /Q "%FILE_OUTPUT%\int_%GROUP_NAME%_%FILE_NAME%"
	IF EXIST "%FILE_OUTPUT%\%GROUP_NAME%_%FILE_NAME%" DEL /F /Q "%FILE_OUTPUT%\%GROUP_NAME%_%FILE_NAME%"

:fCSVh
	echo %CSV_HEADERS% > "%FILE_OUTPUT%\int_%GROUP_NAME%_%FILE_NAME%"
	echo.
	ECHO Getting the list of users from %GROUP_NAME%...
	echo Processing...
	::	Token reference
	:: Token 1 samid
	:: Token 2 upn
	:: Token 3 Last NAME
	:: Token 4 First Name

	SETLOCAL EnableDelayedExpansion
	:: old reference
	:: FOR /F "skip=1 tokens=1-4 delims=, " %P IN ('DSQUERY GROUP -NAME %GROUP_NAME% ^| DSGET GROUP -members ^| DSGET USER -display -upn -samid') DO ECHO %S,%R,%Q,%P,%ROLE%,%USER_TYPE%
	
	FOR /F "skip=1 tokens=1-4 delims=, " %%P IN ('DSQUERY GROUP -NAME %GROUP_NAME% ^| DSGET GROUP -members ^| DSGET USER -display -upn -samid') DO ECHO %%S,%%R,%%Q,%%P,%ROLE%,%USER_TYPE% >> %FILE_OUTPUT%\int_%GROUP_NAME%_%FILE_NAME%

	FOR /F "skip=2 delims=" %%F IN ('FIND /I /V ",,Succeeded," "%FILE_OUTPUT%\int_%GROUP_NAME%_%FILE_NAME%"') DO ECHO %%F>> %FILE_OUTPUT%\%GROUP_NAME%_%FILE_NAME%
	SETLOCAL DisableDelayedExpansion
	echo.
	IF EXIST "%FILE_OUTPUT%\%GROUP_NAME%_%FILE_NAME%" ECHO Successfully created formatted csv file: [%FILE_OUTPUT%\%GROUP_NAME%_%FILE_NAME%]
	echo.

	:: Check if there are any members without a first name
	FINDSTR /B /R /C:"," "%FILE_OUTPUT%\%GROUP_NAME%_%FILE_NAME%" 2> nul 1> nul
	IF %ERRORLEVEL% NEQ 0 GoTo skipW
	echo   !!WARNING!! !!WARNING!! !!WARNING!!
	echo.
	echo The following users don't have a first name:
	Echo.
	FINDSTR /R /B /C:"," "%FILE_OUTPUT%\%GROUP_NAME%_%FILE_NAME%"
	echo.
pause

:skipW
:: Process another group?
	echo.
	echo Process another group?
	echo Choose an action to perform from the list:
	echo.
	echo [1] Yes, process another group
	echo [2] Exit
	echo.
	Choice /c 12
	echo.
	If ERRORLevel 2 GoTo exit
	If ERRORLevel 1 GoTo Main
GoTo exit
:: /**************************************************************************/


::###########################################################################::
::                    START ERROR SECTION
::###########################################################################::

:://///////////////////////////////////////////////////////////////////////////
:: ERROR LEVEL 00's (DEPENDENCY ERROR)
:error00
:: Dependency failure
	color 0C
	cls
	echo.
	echo !! FATAL ERROR !!
	echo.
	echo Running this tool requires "RSAT - Remote Server Administrative Tools"!
	echo RSAT is not available on this system!
	echo Depending on the version of Windows, RSAT can be installed in two manners:
	echo https://www.microsoft.com/en-us/download/details.aspx?id=45520
	echo.
	pause
exit
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:error10
:: Group set error, not found
	mode con:cols=50 lines=25
	color 0E
	cls
	echo.
	echo    !! GROUP NOT FOUND !!
	echo.
	echo Try again...
	echo.
	pause
GoTo sGroup
::///////////////////////////////////////////////////////////////////////////::

:exit
	cls
	mode con:cols=55 lines=25
	color 0B
	:: cleanup
	DEL /F /Q "%FILE_OUTPUT%\int_*" 2> nul
	DEL /F /Q "%FILE_OUTPUT%\*.old" 2> nul
	DEL /F /Q "%FILE_OUTPUT%\AD_Group_Search_Results.txt" 2> nul
	echo.
	::	Open folder
	@explorer "%FILE_OUTPUT%"
	echo Developed by:
	echo David Geeraerts {dgeeraerts.evergreen@gmail.com}
	echo.
	echo.
	echo.
	echo Contributors:
	echo.
	echo.
	echo.
	echo.
	echo.
	echo Copyleft License
	echo GNU GPL (General Public License)
	echo https://www.gnu.org/licenses/gpl-3.0.en.html
	echo.
	Timeout /T 20
	ENDLOCAL
Exit