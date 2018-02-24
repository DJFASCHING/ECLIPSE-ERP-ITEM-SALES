;******************************************************************************
; NAME: ECLIPSE ITEM SALES REPORT CONVERTER
; FILENAME: ITEM-SLS-PRISRT-PRODUCT-XX.AU3
; REVISION: 1.0
;
; TYPE: AUTOIT    VERSION:   V 3.3.12.0        DATE: 2018:02:23
;******************************************************************************
; PROGRAM DESCRIPTION:
;
; Reads .TXT format file download from Eclipse Hold File
;
; This version reads source file directly into an array
;
;******************************************************************************
; REPORT PARAMETERS:
;
;
;******************************************************************************
; SOURCEFILE DESCRIPTION:
;
; Eclipse Report Downloaded in ASCII - LF = CR/LF  text format with headings
;
;******************************************************************************
; PROGRAM BEHAVIOR:
; Reads text file into array, cleans  up data and writes to TAB format file
;
;
;******************************************************************************
; CHANGE NOTES:
;
; 2018:02:23 CREATED
; 2018:02:23 IT WORKS!
; 2018:02:23 Added Beginning of data locator loop instead of relying
; on a fixed start position
;
;******************************************************************************
; TO-DO NOTES:
;
;
;******************************************************************************
; PRODUCT SALES REPORT PARAMETERS:
;
; SELECT SALES BY: 				PRODUCT or BILL-TO CUSTOMER
; BILL-TO CUSTOMER: 			SPECIFY
; PERIOD COLUMN HEADINGS:		QTY/SALES/GP/GP%
; TYPE OF COMPARISON:			NONE
; PRIMARY SORT:					PRODUCT
; SECONDARY SORT:				BLANK
; DETAIL LEVEL:					PRODUCT
;
; With the Detail Level as Product. Product Record ID displays on 3rd line
;
;***********************************************************************************
;------------------------------------------------------------------------------------
;COL	EXCEL	FIELDNAME							VARNAME             POS		LEN
;------------------------------------------------------------------------------------
;001	A		PRODUCT DESCRIPTION					$sA1				1		35
;002	B		ECLID								$sB1  3rd line		15		10
;003	C		QTY									$sC1				37		12
;004	D		SALES								$sD1				50		15
;005	E		COGS GP								$sE1				66		15
;006	F		GP%									$sF1				82		8
;***********************************************************************************
;***********************************************************************************
; USER DEFINED FUNCTIONS (INCLUDES):

;FOR USER INTERACTION:
#include <MsgBoxConstants.au3>
;FOR GUI:
#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
;FOR SQLITE DB:
;#include <SQLite.au3>
;#include <SQLite.dll.au3>
;FOR FILE OPERATIONS:
#include <File.au3>  ;File O/R/W Ops
#include <FileConstants.au3>  ;File Data
;FOR ARRAY / FILE EXPORT FUNCTIONS:
#include <Array.au3>


;******************************************************************************
; INITIALIZATION:
;******************************************************************************

;GLOBALS:
Global $LogPath = @ScriptDir & '\CMDLOG.TXT'  ;gets rewritten later in script

;LOCALS:
Local $sDrive = "", $sDir = "", $sFileName = "", $sExtension = "",  $iFileExists = ""

;KILLSWITCH:
HotKeySet('{NUMPADSUB}', 'Hotkey1')   ;Number Pad Subtract Key
;MsgBox(0, 'KILLSWITCH', 'TERMINATE WITH NUMPAD - ',2)

;*********************************MAIN PROGRAM*********************************
#Region ### START Koda GUI section ### Form=E:\DATA - Copy\_SCRIPTS\AUTOIT\KODA-FORMS\FORM-FILECONV-A3.kxf
$MAIN = GUICreate("MAIN", 580, 400, 278, 148)

$mLbTitle = GUICtrlCreateLabel("ECLIPSE ITEM SALES REPORT CONVERSION", 24, 16, 300, 17)
GUICtrlSetFont(-1, 8, 800, 4, "MS Sans Serif")

$mLbDesc = GUICtrlCreateLabel("CONVERT XXX.TXT TO .TAB FORMAT FOR EXCEL IMPORT", 24, 48, 400, 17)
GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")

$mBSelect = GUICtrlCreateButton("SELECT", 48, 88, 75, 25)

$mBConvert = GUICtrlCreateButton("CONVERT", 232, 88, 75, 25)

$mBLog = GUICtrlCreateButton("LOG", 424, 88, 75, 25)

$mLbSource = GUICtrlCreateLabel("SOURCE FILE", 40, 144, 85, 17)
GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")
$mISource = GUICtrlCreateInput("SOURCE FILE", 133, 144, 425, 21)

$mLbConvert = GUICtrlCreateLabel("CONVERTED FILE", 16, 192, 110, 17)
GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")
$mIConvert = GUICtrlCreateInput("CONVERTED FILE", 133, 190, 425, 21)

$mLbLog = GUICtrlCreateLabel("ERROR LOG", 48, 240, 77, 17)
GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")
$mILog = GUICtrlCreateInput("LOG FILE", 133, 238, 425, 21)

$mLbLines = GUICtrlCreateLabel("LINES / PROC", 32, 288, 88, 17)
GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")
$mILines = GUICtrlCreateInput("LINES", 133, 288, 105, 21)

$mLbStat = GUICtrlCreateLabel("STATUS", 256, 288, 53, 17)
GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")
$mIStat = GUICtrlCreateInput("STATUS", 316, 287, 177, 21)

$mLbNote = GUICtrlCreateLabel("NOTE", 80, 336, 38, 17)
GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")
$mINote = GUICtrlCreateInput("NOTE", 133, 335, 361, 21)

GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###


GUICtrlSetData($mINote, "[NUMPAD]  -  TO TERMINATE")  ;Display note in GUI

While 1
	$nMsg = GUIGetMsg()  ;Idles CPU when there is no events waiting
	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			Exit
		Case $mBSelect
			$sFilePath = GetFile()
			Local $iFileExists = FileExists($sFilePath)
			If $iFileExists Then
				Local $aPathSplit = _PathSplit($sFilePath, $sDrive, $sDir, $sFileName, $sExtension)
				;_ArrayDisplay($aPathSplit, "_PathSplit of " & $sFilePath) ;Displays array contents
				Local $nFilePath = $sDrive & $sDir & $sFileName & "-tab.txt"
				Local $dFilePath = StringUpper($nFilePath)
				;LogFile Creation / Open
				$LogPath = $sDrive & $sDir & $sFileName & "-log.txt"  ;Modify Log file path
				$LogPath = StringUpper($LogPath)
				;Global $hLogOpen = FileOpen($LogPath, $FO_OVERWRITE)
				;If $hLogOpen = -1 Then
				;	MsgBox($MB_SYSTEMMODAL, "", "An error occurred when opening the LOG file.")
				;EndIf
				GUICtrlSetData($mISource, $sFilePath)  ;GUI Info
				GUICtrlSetData($mIConvert, $dFilePath)  ;GUI Info
				GUICtrlSetData($mILog, $LogPath)  ;GUI Info
			EndIf
		Case $mBConvert
			If $iFileExists Then
				FileToArray($sFilePath)
				RWFile($sFilePath,$dFilePath)
			Else
				MsgBox($MB_SYSTEMMODAL, "", "SELECT SOURCE FILE FIRST", 2)
			EndIf
		Case $mBLog
			ShellExecute("Notepad.exe", $LogPath)

	EndSwitch
WEnd

;*********************************MAIN PROGRAM END******************************


;********************************FUNCTIONS SECTION******************************
;KILLSWITCH:

Func Hotkey1()
     MsgBox(0, 'EXIT', 'PROGRAM TERMINATED',2)
	 ;Close Database
	 ;_SQLite_Close($hDskDb)  ;Close Opened DB file
	; Close the opened Source Data File.
    ;FileClose($hFileOpen)
	 Exit
EndFunc     ;==>Hotkey1

;------------------------------------------------------------------------------

Func GetFile()   ; Select File Dialog
    ; Create a constant variable in Local scope of the message to display in FileOpenDialog.
    Local Const $sMessage = "Select a single file of any type."

    ; Display an open dialog to select a file.
	Local $sFileOpenDialog = FileOpenDialog($sMessage, "C:\TEMP", "All (*.*)", $FD_FILEMUSTEXIST)

    If @error Then
        ; Display the error message.
        MsgBox($MB_SYSTEMMODAL, "", "No file was selected.")
		Exit

        ; Change the working directory (@WorkingDir) back to the location of the script directory as FileOpenDialog sets it to the last accessed folder.
        FileChangeDir(@ScriptDir)
    Else
        ; Change the working directory (@WorkingDir) back to the location of the script directory as FileOpenDialog sets it to the last accessed folder.
        FileChangeDir(@ScriptDir)

        ; Replace instances of "|" with @CRLF in the string returned by FileOpenDialog.
        $sFileOpenDialog = StringReplace($sFileOpenDialog, "|", @CRLF)

        ; Display the selected file.
        ;MsgBox($MB_SYSTEMMODAL, "", "You chose the following file:" & @CRLF & $sFileOpenDialog)
    EndIf
	Return $sFileOpenDialog ;Return Value
EndFunc   ;==>GetFile

;------------------------------------------------------------------------------

Func FileToArray($sFilePath)
    ; Read the current script file into an array using the filepath.
    Global $aArray = FileReadToArray($sFilePath)
    If @error Then
        MsgBox($MB_SYSTEMMODAL, "", "There was an error reading the file. @error: " & @error) ; An error occurred reading the current script file.

;	Else
;        For $i = 0 To UBound($aArray) - 1 ; Loop through the array.
;            MsgBox($MB_SYSTEMMODAL, "", $aArray[$i]) ; Display the contents of the array.
;        Next
    EndIf
EndFunc   ;==>FileToArray


;------------------------------------------------------------------------------
;NOT USED UNLESS FILE IS TOO LARGE FOR AN ARRAY (REFERENCE ONLY)
Func FileRdLine($sFilePath)

    ; Open the file for reading and store the handle to a variable.
    Local $hFileOpen = FileOpen($sFilePath, $FO_READ)
    If $hFileOpen = -1 Then
        MsgBox($MB_SYSTEMMODAL, "", "An error occurred when reading the file.")
        Return False
    EndIf

    ; Read the fist line of the file using the handle returned by FileOpen.
    Local $sFileRead = FileReadLine($hFileOpen, 2) ;Read line 2 of file (L1 is header)

    ; Close the handle returned by FileOpen.
    FileClose($hFileOpen)

    ; Display the first line of the file.
    MsgBox($MB_SYSTEMMODAL, "", "First line of the file:" & @CRLF & $sFileRead)
	Return
EndFunc   ;==>FileRdLine

;------------------------------------------------------------------------------
;Read Source File Array and Write to Result file
Func RWFile($sFilePath, $dFilePath)

	;Set Progress Counters
	$iCtr1 = 0
	$iCtr2 = 0
	GUICtrlSetData($mILines, $iCtr2)  ;Set Initial Progress in GUI
	GUICtrlSetData($mIStat, "RUNNING")  ;Set Initial Status in GUI

	;Open Error log file
	Local $hLogOpen = FileOpen($LogPath, $FO_OVERWRITE)
	If $hLogOpen = -1 Then
		MsgBox($MB_SYSTEMMODAL, "", "An error occurred when opening the LOG file.")
	EndIf

	FileWriteLine($hLogOpen, "ERROR LOG CONTENTS: ")

    ; Create or open file for writing
    ;Local $hFileOpenW = FileOpen($sFilePath, $FO_APPEND)
	Local $hFileOpenW = FileOpen($dFilePath, $FO_OVERWRITE)
    If $hFileOpenW = -1 Then
        MsgBox($MB_SYSTEMMODAL, "", "An error occurred when opening the file.")
        Return False
    EndIf

	;Write Header
	$sWrt = "ECLID" & @TAB
	$sWrt = $sWrt & "DESCRIPTION" & @TAB
	$sWrt = $sWrt & "QTY" & @TAB
	$sWrt = $sWrt & "SALES" & @TAB
	$sWrt = $sWrt & "COGS-GP" & @TAB
	$sWrt = $sWrt & "GP-PCT"

	;Write Header to file
	FileWriteLine($hFileOpenW, $sWrt)

	;Declare serveral variables in advance (for syntax check)
	Local $sMrk1, $sMrk2, $sA1, $sB1, $sC1, $sD1, $sE1, $sF1

	$iEnd = UBound($aArray) - 1

	;Find end of header lines
	$i = 0
	While $i <= $iEnd
		$sMrk1 = StringMid ($aArray[$i],1 ,19) ; End of header lines
		If $sMrk1 == "Product Description" Then
			$iStart = $i
			ExitLoop
		EndIf
		$i = $i + 1
	WEnd


	For $i = $iStart to $iEnd
		;MsgBox($MB_SYSTEMMODAL, "", $aArray[$i]) ; Display the contents of the array.


		;Evaluate valid format and position marker
		$sMrk2 = StringMid ($aArray[$i],6 ,7) ; "Product"
		;$sMrk2 = StringStripWS ($sMrk2 , 2) ;Remove trailing whitespace
		;$sMrk2 = StringStripWS ($sMrk2, 1) ;Remove leading whitespace
		;$sMrk2 = StringReplace($sMrk2, ",","")  ;Remove commas

		;003	C		QTY									$sC1				37		12
		$sC1 = StringMid ($aArray[$i],37 ,12)  ;
		;$sC1 = StringStripWS ($sC1, 2) ;Remove trailing whitespace
		$sC1 = StringStripWS ($sC1, 1) ;Remove leading whitespace
		$sC1 = StringReplace($sC1, ",","")  ;Remove commas

		If $sC1 <> "" Then

			If $sMrk2 == "Product" Then

				;002	B		ECLID								$sB1  3rd line		15		10
				$sB1 = StringMid ($aArray[$i],15 ,10)
				$sB1 = StringStripWS ($sB1, 2) ;Remove trailing whitespace
				;$sB1 = StringStripWS ($sB1, 1) ;Remove leading whitespace
				$sB1 = StringReplace($sB1, ",","")  ;Remove commas

				;Write Line to Result File
				$sWrt = $sB1 & @TAB
				$sWrt = $sWrt & $sA1 & @TAB
				$sWrt = $sWrt & $sC1 & @TAB
				$sWrt = $sWrt & $sD1 & @TAB
				$sWrt = $sWrt & $sE1 & @TAB
				$sWrt = $sWrt & $sF1

				FileWriteLine($hFileOpenW, $sWrt)

			Else

				;001	A		PRODUCT DESCRIPTION					$sA1				1		35
				$sA1 = StringMid ($aArray[$i],1 ,35)
				$sA1 = StringStripWS ($sA1, 2) ;Remove trailing whitespace
				;$sA1 = StringStripWS ($sA1, 1) ;Remove leading whitespace
				;$sA1 = StringReplace($sA1, ",","")  ;Remove commas

				;004	D		SALES								$sD1				50		15
				$sD1 = StringMid ($aArray[$i],50 ,15)
				;$sD1 = StringStripWS ($sD1, 2) ;Remove trailing whitespace
				$sD1 = StringStripWS ($sD1, 1) ;Remove leading whitespace
				;$sD1 = StringReplace($sD1, ",","")  ;Remove commas

				;005	E		COGS GP								$sE1				66		15
				$sE1 = StringMid ($aArray[$i],66 ,15)
				;$sE1 = StringStripWS ($sE1, 2) ;Remove trailing whitespace
				$sE1 = StringStripWS ($sE1, 1) ;Remove leading whitespace
				;$sE1 = StringReplace($$sE1, ",","")  ;Remove commas

				;006	F		GP%									$sF1				82		8
				$sF1 = StringMid ($aArray[$i],82 ,8)
				;$sF1 = StringStripWS ($sF1, 2) ;Remove trailing whitespace
				$sF1 = StringStripWS ($sF1, 1) ;Remove leading whitespace
				;$sF1 = StringReplace($$sF1, ",","")  ;Remove commas
			EndIf

		Else
			;Skipped lines written to log
			FileWriteLine($hLogOpen, $aArray[$i])

		EndIf  ;End If $sMrk2 == "Product"


		;Update Progress counter
		$iCtr1 = $iCtr1 + 1

		;Progress counter
		If $iCtr1 == 50 Then
			$iCtr2 = $iCtr2 + $iCtr1
			GUICtrlSetData($mILines, $iCtr2)  ;Display Progress in GUI
			$iCtr1 = 0
		EndIf

	Next ;For $i


	;Final counter update
	$iCtr2 = $iCtr2 + $iCtr1
	GUICtrlSetData($mILines, $iCtr2)  ;Display Progress in GUI
	GUICtrlSetData($mIStat, "DONE")  ;Set Status in GUI

    ; Read the fist line of the file using the handle returned by FileOpen.
    ;Local $sFileRead = FileReadLine($hFileOpenR, 2) ;Read line 2 of file (L1 is header)

    ; Display the first line of the file.
    ;MsgBox($MB_SYSTEMMODAL, "", "First line of the file:" & @CRLF & $sFileRead)

	; Close the handle returned by FileOpen.
	;Close Destination Write File
	FileClose($hFileOpenW)
	;Close the log file
	FileClose($hLogOpen)

	Return
EndFunc   ;==>FileRdLine

;------------------------------------------------------------------------------



