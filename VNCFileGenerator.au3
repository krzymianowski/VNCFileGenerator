#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=favicon-2.ico
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#cs ----------------------------------------------------------------------------

	AutoIt Version: 3.3.16.0
	Author:         Tomasz Krzymianowski

	Script Function:
		Generating VNC cnnection files in specific location.

	Parameters:
		\path:""  ;Create VNC files in this path

#ce ----------------------------------------------------------------------------

#include <File.au3>

Global $gFilePath
Global $gEnableLogging = True
Global $gLogPath = @ScriptDir & "\log.log"
Global $ipAdresses
Global $scriptName = "VNC File generator"
Global $scriptVersion = "1.1"
Global $scriptAuthor = "Tomasz Krzymianowski"

_log("---------------------------" & @CRLF)
_log("Running: " & $scriptName & " " & $scriptVersion & @CRLF)
_log("---------------------------" & @CRLF)

; Get parameters from CMD line
If $CmdLine[0] > 0 Then
	For $i = 1 To $CmdLine[0]
		Local $hasPathLine = StringInStr($CmdLine[$i], "/path:")

		If $hasPathLine == 1 And Not $gFilePath Then
			Local $line = StringTrimLeft($CmdLine[$i], 6)
			Local $validPath = _removeLastSlashFromPath($line)
			Local $isValidPath = StringRegExp($validPath, "^(.+)\\([^\/]+)$")

			If $isValidPath Then
				$gFilePath = $validPath
				ContinueLoop
			Else
				_log('[warning]: Specified "path" parameter is not valid. VNC files will be created in script directory.' & @CRLF)
			EndIf
		EndIf
	Next
EndIf

_main()

Func _main()

	If Not $gFilePath Then
		_log('[warning]: Parameter "path" was not specified. VNC files will be created in script directory.' & @CRLF)
		$gFilePath = @ScriptDir
		_log('VNC Files will be created in: ' & $gFilePath & @CRLF)
	EndIf

	_log('Obtaining active IP addresses...' & @CRLF)
	$ipAdresses = _GetAllIP()
	If IsArray($ipAdresses) Then
		_log($ipAdresses[0] & ($ipAdresses[0] > 1 ? " addresses" : " address") & " found" & @CRLF)

		Local $anyValid = False ;

		For $vIP In $ipAdresses
			Local $isValidIP = StringRegExp($vIP, "\b((25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)(\.|$)){4}\b")
			If $isValidIP Then
				$anyValid = True
				_log('Address: ' & $vIP & " is a valid IP address" & @CRLF)
				_CreateVNCFile($vIP)
			EndIf
		Next

		If Not $anyValid Then
			_log('[error]: Any of obtained adresses was valid.' & @CRLF)
			Exit
		EndIf
	Else
		_log('[error]: Failed to obtain any IP address.' & @CRLF)
		Exit
	EndIf
EndFunc   ;==>_main

Func _CreateVNCFile($vIP, $targetDir = $gFilePath)
	Local $userName = @UserName & "(" & @ComputerName & ")(" & $vIP & ")"

	_log('Creatinv VNC File for: ' & $vIP & @CRLF, 2)

	Local $fileName = "\" & $userName & ".vnc"
	Local $file = FileOpen($targetDir & $fileName, 2)

	If Not ($file == -1) Then
		_log('Empty file was created: ' & $targetDir & $fileName & @CRLF, 2)
		_log('Writing data to file...' & @CRLF, 2)

		Local $fileWriteStatus_line_1 = FileWriteLine($file, "FriendlyName=" & $userName)
		Local $fileWriteStatus_line_2 = FileWriteLine($file, "Host=" & $vIP)

		If $fileWriteStatus_line_1 And $fileWriteStatus_line_2 Then
			_log('Successfully written data to the file.' & @CRLF, 2)
			_log('VNC file successfully created!' & @CRLF, 2)
		Else
			_log('[error]: There were some problems writing the data to the file. ' & @CRLF, 2)
		EndIf

		FileClose($file)
	Else
		_log('[error]: Failed to create file for: ' & $vIP & @CRLF)
		Return 0 ;
	EndIf
EndFunc   ;==>_CreateVNCFile

Func _GetAllIP()
	Local $aRet[1] = [0], $iCount = 1, $iNumAddr
	Local $oWMIService = ObjGet("winmgmts:\\.\root\CIMV2")
	If Not IsObj($oWMIService) Then Return SetError(1, 0, 0)

	Local $sQuery = "select * from win32_networkadapterconfiguration where IPEnabled = True"
	Local $oNetCfgs = $oWMIService.ExecQuery($sQuery, "WQL", 0x10 + 0x20)
	If Not IsObj($oNetCfgs) Then Return SetError(1, 0, 0)

	For $oNetCfg In $oNetCfgs
		$aIPAddr = $oNetCfg.IPAddress
		$iNumAddr = UBound($aIPAddr)
		If UBound($aRet) <= ($iCount + $iNumAddr) Then ReDim $aRet[$iCount * 2 + $iNumAddr]
		For $i = 0 To $iNumAddr - 1
			If Not StringRegExp($aIPAddr[$i], "(?:(?:(25[0-5]|2[0-4]\d|1\d{2}|[1-9]\d|\d)\.){3}(?1))") Then ContinueLoop ; IPv4 adresses only
			$aRet[$iCount] = $aIPAddr[$i]
			$iCount += 1
		Next
	Next
	$aRet[0] = $iCount - 1
	ReDim $aRet[$iCount]
	Return $aRet
EndFunc   ;==>_GetAllIP


Func _removeLastSlashFromPath($path)
	Local $lastChar = StringRight($path, 1)

	If $lastChar == "\" Or $lastChar == "/" Then
		Return _removeLastSlashFromPath(StringTrimRight($path, 1))
	Else
		Return $path
	EndIf
EndFunc   ;==>_removeLastSlashFromPath


Func _log($value, $indent = 0, $filePath = $gLogPath)

	; Add indent to value if given
	If $indent > 0 Then
		Local $tempValue = ""
		For $i = 1 To $indent
			$tempValue &= "-"
		Next
		$value = $tempValue & " " & $value
	EndIf

	;Write log to file
	If $gEnableLogging Then
		Local $hFile = FileOpen($filePath, 1)

		If Not ($hFile == -1) Then

			;Write autor in log file if file is empty
			If FileGetSize($filePath) == 0 Then
				FileWriteLine($hFile, $scriptName & " " & $scriptVersion & ", by: " & $scriptAuthor)
				FileWriteLine($hFile, "")
			EndIf

			_FileWriteLog($hFile, $value)
			If @error Then
				ConsoleWrite('[error]: There were some problems writing the data to the log file. ' & @CRLF)
			EndIf

			FileClose($hFile)
		Else
			ConsoleWrite('[error]: Failed to create log file!' & @CRLF)
		EndIf
	EndIf

	;Print log in console
	ConsoleWrite($value & @CRLF)
EndFunc   ;==>_log
