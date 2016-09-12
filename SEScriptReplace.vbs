'###############################################################################
' Info
'###############################################################################
' SE Script Replace
' Version : 2016091200
'
' How to Use : 
'   Run from the command line
'       $ cscript SEScriptReplace.vbs
'   You will need to know how to change directories or folders and be familiar
'       with the command prompt. Most SE files you will modify will be in
'       C:\Users\<YourUserName>\AppData\Roaming\SpaceEngineers
'   Example Change Directory to the above location ^
'       $ cd\Users\<YourUserName\AppData\Roaming\SpaceEngineers
'   Use tab to auto complete, spaces and special characters don't play will with
'       most CLI's.
'
' General Information (algorithm) : 
'   Opens all files recursively from where the VBS script is and searches for
'       the "program element" in the file if it is XML based.
'   It will then check to see if the contents in that element have a match to
'       the strReplaceProgramStartingWith variable below.
'   If there is a match for the string variable the data within element will be
'       replaced with the contents of the file strNewScriptName defaulted to
'       Script.cs, the same file that is extracted from the workshop sbc files.
'   Backup files are NOT made due to VBS default security perms so backup
'       your files before running this!
'   A list of changes are added to the log file named in the strLogFile variable.
'       The default log file is script-updates.log
'   This script can also be ran in the saved games folder to update the ships
'       that currently exist, not sure if they auto update based on the mod.
'   This script will search for existing platformID variables for the IOS ATI
'       SE in game script. As long as the default is "Exmple" it will be replaced.

' Author(s) :
'   Phil Allen phil@hilands.com
' Last Edited By:
'   phil@hilands.com
'
' Copyright (C) 2016  Philip Allen
'   This program is free software: you can redistribute it and/or modify it under the terms of the GNU Lesser General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.
'   This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU Lesser General Public License for more details.
'   You should have received a copy of the GNU Lesser General Public License along with this program.  If not, see <http://www.gnu.org/licenses/>.
'###############################################################################
' Variables To Change
'###############################################################################
' first line of code should suffice. may need to add "line1"&vblf&"line2" etc.
' not neccessarily starting with but contains...
strReplaceProgramStartingWith = "$IOS ATI v2.13"
' file name of the script that will replace old, must be located in same directory
' you run this from.
strNewScriptName="Script.cs"
' place to log changes
strLogFile = "script-updates.log"
'###############################################################################
' Variables for script
'###############################################################################
boolVerbose = True
'###############################################################################
' XML variables and stuff
'###############################################################################
'Dim xmlDoc, CustomerProducts, Products
Set refXML = CreateObject("Msxml2.DOMDocument")
'refXML.setProperty "SelectionLanguage", "XPath"
'###############################################################################
' get the script file and store it in a string
'   Make sure you remove the BOM as it exists in scripts.
'###############################################################################
Set refFSO = CreateObject("Scripting.FileSystemObject")
Set refNewScript = refFSO.OpenTextFile(strNewScriptName)
strNewScript = refNewScript.ReadAll
refNewScript.Close
' check file for BOM and remove it so it's not added to PB <Program> element
'https://gist.github.com/koma5/1899792
Const UTF8_BOM = "﻿"
If Left(strNewScript,3) = UTF8_BOM Then
	strNewScript = Mid(strNewScript,4)
End If
'###############################################################################
' Find all files in the current directory and iterate all sub folders.
' calling the ReplaceScript function
'###############################################################################
Set WshShell = WScript.CreateObject("WScript.Shell")
strCWD = WshShell.CurrentDirectory

Set objFSO = CreateObject("Scripting.FileSystemObject")
objStartFolder = strCWD
Set objFolder = objFSO.GetFolder(objStartFolder)
'Wscript.Echo objFolder.Path
Set colFiles = objFolder.Files
For Each objFile in colFiles
	'Wscript.Echo objFile.Name
	'Wscript.Echo objFolder.Path & "\" & objFile.Name
	if boolVerbose Then
		Wscript.Echo "Parsing " & objFolder.Path & "\" & objFile.Name & vbcrlf
	End If
	ReplaceScript(objFolder.Path & "\" & objFile.Name)
Next
Wscript.Echo
ShowSubfolders objFSO.GetFolder(objStartFolder)
Sub ShowSubFolders(Folder)
	For Each Subfolder in Folder.SubFolders
		'Wscript.Echo Subfolder.Path
		Set objFolder = objFSO.GetFolder(Subfolder.Path)
		Set colFiles = objFolder.Files
		For Each objFile in colFiles
			'Wscript.Echo objFile.Name
			'Wscript.Echo Subfolder.Path & "\" & objFile.Name
			if boolVerbose Then
				Wscript.Echo "Parsing " & Subfolder.Path & "\" & objFile.Name & vbcrlf
			End If
			ReplaceScript(Subfolder.Path & "\" & objFile.Name)
		Next
		'Wscript.Echo
		ShowSubFolders Subfolder
	Next
End Sub
'###############################################################################
' Search the file for <Program> XML elements.
' If it finds them skim the first line of code for a match of the old version of the script
' then replace it with the file we read above.
'###############################################################################
' we should turn this into a sub/function/method and push every file through this.
' add a make changes boolean then decide if changes should be saved.
Function ReplaceScript(strFileName)
	refXML.load(strFileName)
	boolMadeChange = False
	For Each refNode In refXML.SelectNodes("//Program")
		'Wscript.Echo InStr(refNode.text, strReplaceProgramStartingWith)
		'Wscript.Echo refNode.text

		If (InStr(refNode.text, strReplaceProgramStartingWith) <> 0) Then
			If boolVerbose Then
				Wscript.Echo "Found program code to replace"
			End If
			' need to find the platformid and rip the info in the quotes.
			' string platformID = "XMC-718"; 
			'"string platformID"
			'string platformID = " ";
			'string platformID=" ";
			intPlatformIDStart = InStr(strNewScript, "string platformID")
			'Wscript.Echo "platformid Loc: " & InStr(strNewScript, "string platformID")
			'Wscript.Echo "platformid Loc: " & intPlatformIDStart
			'Wscript.Echo Mid(strNewScript,InStr(strNewScript, "string platformID"), 30)
			intPlatformIDStart = intPlatformIDStart + 18
			'Wscript.Echo "idname start loc: " & InStr(intPlatformIDStart, strNewScript, Chr(34))
			intQuoteStart = InStr(intPlatformIDStart, strNewScript, Chr(34))
			intQuoteEnd = InStr(intQuoteStart+1, strNewScript, Chr(34))
			intQuoteLength = intQuoteEnd-intQuoteStart
			'Wscript.Echo intQuoteStart & " " & intQuoteEnd & " " & intQuoteLength
			'Wscript.Echo "pid : " & Mid(strNewScript, intQuoteStart, (intQuoteStart-intQuoteEnd))
			' this gets the first found string platformid quoted stuff.
			'Wscript.Echo Mid(strNewScript, intQuoteStart+1, intQuoteLength-1)
			strPlatformID = Mid(strNewScript, intQuoteStart+1, intQuoteLength-1)
			' need to find Exmple and replace it with strPlatformID.
			strNewScript = Replace(strNewScript, "Exmple", strPlatformID)

			refNode.text = strNewScript
			boolMadeChange = True
		End If
	Next
	if boolMadeChange Then
		'refXML.save(strFileName)
		strLog = TodayDateTime() & " - Modified " & strFileName & vbcrlf
		boolLog = Logger(strLogFile, strLog)
	End If
End Function
'###############################################################################
'# Logger                                                                      #
'#    write log file Function                                                  #
'#    Takes Input of file name then log text                                   #
'###############################################################################
Function Logger (strFile, strText)
	Set refFSO = CreateObject("Scripting.FileSystemObject")
	if refFSO.FileExists(strFile) then
		set refFile = refFSO.OpenTextFile(strFile, 8) ' 8 for append
	else
		set refFile = refFSO.CreateTextFile(strFile, TRUE)
	end if
	refFile.Write(strText)
	refFile.Close
	Set refFile = nothing
	Logger = true
End Function
'###############################################################################
'# Date and Time                                                               #
'###############################################################################
Function TodayDate ()
	Dim tdStrYYYY: tdStrYYYY = Year(Date) ' current year
	Dim tdStrMM: tdStrMM = Month(Date) ' current month
	Dim tdStrDD: tdStrDD = Day(Date) ' current day
	if tdStrMM < 10 then tdStrMM = "0" & tdStrMM end if ' if Month is 1-9 append 0 to front of digit.
	if tdStrDD < 10 then tdStrDD = "0" & tdStrDD end if ' if Day is 1-9 append 0 to front of digit.
	Dim tdStrDate: tdStrDate = tdstrYYYY & tdStrMM & tdStrDD ' store date in format function can understand
	TodayDate = tdStrDate
End Function
Function TodayTime ()
	Dim ttStrHH: ttStrHH = Hour(Now)
	Dim ttStrmin: ttStrmin = Minute(Now)
	Dim ttStrSS: ttStrSS = Second(Now)
	if ttStrHH < 10 then ttStrHH = "0" & ttStrHH end if ' append 0 digit to front
	if ttStrmin < 10 then ttStrmin = "0" & ttStrmin end if ' append 0 digit to front
	if ttStrSS < 10 then ttStrSS = "0" & ttStrSS end if ' append 0 digit to front
	Dim ttStrTime: ttStrTime = ttStrHH & ":" & ttStrmin & ":" & ttStrSS
	TodayTime = ttStrTime
End Function
Function TodayDateTime ()
	tdtStrDate = TodayDate()
	tdtStrTime = TodayTime()
	Dim tdtStrDateTime: tdtStrDateTime = tdtStrDate & " " & tdtStrTime
	TodayDateTime = tdtStrDateTime
End Function