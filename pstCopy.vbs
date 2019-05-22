'VBScript

	'Scripts purpose is to copy PST's from source folder to target folder structure necessary for CommVault PST import
	'If PST file owner set correctly then script is not necessary.
	'Creates log directory and log file if doesn't exist in source directory\pstCopyLogs\pstCopy.Log

'------------------------------------------------------------------------------------------------------------------------------------------------------------
'
' START SCRIPT
'
'------------------------------------------------------------------------------------------------------------------------------------------------------------

'Check if cscript

On Error Resume Next
checkifcscript

Set WshShell = WScript.CreateObject("WScript.Shell")
Sub checkifcscript
  strengine = LCase(Mid(WScript.FullName, InstrRev(WScript.FullName,"\")+1))
  If Not strengine="cscript.exe" Then
	wscript.echo "Use CMD cscript pstmgr.vbs <OPTIONS>"
    WScript.Quit
  End If
End Sub

'Intro Text
WScript.Echo "CommVault Professional Services"
WScript.Echo "Latest Edit: 8/7/2015"
Wscript.echo
WScript.Echo "/? or /H for Usage Instructions"
Wscript.echo
strMyComputer = "."

'Define Variables
Set dateTime = CreateObject("WbemScripting.SWbemDateTime")					
Set objFSO = CreateObject("Scripting.FileSystemObject")								
Set colNamedArguments = WScript.Arguments.Named
SourceFolder = colNamedArguments.Item("Source")
TargetFolder = colNamedArguments.Item("Destination")
'booTransVaultEV = colNamedArguments.Exists("T")
booRegExpression = colNamedArguments.Item("RegExpression")


'************************************************LOGGING**********************************************************************************
'Set logging variables
strDirectory = "pstCopylogs"
strFile = "\pstCopy.log"
CurrentDate = Now

'Check if log directory already exists. If not create it.
If objFSO.FolderExists(strDirectory) Then
	Set objFolder = objFSO.GetFolder(strDirectory)
	Else
	Set objFolder = objFSO.CreateFolder(strDirectory)
End If

'Check if log file already exists. If not create it.
If objFSO.FileExists(strDirectory & strFile) Then
	Set objFolder = objFSO.GetFolder(strDirectory)
	Else
	Set objFile = objFSO.CreateTextFile(strDirectory & strFile)
End If 

set objFile = nothing
set objFolder = Nothing

' ForAppending = 8 ForReading = 1, ForWriting = 2
Const ForAppending = 8

'Open log file for appending
Set objReportFile = objFSO.OpenTextFile (strDirectory & strFile, ForAppending, True)

objReportFile.WriteBlankLines(2)

objReportFile.WriteLine("----------------------------------------------------------------------------------------------------------------------")
objReportFile.WriteLine("							SCRIPT RUN: " & FormatDateTime(CurrentDate, vbGeneralDate))
objReportFile.WriteLine("----------------------------------------------------------------------------------------------------------------------")

objReportFile.WriteBlankLines(2)


'*********************************************ARGUMENTS*******************************************************************************
'Check if /H /? option provided
If colNamedArguments.Exists("H") OR colNamedArguments.Exists("?") Then
  Call Usage()
  WScript.Quit
End If



'Check for valid source and target folders
If objFSO.FolderExists(SourceFolder) Then
	wscript.echo "Source Folder Exists!"
	objReportFile.WriteLine ("Source Folder Exists!")
	wscript.echo
	Else
	wscript.echo "No Source Folder! Quit Script"
	objReportFile.WriteLine ("No Source Folder! Quit Script")
	wscript.quit
End If

If objFSO.FolderExists(TargetFolder) Then
	wscript.echo "Target Folder Exists!"
	wscript.echo
	objReportFile.WriteLine ("Target Folder Exists!")
	Else
			Err.Clear
			Set objFolder = objFSO.CreateFolder(TargetFolder)
			If Err.number <> 0 Then
			wscript.echo "Target Path Creation not possible! Provide valid path"
			objReportFile.WriteLine ("Target Path Creation not possible! Provide valid path")
			DisplayErrorInfo
			wscript.echo TargetFolder
			wscript.quit
			Else
			wscript.echo "Target Folder doesn't exist. Target Folder Created!"
			objReportFile.WriteLine ("Target Folder doesn't exist. Target Folder Created!")
			
		End If
	
End If

 If booRegExpression Then 'Use recursive search
 MoveFiles
 End If

'************************************************MOVE FILES (TransVault)*************************************************************************************

Sub MoveFiles

'Find regular Expression pattern from user input
Set objRegEx = CreateObject("VBScript.RegExp")
objRegEx.Global = True
objRegEx.Pattern = booRegExpression

'Set Variables
Set objFolder = objFSO.GetFolder(SourceFolder)
Set colFiles = objFolder.Files

'Create target path and move file to target path
For Each objFile in colFiles
    strSearchString = objFile.Name
	strSearchStringPath = objFile.Path
    Set colMatches = objRegEx.Execute(strSearchString)  

    For Each strMatch in colMatches
        strFolderMatch = strMatch.Value
		wscript.echo strFolderMatch
        strFolderName = TargetFolder & "\" & strFolderMatch & "\"
        If Not objFSO.FolderExists(strFolderName) Then
            Set objNewFolder = objFSO.CreateFolder(strFolderName)
        End If

    
	'Move to target path
	If Not objFSO.FileExists(strFolderName & strSearchString) Then
    objFSO.MoveFile objFile.Path, strFolderName
	objReportFile.WriteLine("Moved the PST file from:   " & strSearchStringPath & " To: " & strFolderName)
	wscript.echo "Moved the PST file from:   " & strSearchStringPath & " To: " & strFolderName
	End If
	Next
Next

End Sub




'For errors
Sub DisplayErrorInfo

    WScript.Echo "Error:      : " & Err
    WScript.Echo "Error (hex) : &H" & Hex(Err)
    WScript.Echo "Source      : " & Err.Source
    WScript.Echo "Description : " & Err.Description
	objReportFile.WriteLine ("Error:      : " & Err)
	objReportFile.WriteLine ("Error (hex) : &H" & Hex(Err))
	objReportFile.WriteLine ("Source      : " & Err.Source)
	objReportFile.WriteLine ("Description : " & Err.Description)
    Err.Clear

End Sub

'******************************************************************USAGE INSTRUCTIONS***********************************************************************

Sub Usage()

	Wscript.Echo "********************pstCopy.vbs INFO**********************"
	WScript.Echo "Scripts purpose is to move PST's from source folder to target folder structure necessary for CommVault PST import"
	WScript.Echo "If PST file owner set correctly then script is not necessary."
	WScript.Echo "Uses Regular Expression input by user to find  pattern in filename to use for folder structure"
	Wscript.echo "Uses VBScript Regular Expression only for input!"
	WScript.Echo "Creates log directory and log file if doesn't exist in  <directory where script is run>\pstCopyLogs\pstCopy.Log"
	WScript.Echo "Creates target folder if doesn't exist and path is valid"
	WScript.Echo 
	WScript.Echo "   /Source - Source directory where PST's are located. Not recursive! PST's must be in root of directory"
	WScript.Echo "   /Destination - Target main directory where PST's are to be moved to"
	WScript.Echo "   /RegExpression - Input regular expresion to be used"
	Wscript.echo
	wscript.echo "   Example: cscript pstCopy.vbs /RegExpression:^.*(?=(\.pst)) /source:C:\PST\sourcedir /destination:C:\PST\targetdir"
	wscript.echo "   This creates folders structure based on filename. If filename matched username expression would be a valid solution for PST Archive"
End Sub



'------------------------------------------------------------------------------------------------------------------------------------------------------------
'
' END SCRIPT
'
'------------------------------------------------------------------------------------------------------------------------------------------------------------
