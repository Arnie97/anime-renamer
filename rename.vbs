Option Explicit
'On Error Resume Next

Const strApp = "Arnie's Episode Renamer 8.5"
Dim objWSH, objFSO, objTextStream, strPath, strExt, strList, strLeavings, strPattern, strInfo
Set objWSH = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
strExt = "/rm/rmvb/mkv/mp4/avi/mp3/tta/pdf/rar/zip/7z/ass/sub/txt/"
strList = "\Rename.CSV"
strLeavings = "(\.)?(\[|\()(&|_|\.|~|-)*(\)|\])"
strPattern = "(\[|\()([A-F0-9]){6,8}(\)|\])\.(\S+)"
strInfo = "GB|BIG5|JP|CN|\d{3,4}p|\d{3,4}x\d{3,4}|RV10|(x|H|H\.)26\d|(v|x)[2-4]|SumiSora|CASO|POPGO|EMD|Ktxp|MAGI_ATELIER|SOSG|jump|(BT|ZM)PIG|DMG|DmzJ|www\.\S+\.(com|net|org)(\.cn)?|电影天堂|飘花电影|6v电影|&amp;|%26|_BF|_baofeng|XviD|DviX|\[(RMVB|MKV|MP4)\]|(BD|DVD|TV)(Rip)?|flac|aac|ac3|_mp3"

List
Rename

Private Sub List()
	Dim objFile, objFiles, strArgs, strClipboard
	With objFSO
		If WScript.Arguments.Count = 0 Then
			Message(1)
		Else
			strArgs = WScript.Arguments(0)
			If .FolderExists(strArgs) Then
				strPath = strArgs
			ElseIf .FileExists(strArgs) Then
				strPath = .GetParentFolderName(strArgs)
			Else
				Message(1)
			End If
		End If
		strList = .GetParentFolderName(WScript.ScriptFullName) & strList
		If .FileExists(strList) Then
			.DeleteFile(strList)
		End If
		Set objFiles = .GetFolder(strPath).Files
		Set objTextStream = .OpenTextFile(strList, 2, True)
		For Each objFile In objFiles
			If Not InStr(strExt, "/" & .GetExtensionName(objFile.Name) & "/") = 0 Then
				strClipboard = "/" & AutoTrim(objFile.Name)
				objTextStream.WriteLine objFile.Name & ", " & strClipboard
			End If
		Next
		SetClipboard strClipboard
		Set objTextStream = Nothing
	End With
End Sub

Private Sub Rename()
	Dim strLine, strSource, strTarget, intOrder
	With objFSO
		strSource = .GetParentFolderName(WScript.ScriptFullName) & "\Spread32.exe"
		strSource = .GetDriveName(WScript.ScriptFullName) & "\Applets\Applets\Spread32.exe"
		objWSH.Run Chr(34) & strList & Chr(34), 3, True
		Set objTextStream = .OpenTextFile(strList, 1, True)
		Do While Not objTextStream.AtEndOfStream
			strLine = objTextStream.ReadLine
			strLine = Replace(strLine, ", ", "", vbTextCompare)
			intOrder = InStr(strLine, "/")
			strSource = Left(strLine, intOrder - 1)
			strTarget = Right(strLine, Len(strLine) - intOrder)
			strSource = .BuildPath(strPath, strSource)
			If .FileExists(strSource) Then
				If Not .GetFile(strSource).Name = strTarget Then
					.GetFile(strSource).Name = strTarget
				End If
			End If
		Loop
		Set objTextStream = Nothing
		.DeleteFile strList, True
		Message(0)
	End With
End Sub

Private Function AutoTrim(strOriginal)
	Dim objRegExp
	Set objRegExp = New RegExp
	If InStr(AutoTrim, "%") = 0 Then
		AutoTrim = strOriginal
	Else
		Dim intLoop, intChar, strChar
		intLoop = 1
		Do While intLoop <= Len(strOriginal)
			strChar = Mid(strOriginal, intLoop, 1)
			intLoop = intLoop + 1
			If strChar = "%" Then
				intChar = Val("&H" & Mid(strOriginal, intLoop, 2))
				If intChar >= 128 Then
					intChar = intChar * 256 + Val("&H" & Mid(strOriginal, intLoop + 3, 2))
					intLoop = intLoop + 5
				Else
					intLoop = intLoop + 2
				End If
				AutoTrim = AutoTrim & Chr(intChar)
			Else
				AutoTrim = AutoTrim & strChar
			End If
		Loop
	End If
	With objRegExp
		.Global = True
		.IgnoreCase = True
		.Pattern = strInfo
		AutoTrim = .Replace(AutoTrim, "")
		.Pattern = strLeavings
		AutoTrim = .Replace(AutoTrim, "")
		.Pattern = strPattern
		AutoTrim = .Replace(AutoTrim, ".$4")
		.Pattern = "\]\.\["
		AutoTrim = .Replace(AutoTrim, "][")
	End With
End Function

Private Sub SetClipboard(strText)
	Dim objForm, objTextBox
	Set objForm = CreateObject("Forms.Form.1")
	Set objTextBox = objForm.Controls.Add("Forms.TextBox.1").Object
	With objTextBox
		.MultiLine = True
		.Text = strText
		.SelStart = 0
		.SelLength = objTextBox.TextLength
		.Copy
	End With
	Set objTextBox = Nothing
	Set objForm = Nothing
End Sub

Private Sub Message(intMsg)
	Select Case intMsg
		Case 0
			Dim strMsg0
			strMsg0 = "Rename operations sucessfully completed!"
			MsgBox strMsg0, vbOKOnly + vbInformation + vbSystemModal, strApp
		Case 1
			Dim strMsg1
			strMsg1 = "Valid command parameters not found." & vbCrLf & vbCrLf & "Would you like to rename the animations in this script's folder?" & vbCrLf & "Note: Click " & Chr(34) & "No" & Chr(34) & " to cancel renaming and exit." & vbCrLf & vbCrLf & "Usage instructions:" & vbCrLf & vbCrLf & "Go to the folder in which the animations need renaming are, " & vbCrLf & "and drag'n'drop the folder or a file in the folder to this script." & vbCrLf & "Or: Just put the this script file into the folder and run it."
			Select Case MsgBox(strMsg1, vbYesNo + vbInformation + vbSystemModal, strApp)
				Case vbYes
					strPath = objFSO.GetParentFolderName(WScript.ScriptFullName)
					Exit Sub
				Case Else
					WScript.Quit
			End Select
		Case 2
			Dim strMsg2
			strMsg2 = "No animations found in the specified folder." & vbCrLf & vbCrLf & "Would you like to retry listing?" & vbCrLf & "Note: Click " & Chr(34) & "Cancel" & Chr(34) & " to cancel renaming and exit."
			Select Case MsgBox(strMsg2, vbRetryCancel + vbExclamation + vbSystemModal, strApp)
				Case vbRetry
					objWSH.Run WScript.ScriptFullName & Chr(32) & strPath
				Case Else
					WScript.Quit
			End Select
	End Select
End Sub
