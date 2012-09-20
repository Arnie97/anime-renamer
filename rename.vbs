Option Explicit
'On Error Resume Next

Const strApp = "Arnie's Episode Renamer 8.2"
Dim objWSH, objFSO, objTextStream, strPath, strList, strPattern, strInfo
Set objWSH = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
strList = "\Rename.CSV"
strPattern = "(\[|\()([A-F0-9]){6,8}(\)|\])\.(\S+)"
strInfo = "GB/BIG5/Jp/Cn/720p/576p/480p/360p/1280x720/1024x576/848x480/840x480/720x480/720x396/624x352/RV10/624X352/640X480/V2/V3/SumiSora/CASO/POPGO/EMD/Ktxp/MAGI_ATELIER/jump/BTPIG/ZMPIG/DMG/DmzJ/www.dy2018.com/www.dygod.net/电影天堂/www.piaohua.com/飘花电影/&amp;/%26/_SP/_BF/_baofeng/x264/H264/H.264/XviD/DviX/DVDRip/DVD/_flac/_aac/_ac3/_mp3/&/.[_]/[_]/.[]/[]/()"

List
Rename

Private Sub List()
	Dim objFile, objFiles, strArgs, strClipboard
	With objFSO
		strList = .GetParentFolderName(WScript.ScriptFullName) & strList
		If .FileExists(strList) Then
			.DeleteFile(strList)
		End If
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
		Set objFiles = .GetFolder(strPath).Files
		Set objTextStream = .OpenTextFile(strList, 2, True)
		For Each objFile In objFiles
			If Not InStr("/rm/rmvb/mkv/mp4/avi/mp3/pdf/rar/zip/7z/ass/", "/" & .GetExtensionName(objFile.Name) & "/") = 0 Then
				strClipboard = AutoTrim(objFile.Name)
				objTextStream.WriteLine objFile.Name & ", /" & strClipboard
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
	Dim objRegExp, strKeyword, strKeywords, intOrder
	Set objRegExp = New RegExp
	AutoTrim = strOriginal
	strKeyWords = Split(strInfo, "/",  - 1, vbTextCompare)
	For Each strKeyword In strKeyWords
		Do
			intOrder = InStr(LCase(AutoTrim), LCase(strKeyword))
			If intOrder = 0 Then
				Exit Do
			Else
				AutoTrim = Left(AutoTrim, intOrder - 1) & Right(AutoTrim, Len(AutoTrim) - Len(strKeyword) - intOrder + 1)
			End If
		Loop
	Next
	With objRegExp
		.Global = True
		.IgnoreCase = True
		.Pattern = strPattern
		AutoTrim = .Replace(AutoTrim, ".$4")
	End With
	AutoTrim = Replace(AutoTrim, "].[", "][", vbTextCompare)
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

Private Function Hex2ASCII(strHex)
	Dim strChr, intLoop, intChr
	strHex = UCase(strHex)
	For intLoop = 1 To Len(strHex)
		strChr = Mid(strHex, Len(strHex) - intLoop + 1, 1)
		Select Case strChr
			Case "A", "B", "C", "D", "E", "F"
				intChr = intChr + (Asc(strChr) - 65 + 10) * 16 ^ (intLoop - 1)
			Case 1, 2, 3, 4, 5, 6, 7, 8, 9, 0
				intChr = intChr + strChr * 16 ^ (intLoop - 1)
			Case Else
		End Select
	Next
	Hex2ASCII = ChrW(intChr)
End Function
