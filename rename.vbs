Option Explicit
'On Error Resume Next

Dim objWSH, objFSO, objTextStream, strPath, strList, strInfo
Set objWSH = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
strPath = objFSO.GetParentFolderName(WScript.ScriptFullName)
strList = strPath & "\Rename.CSV"
strInfo = "GB/BIG5/Jp/Cn/720p/576p/480p/360p/1280x720/1024x576/848x480/840x480/720x480/720x396/624x352/RV10/624X352/640X480/V2/V3/SumiSora/CASO/POPGO/EMD/Ktxp/jump/BTPIG/ZMPIG/DMG/DmzJ/www.dy2018.com/www.dygod.net/µÁ”∞ÃÏÃ√/&amp;/%26/_SP/_BF/_baofeng/x264/H264/H.264/XviD/DviX/DVDRip/DVD/flac/aac/AC3/_mp3/.[&]/[_]/[&]/[]/()"

List
Rename
MsgBox "Rename operations sucessfully completed!", vbOKOnly + vbInformation, "Arnie's TV Animation Renamer"

Private Sub List()
	Dim objFile, objFiles, strArgs
	With objFSO
		If .FileExists(strList) Then
			.DeleteFile(strList)
		End If
		Set objFiles = .GetFolder(strPath).Files
		Set objTextStream = .OpenTextFile(strList, 2, True)
		For Each objFile In objFiles
			If Not InStr("/rm/rmvb/mkv/mp4/avi/mp3/", "/" & _
				.GetExtensionName(objFile.Name) & "/") = 0 Then
				objTextStream.WriteLine objFile.Name & ", /" & AutoTrim(objFile.Name)
			End If
		Next
		Set objTextStream = Nothing
	End With
End Sub

Private Sub Rename()
	Dim strLine, strSource, strTarget, intOrder
	With objFSO
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
		SetClipboard strTarget
		Set objTextStream = Nothing
		.DeleteFile strList, True
	End With
End Sub

Private Function AutoTrim(strOriginal)
	Dim strKeyword, strKeywords, intOrder
	AutoTrim = strOriginal
	strKeyWords = Split(strInfo, "/", -1, vbTextCompare)
	For Each strKeyword In strKeyWords
		Do
			intOrder = InStr(LCase(AutoTrim), LCase(strKeyword))
			If intOrder = 0 Then
				Exit Do
			Else
				AutoTrim = Left(AutoTrim, intOrder - 1) & Right(AutoTrim, _
					Len(AutoTrim) - Len(strKeyword) - intOrder + 1)
			End If
		Loop
	Next
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
