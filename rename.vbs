Option Explicit
'On Error Resume Next

Const strInfo = "[GB]/[BIG5]/[Jp/Cn]/[720p]/[480p]/[1280x720]/[848x480]/[RV10]/XviD/V2/V3/amp;/.[CASO&SumiSora]/.[SumiSora&Ktxp]/[]"
Dim objWSH, objFSO
Set objWSH = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

List
Spreadsheet
MsgBox "Rename operations sucessfully completed!", vbOKOnly + vbInformation, "Arnie's TV Animation Renamer"

Private Sub List()
	Dim objExec, objTextStream, strLine, strClipboard
	If objFSO.FileExists("Rename.CMD") Or objFSO.FileExists("Rename.CSV") Then
		objFSO.DeleteFile("Rename.C*")
	End If
	Set objExec = objWSH.Exec("cmd.exe /c dir /b *.mp4 *.mkv *.rmvb *.avi")
	Set objTextStream = objFSO.OpenTextFile("Rename.CSV", ForWriting, True)
	With objTextStream
		.WriteLine "@Echo Off"
		Do While Not objExec.StdOut.AtEndOfStream
			strLine = objExec.StdOut.ReadLine
			strClipboard = AutoTrim(strLine, strInfo)
			strLine = "Ren \," & strLine & ",/," & strClipboard & ",\"
			.WriteLine strLine
			strLine = vbNullString
		Loop
		.WriteLine "Del Rename.CMD"
	End With
	Set objTextStream = Nothing
	SetClipboard strClipboard
End Sub

Private Sub SpreadSheet()
	Dim objTextStream, strAll
	objWSH.Run "Rename.CSV", vbMaximizedFocus, True
	objFSO.MoveFile "Rename.CSV", "Rename.CMD"
	Set objTextStream = objFSO.OpenTextFile("Rename.CMD", ForReading, True)
	strAll = objTextStream.ReadAll
	strAll = Replace(strAll, "\", Chr(34), vbTextCompare)
	strAll = Replace(strAll, "/", Chr(34) & Chr(32) & Chr(34), vbTextCompare)
	strAll = Replace(strAll, ",", vbNullString)
	Set objTextStream = objFSO.OpenTextFile("Rename.CMD", ForWriting, True)
	objTextStream.Write(strAll)
	Set objTextStream = Nothing
	objWSH.Run "Rename.CMD", vbMinimizedNoFocus
End Sub

Private Function AutoTrim(strOriginal)
	Dim strKeyword, strKeywords
	AutoTrim = strOriginal
	strKeyWords = Split(strInfo, "/", -1, vbTextCompare)
	For Each strKeyword In strKeyWords
		AutoTrim = Replace(AutoTrim, strKeyword, "", vbTextCompare)
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
