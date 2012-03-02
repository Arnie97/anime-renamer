Option Explicit
'On Error Resume Next

Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Const vbHide = 0
Const vbMaximizedFocus = 3
Const vbMinimizedNoFocus = 7
Const vbInfomation = 64
Const strInfo = "[GB]/[BIG5]/[Jp/Cn]/[720p]/[480p]/[1280x720]/[848x480]/[RV10]/XviD/V2/V3/amp;/.[CASO&SumiSora]/.[SumiSora&Ktxp]/[]"
Dim objWSH,objFSO
Set objWSH = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

List
Spreadsheet
MsgBox "Rename operations sucessfully completed!", vbOKOnly + vbInformation, "Arnie's TV Animation Renamer"

Private Sub List()
	Dim objExec,objTextStream,strLine,strClipboard
	If objFSO.FileExists("Rename.CMD") Or objFSO.FileExists("Rename.CSV") Then
		objFSO.DeleteFile("Rename.C*")
	End If
	Set objExec = objWSH.Exec("cmd.exe /c dir /b *.mp4 *.mkv *.rmvb *.avi")
	Set objTextStream = objFSO.OpenTextFile("Rename.CSV", ForWriting, True)
	With objTextStream
		.WriteLine "@Echo Off"
		Do While Not objExec.StdOut.AtEndOfStream
			strLine = objExec.StdOut.ReadLine
			strClipboard = strClipboard & AutoTrim(strLine, strInfo) & vbCrLf
			strLine = "Ren \," & strLine & ",/,,\"
			.WriteLine strLine
			strLine = Null
		Loop
		.WriteLine "Del Rename.CMD"
	End With
	Set objTextStream = Nothing
	SetClipboard strClipboard
End Sub

Private Sub SpreadSheet()
	Dim objTextStream,strAll
	objWSH.Run "Rename.CSV", vbMaximizedFocus, True
	objFSO.MoveFile "Rename.CSV", "Rename.CMD"
	Set objTextStream = objFSO.OpenTextFile("Rename.CMD", ForReading, True)
	strAll = objTextStream.ReadAll
	strAll = Replace(strAll, "\", Chr(34), vbTextCompare)
	strAll = Replace(strAll, "/", Chr(34) & Chr(32) & Chr(34), vbTextCompare)
	strAll = AutoTrim(strAll, ",")
	Set objTextStream = objFSO.OpenTextFile("Rename.CMD", ForWriting, True)
	objTextStream.Write(strAll)
	Set objTextStream = Nothing
	objWSH.Run "Rename.CMD", vbMinimizedNoFocus
End Sub

Private Function AutoTrim(strSource, strReplace)
	Dim strKeyword,strKeywords
	AutoTrim = strSource
	strKeyWords = Split(strReplace, "/", -1, vbTextCompare)
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
