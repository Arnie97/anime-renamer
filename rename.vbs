Option Explicit
'On Error Resume Next

Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Const vbHide = 0
Const vbMaximizedFocus = 3
Const vbMinimizedNoFocus = 7
Const vbInfomation = 64

Dim objWSH,objFSO
Set objWSH = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

List
Spreadsheet
MsgBox "Rename operations sucessfully completed!", vbOKOnly + vbInformation, "Arnie's TV Animation Renamer"

Private Sub List()
	Dim objExec,objTextStream,strLine
	If objFSO.FileExists("Rename.CMD") Or objFSO.FileExists("Rename.CSV") Then
		objFSO.DeleteFile("Rename.C*")
	End If
	Set objExec = objWSH.Exec("cmd.exe /c dir /b *.mp4 *.mkv *.rmvb *.avi")
	Set objTextStream = objFSO.OpenTextFile("Rename.CSV", ForWriting, True)
	With objTextStream
		.WriteLine "@Echo Off"
		Do While Not objExec.StdOut.AtEndOfStream
			strLine = objExec.StdOut.ReadLine
			strLine = "Ren \," & strLine & ",/,,\"
			.WriteLine strLine
			strLine = Null
		Loop
		.WriteLine "Del Rename.CMD"
	End With
	Set objTextStream = Nothing
End Sub

Private Sub SpreadSheet()
	Dim objTextStream,strAll
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
