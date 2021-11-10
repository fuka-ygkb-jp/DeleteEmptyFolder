'--------------------------------------------------------------------------
' 空フォルダ削除君１号　２００２年度版
' 2002/07/03 不可思議絵の具　(http://ygkb.jp/)
'--------------------------------------------------------------------------
'-------------------- 初期化
Option Explicit

'定数
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
Const MYNAME = "空フォルダ削除君１号"

'変数
Dim sCWD			'自身のパス
Dim oFSO			'FileSystemObject
Dim oFolder			'(FolderObject)		親フォルダ自身
Dim FolderName		'(FolderObject)		子フォルダ自身
Dim oFolderCol1		'(FolderCollection)	親フォルダ内の子フォルダ一覧
Dim oFolderCol2		'(FolderCollection)	子フォルダ内のフォルダ一覧
Dim oFilesCol		'(FilesCollection)	子フォルダ内のファイル一覧
Dim lCounter		'(Long)				処理件数格納用カウンタ
Dim sTarget			'(String)			削除対象フォルダ
Dim oLog			'(FileObject)		ログファイル
Dim oLogStream		'(TextStreamObject)	開いたログファイル

Set oFSO = CreateObject("Scripting.FileSystemObject")
lCounter = 0
sCWD     = oFSO.GetParentFolderName(WScript.ScriptFullName)


'-------------------- 主処理
If WScript.Arguments.Count = 0 Then
	MsgBox "検索対象とするフォルダをドロップして下さい。", vbOKOnly + vbInformation, MYNAME
	WScript.Quit
End If

'Create Logfile
oFSO.CreateTextFile sCWD & "\DeleteEmptyFolder.log"
Set oLog = oFSO.GetFile(sCWD & "\DeleteEmptyFolder.log")
Set oLogStream = oLog.OpenAsTextStream(ForWriting, TristateUseDefault)


'サーチ開始
Dim Target
for each Target in WScript.Arguments
	ListupEmptyFolder(Target)
next

oLogStream.Close


If lCounter = 0 Then
	MsgBox "削除対象フォルダが無いので何もせず終了します。", vbOKOnly + vbInformation, MYNAME
	oLog.Delete
	WScript.Quit
End If


'結果表示
MsgBox "検索完了。" & vbCRLF & _
       "削除対象フォルダは" & lCounter & "個です。", vbOKOnly + vbInformation, MYNAME

'フォルダ削除処理
If MsgBox("本当に削除処理を実行しますか？", vbYesNo + vbExclamation, MYNAME) = 6 Then
	Set oLogStream = oLog.OpenAsTextStream(ForReading, TristateUseDefault)

	lCounter = 0
	Do While oLogStream.AtEndOfStream <> True
		sTarget = oLogStream.ReadLine

		If sTarget <> "" Or oFSO.FolderExists(sTarget) = True Then	'空行or無いフォルダは飛ばす
			oFSO.DeleteFolder sTarget, True
			lCounter = lCounter + 1
		End If
	Loop

	MsgBox "削除完了。" & vbCRLF & _
	       "削除されたフォルダは" & lCounter & "個です。", vbOKOnly + vbInformation, MYNAME

	oLogStream.Close
End If


'-------------------- 関数
'sPath以下を再帰で辿りながら、空のフォルダをログファイルにリストアップしてゆく
Function ListupEmptyFolder(sPath)
	Set oFolder     = oFSO.GetFolder(sPath)
	Set oFolderCol1 = oFolder.SubFolders

	Dim Folder
	For Each Folder in oFolderCol1
		Set oFolderCol2 = Folder.SubFolders
		Set oFilesCol   = Folder.Files
		If (oFolderCol2.Count = 0) And (oFilesCol.Count = 0) Then
			lCounter = lCounter + 1
			oLogStream.WriteLine sPath & "\" & Folder.name
		end if
		ListupEmptyFolder(sPath & "\" & Folder.name)
	Next
End Function
