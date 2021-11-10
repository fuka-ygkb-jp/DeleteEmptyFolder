'--------------------------------------------------------------------------
' ��t�H���_�폜�N�P���@�Q�O�O�Q�N�x��
' 2002/07/03 �s�v�c�G�̋�@(http://ygkb.jp/)
'--------------------------------------------------------------------------
'-------------------- ������
Option Explicit

'�萔
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
Const MYNAME = "��t�H���_�폜�N�P��"

'�ϐ�
Dim sCWD			'���g�̃p�X
Dim oFSO			'FileSystemObject
Dim oFolder			'(FolderObject)		�e�t�H���_���g
Dim FolderName		'(FolderObject)		�q�t�H���_���g
Dim oFolderCol1		'(FolderCollection)	�e�t�H���_���̎q�t�H���_�ꗗ
Dim oFolderCol2		'(FolderCollection)	�q�t�H���_���̃t�H���_�ꗗ
Dim oFilesCol		'(FilesCollection)	�q�t�H���_���̃t�@�C���ꗗ
Dim lCounter		'(Long)				���������i�[�p�J�E���^
Dim sTarget			'(String)			�폜�Ώۃt�H���_
Dim oLog			'(FileObject)		���O�t�@�C��
Dim oLogStream		'(TextStreamObject)	�J�������O�t�@�C��

Set oFSO = CreateObject("Scripting.FileSystemObject")
lCounter = 0
sCWD     = oFSO.GetParentFolderName(WScript.ScriptFullName)


'-------------------- �又��
If WScript.Arguments.Count = 0 Then
	MsgBox "�����ΏۂƂ���t�H���_���h���b�v���ĉ������B", vbOKOnly + vbInformation, MYNAME
	WScript.Quit
End If

'Create Logfile
oFSO.CreateTextFile sCWD & "\DeleteEmptyFolder.log"
Set oLog = oFSO.GetFile(sCWD & "\DeleteEmptyFolder.log")
Set oLogStream = oLog.OpenAsTextStream(ForWriting, TristateUseDefault)


'�T�[�`�J�n
Dim Target
for each Target in WScript.Arguments
	ListupEmptyFolder(Target)
next

oLogStream.Close


If lCounter = 0 Then
	MsgBox "�폜�Ώۃt�H���_�������̂ŉ��������I�����܂��B", vbOKOnly + vbInformation, MYNAME
	oLog.Delete
	WScript.Quit
End If


'���ʕ\��
MsgBox "���������B" & vbCRLF & _
       "�폜�Ώۃt�H���_��" & lCounter & "�ł��B", vbOKOnly + vbInformation, MYNAME

'�t�H���_�폜����
If MsgBox("�{���ɍ폜���������s���܂����H", vbYesNo + vbExclamation, MYNAME) = 6 Then
	Set oLogStream = oLog.OpenAsTextStream(ForReading, TristateUseDefault)

	lCounter = 0
	Do While oLogStream.AtEndOfStream <> True
		sTarget = oLogStream.ReadLine

		If sTarget <> "" Or oFSO.FolderExists(sTarget) = True Then	'��sor�����t�H���_�͔�΂�
			oFSO.DeleteFolder sTarget, True
			lCounter = lCounter + 1
		End If
	Loop

	MsgBox "�폜�����B" & vbCRLF & _
	       "�폜���ꂽ�t�H���_��" & lCounter & "�ł��B", vbOKOnly + vbInformation, MYNAME

	oLogStream.Close
End If


'-------------------- �֐�
'sPath�ȉ����ċA�ŒH��Ȃ���A��̃t�H���_�����O�t�@�C���Ƀ��X�g�A�b�v���Ă䂭
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
