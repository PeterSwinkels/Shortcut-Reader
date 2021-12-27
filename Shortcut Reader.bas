Attribute VB_Name = "ShortcutReaderModule"
'This module contains the Shortcut Reader.
Option Explicit

'This procedure displays the target of the shortcut file specified by the user.
Public Sub Main()
Dim Directory As String
Dim Folder As Shell32.Folder
Dim FolderItem As Shell32.FolderItem
Dim Path As String
Dim Shell As New Shell32.Shell
Dim ShellLink As Shell32.ShellLinkObject
Dim ShortCut As String

Path = InputBox$("Enter the path of a shortcut file:")
 If Path = Empty Then End

SplitPath Path, Directory, ShortCut
 If Directory = Empty Then Directory = CurDir$
 If ShortCut = Empty Then ShortCut = Path
 
Set Folder = Shell.NameSpace(Directory)
Set FolderItem = Folder.ParseName(ShortCut)
Set ShellLink = FolderItem.GetLink

MsgBox "Target of shortcut:" & vbCrLf & ShellLink.Path, vbInformation
End Sub


'This procedure returns the directory and/or filename parts of the specified path.
Public Sub SplitPath(ByVal Path As String, Optional Directory As String, Optional FileName As String)
Dim Index As Long, NextIndex As Long

Directory = Empty
FileName = Empty
Index = InStr(Path, "\")
NextIndex = 0
 Do
  NextIndex = InStr(Index + 1, Path, "\")
   If NextIndex = 0 Then Exit Do
  Index = NextIndex
 Loop
 
 If Not Index = 0 Then
  Directory = Left$(Path, Index)
  FileName = Mid$(Path, Index + 1)
 End If
End Sub


