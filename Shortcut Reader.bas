Attribute VB_Name = "ShortcutReaderModule"
'This module contains this program's core procedures.
Option Explicit

'This procedure is executed when this program is started and displays the target of the shortcut file specified by the user.
Public Sub Main()
On Error GoTo ErrorTrap
Dim Shell As New Shell32.Shell
Dim ShortcutFileName As String
Dim ShortcutFileO As Shell32.FolderItem
Dim ShortcutFolderO As Shell32.Folder
Dim ShortcutDirectory As String
Dim ShortcutLinkO As Shell32.ShellLinkObject
Dim ShortcutPath As String

   ChDrive Left$(App.Path, InStr(App.Path, ":"))
   ChDir App.Path

   ShortcutPath = InputBox$("Enter the path of a shortcut file:")
 
   If Not ShortcutPath = vbNullString Then
      SplitPath ShortcutPath, ShortcutDirectory, ShortcutFileName
      If ShortcutDirectory = vbNullString Then ShortcutDirectory = CurDir$()
      If ShortcutFileName = vbNullString Then ShortcutFileName = ShortcutPath
       
      Set ShortcutFolderO = Shell.NameSpace(ShortcutDirectory)
      If ShortcutFolderO Is Nothing Then
         MsgBox "The specified directory or drive could not be accessed.", vbExclamation
      Else
         Set ShortcutFileO = ShortcutFolderO.ParseName(ShortcutFileName)
         If ShortcutFileO Is Nothing Then
            MsgBox "The specified shortcut file could not be accessed.", vbExclamation
         Else
            If ShortcutFileO.IsLink() Then
               Set ShortcutLinkO = ShortcutFileO.GetLink()
               MsgBox "Target of shortcut:" & vbCrLf & ShortcutLinkO.Path(), vbInformation
            Else
               MsgBox "The link could not be retrieved from the specified shortcut file.", vbExclamation
            End If
         End If
      End If
   End If
   
EndRoutine:
   Exit Sub
   
ErrorTrap:
   MsgBox Err.Description & vbCr & "Error code: " & CStr(Err.Number), vbExclamation
   Resume EndRoutine
End Sub

'This procedure returns the directory and filename parts of the specified path.
Private Sub SplitPath(ByVal Path As String, ByRef Directory As String, ByRef FileName As String)
Dim Position As Long

   Directory = vbNullString
   FileName = vbNullString
   Position = InStrRev(Path, "\")
   
   If Position > 0 Then
      Directory = Left$(Path, Position)
      FileName = Mid$(Path, Position + 1)
   End If
End Sub


