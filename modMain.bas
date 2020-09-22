Attribute VB_Name = "modMain"
Option Explicit

'this sub fires before anything else and checks to see if unzip32.dll is
'installed, if not it asks if it should auto install it from the included res file
'then it displays the main form
Sub Main()

Dim strUnzipFile As String, intFileNum As Integer
  
  strUnzipFile = Environ("windir") & IIf(InStr(Environ("OS"), "NT"), "\system32", "\system") & "\Unzip32.dll"

  On Error Resume Next
  If Dir$(strUnzipFile) = vbNullString Then
      If MsgBox("The unzip32.dll file does not appear to be installed on your system! It is a required file, to use the zip file searching capabilities of Code Hunter. Would you like Code Hunter to auto install this file for you?", vbCritical + vbYesNo) = vbYes Then
          intFileNum = FreeFile
          Open strUnzipFile For Binary As #intFileNum
            Put #intFileNum, , LoadResData(101, "CUSTOM")
          Close intFileNum
      Else
          MsgBox "Code Hunter may not function properly without the unzip32.dll installed. No zip files will be checked.", vbCritical
      End If
   End If
   
   frmMain.Show
      

End Sub
