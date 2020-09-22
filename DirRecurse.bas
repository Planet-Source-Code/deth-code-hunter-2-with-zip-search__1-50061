Attribute VB_Name = "DirRecurse"
Option Explicit
Option Compare Text

'Directory Recursion Bas Module - By Deth

Private Const Period As String = "."

Public Cancelled As Boolean

'retreives all the files in a folder, does not recurse into subfolders
'full file mask capability! you can use multiple mask list like so "*.exe;*.ocx;*.dll"
Sub GetFiles(Files As Collection, ByVal Path As String, Optional ByVal Mask As String = "*.*")

  Dim sFile As String, MaskArr() As String, X As Long

    If InStr(Mask, ";") Then
        MaskArr = Split(Mask, ";")
      Else
        ReDim MaskArr(0) As String
        MaskArr(0) = Mask
    End If

    Path = FormatPath(Path) & "\"
    On Error Resume Next
        For X = 0 To UBound(MaskArr)
            sFile = Dir$(Path & "\" & MaskArr(X))
            Do While LenB(sFile) > 0 And Not Cancelled
                If (GetAttr(FormatPath(Path & sFile)) And vbDirectory) <> vbDirectory Then
                    Files.Add Path & sFile
                End If
                sFile = Dir$
            Loop
        Next X

End Sub

'returns all folders inside a folder, no subfolder recursion
Sub GetFolders(Folders As Collection, ByVal Path As String)

  Dim sFolder As String

    Path = FormatPath(Path) & "\"
    sFolder = Dir$(Path, vbDirectory)
    Do While (Len(sFolder) <> 0) And Not Cancelled
        If Not (InStr(sFolder, String$(Len(sFolder), Period)) > 0) Then
            If (GetAttr(Path & sFolder) And vbDirectory) = vbDirectory Then
                Folders.Add Path & sFolder
            End If
        End If
        sFolder = Dir$
    Loop

End Sub


'gets all the folders including subfolders starting from path, and descending
Sub RecurseFolders(Folders As Collection, ByVal Path As String)

  Dim sFolder As String
  Dim colNew As Collection

    On Error Resume Next
        Path = FormatPath(Path)
        Folders.Add Path
        Path = Path & "\"
        Set colNew = New Collection
        DoEvents

        sFolder = Dir$(Path, vbDirectory)
        Do While (LenB(sFolder) > 0) And Not Cancelled
            If Not (InStr(sFolder, String$(Len(sFolder), Period)) > 0) Then
                If (GetAttr(Path & sFolder) And vbDirectory) = vbDirectory Then
                    colNew.Add Path & sFolder
                End If
            End If
            sFolder = Dir$
        Loop

        If colNew.Count > 0 Then
            Do While colNew.Count > 0 And Not Cancelled
                RecurseFolders Folders, colNew(1)
                colNew.Remove 1
            Loop
        End If

End Sub

'returns all files in current folder including subfolders
Sub RecurseFiles(Files As Collection, ByVal Path As String, Optional ByVal Mask As String = "*.*")

    On Error Resume Next
      Dim sFolder As String, colFolders As New Collection

        Path = FormatPath(Path)
        GetFiles Files, Path, Mask
        Path = Path & "\"
        DoEvents
        
        sFolder = Dir$(Path, vbDirectory)
        Do While (LenB(sFolder) > 0) And Not Cancelled
            If Not (InStr(sFolder, String$(Len(sFolder), Period)) > 0) Then
                If (GetAttr(Path & sFolder) And vbDirectory) = vbDirectory Then
                    colFolders.Add Path & sFolder
                End If
            End If
            sFolder = Dir$
        Loop

        If colFolders.Count > 0 Then
            Do While (colFolders.Count > 0) And Not Cancelled
                RecurseFiles Files, colFolders(1), Mask
                colFolders.Remove 1
            Loop
        End If

End Sub

'simple function that does some checking of a folderpath
'and removes trailing slashes
Function FormatPath(ByVal FolderPath As String) As String

    On Error Resume Next
        If LenB(FolderPath) > 4 Then
            Do Until Right$(FolderPath, 1) <> "\"
                FolderPath = Left$(FolderPath, Len(FolderPath) - 1)
            Loop
            FolderPath = Replace$(FolderPath, "/", "\")
        End If

        If LenB(FolderPath) > 4 Then
            FormatPath = Left$(FolderPath, 2) & Replace$(Mid$(FolderPath, 3), "\\", "\")
          Else
            FormatPath = FolderPath
        End If

End Function

'brute force folder delete, and all subfolders...
Sub NukeFolders(ByVal FolderPath As String)

  Dim colFolders As Collection, X As Long

  'gather all folders
  Set colFolders = New Collection
  RecurseFolders colFolders, FolderPath
   
  '2 loops, (this gives the os time to release read-only settings on folders)
  
  'first delete all files
  X = colFolders.Count
  Do While X > 0
      FileKill colFolders(X) & "\*.*"
      X = X - 1
  Loop
  
  'delete all folders,working backwards
  On Error Resume Next
  Do While colFolders.Count > 0
      RmDir colFolders(colFolders.Count)
      colFolders.Remove colFolders.Count
  Loop
  On Error GoTo 0
  
End Sub

Public Sub FileKill(ByVal FilePath As String)

    On Error Resume Next
        SetAttr FilePath, vbNormal
        Kill FilePath
    On Error GoTo 0
    
End Sub

