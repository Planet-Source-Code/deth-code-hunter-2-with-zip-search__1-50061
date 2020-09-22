VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmMain 
   Caption         =   "Code Hunter v2"
   ClientHeight    =   5145
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9105
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5145
   ScaleWidth      =   9105
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCopyCode 
      Caption         =   "Copy"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7875
      TabIndex        =   7
      Top             =   90
      Width           =   915
   End
   Begin VB.CommandButton cmdStartSearch 
      Caption         =   "Search"
      Height          =   375
      Left            =   5940
      TabIndex        =   6
      Top             =   90
      Width           =   960
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   5
      Top             =   4815
      Width           =   9105
      _ExtentX        =   16060
      _ExtentY        =   582
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtDisplayFunction 
      Height          =   1500
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   3060
      Width           =   8700
   End
   Begin VB.CommandButton cmdCancelSearch 
      Caption         =   "Cancel"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6930
      TabIndex        =   3
      Top             =   90
      Width           =   915
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   375
      Left            =   5445
      TabIndex        =   2
      Top             =   90
      Width           =   420
   End
   Begin VB.TextBox txtFolderPath 
      Height          =   330
      Left            =   90
      TabIndex        =   1
      Top             =   90
      Width           =   5280
   End
   Begin MSComctlLib.ListView lvResults 
      Height          =   2490
      Left            =   45
      TabIndex        =   0
      Top             =   495
      Width           =   8700
      _ExtentX        =   15346
      _ExtentY        =   4392
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "File"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "In Folder"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "In Function"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "At Line"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Program: Code Hunter
'Author: Lewis Miller
'Email: dethbomb@hotmail.com
'Date: 11/22/03

'api function to keep track of time
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
'api to open files
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'these are public so we can change them from a different form
Public SearchMethod   As VbCompareMethod
Public SearchString   As String
Public FileMask       As String
Public SearchZips     As Boolean

'an array to store items to search for
Dim Search()          As String
Dim SearchCount       As Long

'open a folder form, to select a folder
Private Sub cmdBrowse_Click()

  Dim strFolderPath As String

    'use txtFolderPath as our start path if we can
    If LenB(txtFolderPath.Text) > 2 Then
        'turn on error handling
        On Error Resume Next
        'check to see if the folder exists
        If Dir$(FormatPath(txtFolderPath.Text) & "\nul", vbDirectory) <> vbNullString Then
           strFolderPath = FormatPath(txtFolderPath.Text)
        End If
        'turn off error handling
        On Error GoTo 0
    End If

    'show the browse window
    strFolderPath = ShowBrowse(Me, "Select Code Folder", strFolderPath, False)

    'check to see if a folder was selected and fill in folder path textbox
    If Len(strFolderPath) > 0 Then
        txtFolderPath.Text = strFolderPath
        ChDir strFolderPath
    End If

End Sub

'*********************************************************************************
'Note: I broke this function up into two parts, because it was getting too big.
'      Now this section loads the files, while the ProcessFile() function actually
'      searches the file.

'the main search 'engine'
'gathers together all files and then looks through them
'each for a matching item for searchstring
'*********************************************************************************
Public Sub StartSearch()

  Dim colFiles             As Collection
  Dim colZipFiles          As Collection
  Dim strZipFolder         As String
  Dim varCurrentFile       As Variant
  Dim blnProcessingZip     As Boolean
  Dim lngTotalTime         As Long
  Dim lngTotalLines        As Long
  Dim lngTotalFiles        As Long
  Dim lngTotalBytes        As Double
  Dim X                    As Long
  
    'check to make sure search string is valid
    If LenB(SearchString) = 0 Then
        MsgBox "Invalid Search String!", vbCritical
        GoTo Quit
    End If

    'make sure we have a file mask
    If LenB(FileMask) = 0 Then
        FileMask = "*.bas;*.frm;*.ctl"
    End If

    'see if we will look through zip files also, and make sure
    'the zip file mask is in the file mask
    If SearchZips Then
        If MsgBox("You have chosen to include zip files in your search, all zip files will be unzipped to a folder with the same name as the zip file. Are you sure You want to leave this option enabled? (note: all folders created are auto deleted after scanned for files)", vbCritical + vbYesNo) = vbYes Then
            If InStr(1, FileMask, ".zip", vbTextCompare) = 0 Then
                FileMask = FileMask & ";*.zip"
            End If
        Else
            SearchZips = False
        End If
    End If

    'clear list items if any
    lvResults.ListItems.Clear
    'clear function display text
    txtDisplayFunction.Text = ""

    'initialize and seperate search items into an array
    'to be used by the ContainsSearchItem() function
    If InStr(SearchString, vbNullChar) Then
        Search = Split(SearchString, vbNullChar)
        SearchCount = UBound(Search) + 1
      Else
        ReDim Search(0) As String
        Search(0) = SearchString
        SearchCount = 1
    End If

    'initialize variables that need it
    lngTotalTime = timeGetTime
    Set colFiles = New Collection

    'load all the files in the folder
    Status "Loading File List. Please Wait..."
    DoEvents
    Call RecurseFiles(colFiles, txtFolderPath, FileMask)

    'now loop through each file and search for matches to search items
    If colFiles.Count > 0 Then
        For Each varCurrentFile In colFiles
        
           On Error Resume Next
           lngTotalBytes = lngTotalBytes + FileLen(CStr(varCurrentFile))
           On Error GoTo 0
           
           'see if its a zip file
             If StrComp(Right$(varCurrentFile, 4), ".ZIP", 1) = 0 Then
                  Set colZipFiles = New Collection
                 'we have to keep track of unzip folders, so we can delete them
                  strZipFolder = UnPackZipFiles(colZipFiles, varCurrentFile, FileMask)
                  If colZipFiles.Count > 0 Then
                      X = 1
                      'loop through unzipped files and see if theres zips within
                      'zips, if not process it
CheckAgain:
                      Do While X < colZipFiles.Count + 1
                          On Error Resume Next
                          lngTotalBytes = lngTotalBytes + FileLen(colZipFiles(X))
                          On Error GoTo 0

                          If StrComp(Right$(colZipFiles(X), 4), ".ZIP", 1) = 0 Then
                              Call UnPackZipFiles(colZipFiles, colZipFiles(X), FileMask)
                              colZipFiles.Remove X
                              GoTo CheckAgain
                           Else
                              lngTotalFiles = lngTotalFiles + 1
                               'unfreeze our app
                              If lngTotalFiles Mod 5 = 0 Then
                                  DoEvents
                              End If
                              
                              lngTotalLines = lngTotalLines + ProcessFile(colZipFiles(X))
                           End If
                           X = X + 1
                       Loop
                       
                       'delete the folder (hehe, found this function with the old version of code hunt)
                       Status "Nuking " & strZipFolder
                       NukeFolders strZipFolder
                  End If
              Else
                  lngTotalFiles = lngTotalFiles + 1
                  'unfreeze our app
                  If lngTotalFiles Mod 5 = 0 Then
                      DoEvents
                  End If
                  
                  lngTotalLines = lngTotalLines + ProcessFile(varCurrentFile)
              End If
              If Cancelled Then GoTo Quit
         Next varCurrentFile
    End If
    
Quit:
    'do a second folder delete for leftover zip folders that didnt get deleted
    'note: we are re-using the colfiles collection for folders so we dont have
    'to dim another variable
    If SearchZips = True Then
      Set colFiles = New Collection
      RecurseFolders colFiles, txtFolderPath
      X = colFiles.Count
      If X > 0 Then
          On Error Resume Next
          Do While X > 0
              If InStr(colFiles(X), " [ZIP]") Then
                 Status "Nuking " & colFiles(X)
                 If (X Mod 10) = 0 Then DoEvents
                 FileKill colFiles(X) & "\*.*"
                 RmDir colFiles(X)
              End If
              X = X - 1
          Loop
          On Error GoTo 0
      End If
    End If
    
    'all done! :)
    cmdStartSearch.Enabled = True
    cmdCancelSearch.Enabled = False
    cmdCopyCode.Enabled = (lvResults.ListItems.Count > 0)
    Cancelled = False
    Status "Search Complete"

    'show some stats
    txtDisplayFunction.Text = "Files Searched" & vbTab & "= " & lngTotalFiles & vbCrLf & _
                              "Total Results" & vbTab & "= " & CStr(lvResults.ListItems.Count) & vbCrLf & _
                              "Total Bytes" & vbTab & "= " & FormatBytes(lngTotalBytes) & vbCrLf & _
                              "Total Lines" & vbTab & "= " & CStr(lngTotalLines) & vbCrLf & _
                              "Total Time" & vbTab & "= " & CStr(CalculateTime(timeGetTime - lngTotalTime)) & vbCrLf

End Sub

'this function actually processes the code file, and
'returns the amount of lines processed
Function ProcessFile(ByVal strCurrentFile As String) As Long

  Dim lvListItem           As MSComctlLib.ListItem
  Dim intFileNum           As Integer
  Dim lngCurrentLine       As Long
  Dim strNextLine          As String
  Dim strCurrentFunction   As String
  Dim strCurrentCode       As String
  Dim blnFoundItem         As Boolean
  Dim blnInsideFunction    As Boolean
  Dim blnFoundAttribute    As Boolean

    'show filepath in status bar
    Status "Searching " & CStr(strCurrentFile)

    'get a free file number
    intFileNum = FreeFile
    'open the file
    Open strCurrentFile For Input As intFileNum

    'all vb files have hidden attributes at the beginning of the file
    'so we loop until we find it, if not, its not a vb file...
    Do While Not EOF(intFileNum)
        Line Input #intFileNum, strNextLine    'grab line of code
        ProcessFile = ProcessFile + 1
        If StrComp(Left$(Trim$(strNextLine), 12), "ATTRIBUTE VB", 1) = 0 Then
            blnFoundAttribute = True
        Else
            If blnFoundAttribute Then
                blnFoundAttribute = False
                Exit Do
            End If
        End If
     Loop

    'now loop thru the actual code lines
    Do While Not EOF(intFileNum)

        Line Input #intFileNum, strNextLine    'grab line of code
        ProcessFile = ProcessFile + 1

        lngCurrentLine = lngCurrentLine + 1

        If LenB(strNextLine) > 0 Then          'check length
            strNextLine = TabTrim(strNextLine) 'trim excess
            If LenB(strNextLine) > 0 Then      'check length again

                'quick check to see if it could be a function
                If ContainsWord("Function ,Sub ,Property ", strNextLine) Then
                    'double check
                    If StartWord("Declare ,Private ,Public ,Friend ,Static ,Function ,Sub ,Property ", strNextLine) Then
                        'grab function name
                        strCurrentFunction = ParseFunctionName(strNextLine)
                        'store code
                        strCurrentCode = strNextLine & vbCrLf
                        'flip switch
                        blnInsideFunction = True

                        'look for search items in current line of code
                        If ContainsSearchItem(strNextLine) Then
                            blnFoundItem = True
                          Else
                            blnFoundItem = False
                        End If

                        'is it an API function?
                        If IsAPIFunction(strNextLine) Then
                            'check for line continuations
                            Do While RightCheck(strNextLine, "_") And Not EOF(intFileNum)
                                'grab line of code
                                Line Input #intFileNum, strNextLine
                                'increment line count
                                ProcessFile = ProcessFile + 1
                                'lngTotalLines = lngTotalLines + 1
                                'store it
                                strCurrentCode = strCurrentCode & (strNextLine & vbCrLf)
                                If Cancelled Then GoTo Quit
                            Loop
                            'look for search items
                            blnFoundItem = ContainsSearchItem(strCurrentCode)
                          Else
                            'not an api call so its a regular function, sub, or property
                            'grab all the code for this code block
                            Do While blnInsideFunction And (Not EOF(intFileNum))
                                'grab line of code
                                Line Input #intFileNum, strNextLine
                                'increment line counter
                                ProcessFile = ProcessFile + 1
                                'lngTotalLines = lngTotalLines + 1
                                'store it
                                strCurrentCode = strCurrentCode & (strNextLine & vbCrLf)
                                lngCurrentLine = lngCurrentLine + 1
                                'check to see that its not a blank line
                                If LenB(strNextLine) > 0 Then
                                    'look for end line
                                    If ContainsWord("Function,Sub,Property", strNextLine) Then
                                        If StartWord("End ", strNextLine) Then
                                            blnInsideFunction = False
                                        End If
                                    End If
                                End If
                                'check for search items in current line
                                If ContainsSearchItem(strNextLine) Then
                                    blnFoundItem = True
                                End If
                                If Cancelled Then GoTo Quit
                            Loop

                        End If

                        'did we find anything?
                        If blnFoundItem Then
                            'yes so add it to the list
                            blnFoundItem = False
                            Set lvListItem = lvResults.ListItems.Add(, , Mid$(CStr(strCurrentFile), InStrRev(CStr(strCurrentFile), "\") + 1))
                            With lvListItem
                                'store the code in the tag property
                                .Tag = Left$(strCurrentCode, Len(strCurrentCode) - 2)
                                'file name
                                .SubItems(1) = Left$(CStr(strCurrentFile), InStrRev(CStr(strCurrentFile), "\") - 1)
                                'function name
                                .SubItems(2) = strCurrentFunction
                                'code line number
                                .SubItems(3) = CStr(lngCurrentLine)
                                .EnsureVisible
                            End With
                        End If
                        
                    End If
                End If
            End If
        End If
        If Cancelled Then GoTo Quit
    Loop

Quit:
    Close intFileNum

End Function

'this function unpacks a zip file into a folder, then loads the files
'into a collection
Function UnPackZipFiles(colFiles As Collection, ByVal strZipPath As String, ByVal strFileMask As String) As String

  Dim lngReturn As Long, lngPathLen As Long
    
    lngPathLen = Len(strZipPath)
    If lngPathLen > 0 Then
        lngReturn = InStrRev(strZipPath, "\")
        If lngReturn < lngPathLen - 2 Then
            lngReturn = InStr(lngReturn + 1, strZipPath, ".")
            If lngReturn > 0 Then
                
                'get a new folder path from file path, removing '.zip'
                UnPackZipFiles = Left$(strZipPath, lngReturn - 1) & " [ZIP]"
                'unzip the files
                On Error Resume Next
                  Status "Unzipping " & strZipPath
                  lngReturn = VBUnzip(strZipPath, UnPackZipFiles, 0, 0, 0, 2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0)
                    'see if its been unzipped
                    If Dir$(UnPackZipFiles & "\nul") <> vbNullString Then
                        On Error GoTo 0
                        If InStr(1, strFileMask, "zip", vbTextCompare) = 0 Then
                            strFileMask = strFileMask & ";*.zip"
                        End If
                        'load all the files that have been unzipped
                        RecurseFiles colFiles, UnPackZipFiles, strFileMask
                    End If
             End If
        End If
    End If

End Function


Sub Status(ByVal strStatusText As String)

    StatusBar1.SimpleText = strStatusText

End Sub

Function ContainsSearchItem(ByVal strLine As String) As Boolean

  Dim X As Long

    For X = 0 To SearchCount - 1
        If InStr(1, strLine, Search(X), SearchMethod) Then
            ContainsSearchItem = True
            Exit Function
        End If
    Next X

End Function


Private Sub cmdStartSearch_Click()
     
    SearchString = ""
    frmFind.Show , Me

End Sub


Private Sub cmdCancelSearch_Click()

    Cancelled = True

End Sub

Private Sub cmdCopyCode_Click()

    If Not (lvResults.SelectedItem Is Nothing) Then
        Clipboard.Clear
        Clipboard.SetText lvResults.SelectedItem.Tag
    End If

End Sub

Private Sub Form_Load()

    txtFolderPath = GetSetting(App.Title, "Settings", "LastFolder", "C:\")

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    SaveSetting App.Title, "Settings", "LastFolder", txtFolderPath

    On Error Resume Next
        Unload frmFind

End Sub

Private Sub Form_Resize()

    On Error Resume Next

        If WindowState <> 1 Then
            With txtDisplayFunction
                .Height = Me.Height \ 4
                lvResults.Width = ScaleWidth - 100
                lvResults.Height = ScaleHeight - (550 + StatusBar1.Height + .Height)
                .Top = lvResults.Top + lvResults.Height + 50
                .Width = ScaleWidth - 100
                .Left = lvResults.Left
            End With
        End If

End Sub

Private Sub lvResults_DblClick()

   Dim strFilePath As String

  'open file if double clicked

    If Not (lvResults.SelectedItem Is Nothing) Then
        With lvResults.SelectedItem
            'if it was in a zip file, show zip file, else show normal file
            If InStr(1, .SubItems(1), " [ZIP]", vbTextCompare) Then
                 strFilePath = Mid$(.SubItems(1), 1, InStr(1, .SubItems(1), " [ZIP]", 1) - 1) & ".zip"
            Else
                 strFilePath = .SubItems(1) & "\" & .Text
            End If
            Call ShellExecute(Me.hWnd, "open", strFilePath, vbNullString, "", 1)
        End With
    End If

End Sub

Private Sub lvResults_ItemClick(ByVal Item As MSComctlLib.ListItem)

    txtDisplayFunction.Text = Item.Tag
    Status Search(0) & " found in " & Item.Text & " in function " & Item.SubItems(2) & " @ line #" & Item.SubItems(3)

End Sub

