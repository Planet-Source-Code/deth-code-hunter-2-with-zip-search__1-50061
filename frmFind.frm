VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3330
   Icon            =   "frmFind.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   3330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Remove"
      Height          =   375
      Left            =   2160
      TabIndex        =   10
      Top             =   2520
      Width           =   1050
   End
   Begin VB.ListBox lstSearchItems 
      Height          =   1620
      Left            =   90
      TabIndex        =   9
      Top             =   405
      Width           =   3120
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   375
      Left            =   1080
      TabIndex        =   8
      Top             =   2520
      Width           =   1005
   End
   Begin VB.CheckBox chkCheckZips 
      Caption         =   "Search Zip Files"
      Height          =   195
      Left            =   1620
      TabIndex        =   7
      Top             =   3690
      Width           =   1500
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2025
      TabIndex        =   6
      Top             =   4140
      Width           =   1140
   End
   Begin VB.CheckBox chkCaseSensitive 
      Caption         =   "Case Sensitive"
      Height          =   195
      Left            =   180
      TabIndex        =   4
      Top             =   3690
      Width           =   1365
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Start"
      Height          =   375
      Left            =   855
      TabIndex        =   3
      Top             =   4140
      Width           =   1095
   End
   Begin VB.TextBox txtFilemask 
      Height          =   330
      Left            =   90
      TabIndex        =   2
      Text            =   "*.bas;*.frm;*.ctl"
      Top             =   3195
      Width           =   3075
   End
   Begin VB.TextBox txtSearchString 
      Height          =   375
      Left            =   90
      TabIndex        =   0
      Top             =   2070
      Width           =   3120
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter search items below to the list"
      Height          =   240
      Index           =   2
      Left            =   135
      TabIndex        =   5
      Top             =   90
      Width           =   3030
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "File Mask:"
      Height          =   240
      Index           =   0
      Left            =   135
      TabIndex        =   1
      Top             =   2970
      Width           =   825
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSearch_Click()
  
  Dim X As Long
  
    If lstSearchItems.ListCount > 0 Then
        Me.Hide   'hide this form
        DoEvents
        With frmMain
            .SearchZips = CBool(chkCheckZips.Value)
            For X = 0 To lstSearchItems.ListCount - 1
              .SearchString = .SearchString & lstSearchItems.List(X) & vbNullChar  'located in frmMain declarations
            Next X
            .SearchString = Left$(.SearchString, Len(.SearchString) - 1)
            .SearchMethod = IIf(chkCaseSensitive.Value = 1, 0, 1)  'located in frmMain declarations
            .FileMask = txtFilemask.Text                           'located in frmMain declarations
            .cmdStartSearch.Enabled = False
            .cmdCancelSearch.Enabled = True
            .StartSearch                                           'public Sub (method) In frmMain
        End With
    Else
        MsgBox "Invalid Amount of Search Items", vbCritical
    End If

End Sub

Private Sub cmdCancel_Click()
    
    Unload Me
    
End Sub

Private Sub Command1_Click()
  
  With txtSearchString
      If .Text <> "" Then
          lstSearchItems.AddItem .Text
          .Text = ""
          SaveSearchList
      Else
          MsgBox "Please enter an item to search for first!", vbCritical
          .SetFocus
      End If
  End With
  
End Sub

Private Sub Command2_Click()
  
  With lstSearchItems
      If .ListCount > 0 Then
          If .ListIndex <> -1 Then
              If MsgBox("Are you sure you want to remove " & .List(.ListIndex) & " ?", vbQuestion + vbYesNo) = vbYes Then
                  .RemoveItem .ListIndex
              End If
          Else
              MsgBox "Nothing selected.", vbCritical
          End If
      End If
  End With
  
End Sub

Private Sub Form_Load()

    LoadSearchList
    txtFilemask.Text = GetSetting(App.Title, "Settings", "LastMask", txtFilemask.Text)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    SaveSearchList
    SaveSetting App.Title, "Settings", "LastMask", txtFilemask.Text
    
End Sub

Sub SaveSearchList()

 Dim X As Long
   With lstSearchItems
       SaveSetting App.Title, "Settings", "SearchCount", CStr(.ListCount)
       If .ListCount > 0 Then
            For X = 0 To .ListCount - 1
                SaveSetting App.Title, "Settings", "SearchItem" & CStr(X), .List(X)
            Next X
       End If
   End With
   
End Sub

Sub LoadSearchList()

 Dim X As Long, SearchAmount As Long
  
  lstSearchItems.Clear
  SearchAmount = CLng(GetSetting(App.Title, "Settings", "SearchCount", "0"))
  If SearchAmount > 0 Then
      For X = 0 To SearchAmount - 1
          lstSearchItems.AddItem GetSetting(App.Title, "Settings", "SearchItem" & CStr(X), "")
      Next X
  End If
  
End Sub
