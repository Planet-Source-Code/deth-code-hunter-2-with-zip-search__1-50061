Attribute VB_Name = "BrowseDialog"
'=====================================================================================
' Browse for a Folder using SHBrowseForFolder API function with a callback
' function BrowseCallbackProc. Can also include files

'Original Code By:
' Stephen Fonnesbeck
' steev@xmission.com
' http://www.xmission.com/~steev
' Feb 20, 2000

'Modified By:
' Lewis Miller (aka Deth)
' dethbomb@hotmail.com

Option Explicit

Public Enum CsIdlInfo
    CSIDL_DESKTOP = &H0
    CSIDL_PROGRAMS = &H2
    CSIDL_PRINTERS = &H4
    CSIDL_PERSONAL = &H5
    CSIDL_FAVORITES = &H6
    CSIDL_STARTUP = &H7
    CSIDL_RECENT = &H8
    CSIDL_SENDTO = &H9
    CSIDL_STARTMENU = &HB
    CSIDL_NETWORK = &H12
    CSIDL_NETHOOD = &H13
    CSIDL_FONTS = &H14
    CSIDL_TEMPLATES = &H15
End Enum

Private Type ITEMID
    cb As Long
    abID As Byte
End Type

Private Type ITEMIDLIST
    mkid As ITEMID
End Type


Private Enum BROWSE_OPTIONS
    BIF_RETURNONLYFSDIRS = &H1
    BIF_DONTGOBELOWDOMAIN = &H2
    BIF_STATUSTEXT = &H4&
    BIF_RETURNFSANCESTORS = &H8
    BIF_EDITBOX = &H10
    BIF_VALIDATE = &H20
    BIF_NEWDIALOGSTYLE = &H40
    BIF_BROWSEINCLUDEURLS = &H80
    BIF_BROWSEFORCOMPUTER = &H1000
    BIF_BROWSEFORPRINTER = &H2000
    BIF_BROWSEINCLUDEFILES = &H4000
    BIF_SHAREABLE = &H8000
End Enum

Private Const MAX_PATH = 260

Private Const WM_USER = &H400

Private Const BFFM_INITIALIZED As Long = 1
Private Const BFFM_SELCHANGED As Long = 2
Private Const BFFM_VALIDATEFAILEDA As Long = 3
Private Const BFFM_VALIDATEFAILEDW As Long = 4

Private Const BFFM_SETSTATUSTEXTA As Long = (WM_USER + 100)
Private Const BFFM_ENABLEOK As Long = (WM_USER + 101)
Private Const BFFM_SETSELECTIONA As Long = (WM_USER + 102)
Private Const BFFM_SETSELECTIONW As Long = (WM_USER + 103)
Private Const BFFM_SETSTATUSTEXTW As Long = (WM_USER + 104)

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
'Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)

Private Type BrowseInfo
    hwndOwner      As Long
    pIDLRoot       As Long
    pszDisplayName As String
    lpszTitle      As String
    ulFlags        As Long
    lpfnCallback   As Long
    lParam         As Long
    iImage         As Long
End Type

Private strCurrentDirectory As String

'get special folders (ie My Documents,Program Files)
Public Function SpecialFolder(CSIDL As CsIdlInfo) As String
    
    Dim FolderPath As String * MAX_PATH
    Dim lngReturn As Long, IDL As ITEMIDLIST
    
    On Error Resume Next
    lngReturn = SHGetSpecialFolderLocation(0&, CSIDL, IDL)
    If lngReturn = 0 Then
        lngReturn = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal FolderPath$)
        SpecialFolder = Left$(FolderPath, InStr(FolderPath, vbNullChar) - 1)
    End If

End Function

'shows a folder selector form
Public Function ShowBrowse(Owner As Form, ByVal DialogTitle As String, ByVal StartDir As String, Optional ByVal IncludeFiles As Boolean) As String

  Dim lpIDList As Long
  Dim strBuffer As String
  Dim tBrowseInfo As BrowseInfo

    On Error Resume Next
        
    If LenB(StartDir) = 0 Then
       If LenB(strCurrentDirectory) > 0 Then
           StartDir = Left$(strCurrentDirectory, Len(strCurrentDirectory) - 1)
       Else
           StartDir = CurDir$
       End If
    End If
        
    strCurrentDirectory = StartDir & vbNullChar

    With tBrowseInfo
        .hwndOwner = Owner.hWnd
        .lpszTitle = DialogTitle & vbNullChar
        If IncludeFiles Then
            .ulFlags = BIF_BROWSEINCLUDEFILES
        End If
        If InStr(1, Environ$("OS"), "NT", vbTextCompare) Then
            .ulFlags = .ulFlags Or BIF_NEWDIALOGSTYLE
        End If
        .ulFlags = .ulFlags Or BIF_DONTGOBELOWDOMAIN Or BIF_STATUSTEXT
        .lpfnCallback = FunctionAddress(AddressOf BrowseCallbackProc)
    End With

    lpIDList = SHBrowseForFolder(tBrowseInfo)

    If (lpIDList) Then
        strBuffer = Space$(MAX_PATH)
        If SHGetPathFromIDList(lpIDList, strBuffer) Then
            strBuffer = Left$(strBuffer, InStr(strBuffer, vbNullChar) - 1)
        End If
        Call CoTaskMemFree(lpIDList)
        ShowBrowse = strBuffer
    End If

End Function

'used by the showbrowse() function
Private Function BrowseCallbackProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lp As Long, ByVal pData As Long) As Long

  Dim lpIDList As Long
    Dim lngReturn As Long
    Dim strBuffer As String

    On Error Resume Next

    Select Case uMsg

        Case BFFM_INITIALIZED
            Call SendMessage(hWnd, BFFM_SETSELECTIONA, 1, ByVal strCurrentDirectory)

        Case BFFM_SELCHANGED
            strBuffer = Space$(MAX_PATH)
            lngReturn = SHGetPathFromIDList(lp, strBuffer)
            If lngReturn = 1 Then
                If strBuffer <> strCurrentDirectory Then
                    Call SendMessage(hWnd, BFFM_SETSTATUSTEXTA, 0, ByVal strBuffer)
                End If
            End If

    End Select

    BrowseCallbackProc = 0

End Function

' Assign a function pointer to a variable.
Private Function FunctionAddress(Address As Long) As Long

    FunctionAddress = Address

End Function

