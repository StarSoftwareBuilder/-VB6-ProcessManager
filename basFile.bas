Attribute VB_Name = "basFile"
Option Explicit

'Get icon


Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbSizeFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl&, ByVal i&, ByVal hdcDest&, ByVal x&, ByVal y&, ByVal flags&) As Long
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Private Const SEE_MASK_INVOKEIDLIST = &HC
Private Const SEE_MASK_NOCLOSEPROCESS = &H40
Private Const SEE_MASK_FLAG_NO_UI = &H400
Public Declare Function ShellExecuteEx Lib "shell32" Alias "ShellExecuteExA" (SEI As SHELLEXECUTEINFO) As Long

Private Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * 260
    szTypeName As String * 80
End Type
Public Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hWnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type
Private Const SHGFI_DISPLAYNAME = &H200
Private Const SHGFI_EXETYPE = &H2000
Private Const SHGFI_SYSICONINDEX = &H4000
Private Const SHGFI_LARGEICON = &H0
Private Const SHGFI_SMALLICON = &H1
Private Const ILD_TRANSPARENT = &H1
Private Const SHGFI_SHELLICONSIZE = &H4
Private Const SHGFI_TYPENAME = &H400
Private Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME _
Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX _
Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE

Private SIconInfo As SHFILEINFO


'Dimensionalize SIconInfo as SHFILEINFO type structure
Public Sub GetIcon(icPath$, pDisp As PictureBox)
pDisp.Cls
Dim hImgSmall&: hImgSmall = SHGetFileInfo(icPath$, 0&, SIconInfo, Len(SIconInfo), BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
'call SHGetFileInfo to return a handle to the icon associated with the specified file
 ImageList_Draw hImgSmall, SIconInfo.iIcon, pDisp.hDC, 0, 0, ILD_TRANSPARENT
 'Draw the icon to the specified picturebox control
End Sub
Public Sub GetLargeIcon(icPath$, pDisp As PictureBox)
pDisp.Cls
Dim hImgLrg&: hImgLrg = SHGetFileInfo(icPath$, 0&, SIconInfo, Len(SIconInfo), BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)
'call SHGetFileInfo to return a handle to the icon associated with the specified file
 ImageList_Draw hImgLrg, SIconInfo.iIcon, pDisp.hDC, 0, 0, ILD_TRANSPARENT
 'Draw the icon to the specified picturebox control
End Sub
Public Sub ShowProperties(sFileName As String, hwndOwner As Long)
    '##############################################################################################
    'Displays the Properties of the specified file
    '##############################################################################################
    
    Dim SEI As SHELLEXECUTEINFO
    
    With SEI
        .cbSize = Len(SEI)
        .fMask = SEE_MASK_NOCLOSEPROCESS Or SEE_MASK_INVOKEIDLIST Or SEE_MASK_FLAG_NO_UI
        .hWnd = hwndOwner
        .lpVerb = "properties"
        .lpFile = sFileName
        .lpParameters = vbNullChar
        .lpDirectory = vbNullChar
        .nShow = 0
        .hInstApp = 0
        .lpIDList = 0
    End With
    
    Call ShellExecuteEx(SEI)
End Sub

Public Function zDeletefile(sFile) As Boolean
SetAttr sFile, vbNormal
DeleteFile sFile
Dim x As String
x = sFile
zDeletefile = Not FileExists(x)
End Function
Public Function FileExists(sFile As String) As Boolean
On Error Resume Next
FileExists = ((GetAttr(sFile) And vbDirectory) = 0)
End Function

Sub Main()
Shell "regsvr32 /s UniControls_v2.0.ocx"
frmMain.Show
End Sub

