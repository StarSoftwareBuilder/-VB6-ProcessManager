Attribute VB_Name = "basViewHid"

Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Boolean
Public Function EnumWindowsProc(ByVal hwnd As Long, ByVal lParam As Long) As Boolean
    Dim ID As Long
    Dim CoChua As Boolean
    Dim i As Integer
    CoChua = False
    GetWindowThreadProcessId hwnd, ID
    
    With frmMain.lstPro
        For i = 0 To .ListCount - 1
            If ID = Val(.List(i)) Then CoChua = True
        Next
        If CoChua = False Then .AddItem ID
    End With
    
If CoChua = False Then
    If CheckID(ID) <> ID Then
        Dim tmp As String
        tmp = ProcessPathByPID(ID)
        frmMain.SoLuong = frmMain.SoLuong + 1
        CoChua = False
        With frmMain

            Set lsv = .LV1.ListItems.Add()
            lsv.Text = GetFileName(tmp)
            lsv.SubItems(1) = tmp
            lsv.SubItems(2) = ID
            lsv.ForeColor = vbRed
'            lsv.Font.Bold = True
        End With
        
    End If
End If

    EnumWindowsProc = True
End Function
Public Function GetFileName(ByVal sPath As String) As String
GetFileName = Mid(sPath, InStrRev(sPath, "\") + 1)
End Function

