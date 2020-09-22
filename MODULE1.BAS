Attribute VB_Name = "Module1"
Declare Function FindWindow Lib "User32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function GetClassName& Lib "User32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Declare Function GetParent Lib "User32" (ByVal hwnd As Long) As Long
Declare Function GetWindow Lib "User32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function GetWindowText Lib "User32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetCursorPos Lib "User32" (lpPoint As PointApi) As Long
Declare Function SendMessageByNum& Lib "User32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Declare Function SendMessageByString Lib "User32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Declare Function WindowFromPoint Lib "User32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Public Type PointApi
X As Long
Y As Long
End Type
Function FindChildByClass(parentw, childhand)
    Dim Firs As Integer, Firss As Integer, Room As String
    Firs = GetWindow(parentw, 5)
    If UCase(Mid(GC(Firs), 1, Len(childhand))) Like UCase(childhand) Then GoTo Digital
        Firs = GetWindow(parentw, 5)
    If UCase(Mid(GC(Firs), 1, Len(childhand))) Like UCase(childhand) Then GoTo Digital
    While Firs
        Firs = GetWindow(parentw, 5)
        If UCase(Mid(GC(Firss), 1, Len(childhand))) Like UCase(childhand) Then GoTo Digital
        Firs = GetWindow(Firs, 2)
        If UCase(Mid(GC(Firs), 1, Len(childhand))) Like UCase(childhand) Then GoTo Digital
        Wend
        FindChildByClass = 0
Digital:
Room$ = Firs
FindChildByClass = Room$
End Function

Function GC(Child)
    Dim B As String, G As Integer
    B = Space(250)
    G = GetClassName(Child, B, 250)
    GC = B
End Function

Function Tex(Child)
    Dim GT As Integer, TrimSpace As String, GetString As String
    GT = SendMessageByNum(Child, 14, 0&, 0&)
    TrimSpace$ = Space$(GT)
    GetString$ = SendMessageByString(Child, 13, GT + 1, TrimSpace$)
    Tex = TrimSpace$
End Function

Public Sub GetExplorerText()
Dim a As PointApi
ie = FindWindow("ieframe", vbNullString)
If ie <> 0 Then
Call GetCursorPos(a)
z = WindowFromPoint(a.X, a.Y)
aa = GetParent(z)
B = GetParent(aa)
Clipboard.Clear
    If B = ie Or z = ie Or aa = ie Then
        ie = FindWindow("ieframe", vbNullString)
        worker = FindChildByClass(ie, "workera")
        rebar = FindChildByClass(worker, "rebarwindow32")
        combo32 = FindChildByClass(rebar, "comboboxex32")
        newadd = Left(Tex(combo32), 150)
            If newadd <> Form1.Text2.Text Then
            SendKeys "^{a}", True
            SendKeys "^{c}", True
            Form1.Text2.Text = newadd
            Form1.Text1.Text = Clipboard.GetText
            End If
    End If
End If
End Sub
