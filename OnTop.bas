Attribute VB_Name = "Module1"
Declare Sub SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal _
    hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx _
    As Long, ByVal cy As Long, ByVal wFlags As Long)
    Global Const HWND_TOPMOST = -1
    Global Const HWND_NOTOPMOST = -2
    Global AOTValue As Boolean



