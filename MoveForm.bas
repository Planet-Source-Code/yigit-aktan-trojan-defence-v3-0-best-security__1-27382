Attribute VB_Name = "MoveForm"
'Sets the API Calls Used In This App
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Declare Sub ReleaseCapture Lib "user32" ()
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
'Sets the FormDrag Settings
Public Sub FormDrag(TheForm As Form)
    ReleaseCapture
    Call SendMessage(TheForm.hwnd, &HA1, 2, 0&)
End Sub


