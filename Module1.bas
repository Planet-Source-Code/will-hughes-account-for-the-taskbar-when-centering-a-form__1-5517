Attribute VB_Name = "Module1"
'Coded By: Will Hughes
Private Declare Function GetSystemMetrics Lib "user32" _
    (ByVal nIndex As Long) As Long
Private Declare Function GetWindowLong Lib "user32" _
    Alias "GetWindowLongA" (ByVal hwnd As Long, _
    ByVal nIndex As Long) As Long
Private Const SM_CXFULLSCREEN = 16
Private Const SM_CYFULLSCREEN = 17
Public Sub CenterForm(Frm As Form)
'Coded By: Will Hughes
    Dim Left As Long, Top As Long
    Left = (Screen.TwipsPerPixelX _
        * (GetSystemMetrics(SM_CXFULLSCREEN) / 2)) - _
        (Frm.Width / 2)
    Top = (Screen.TwipsPerPixelY * _
        (GetSystemMetrics(SM_CYFULLSCREEN) / 2)) - _
        (Frm.Height / 2)
    Frm.Move Left, Top
End Sub

