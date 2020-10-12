Attribute VB_Name = "Module1"
Declare Function FindWindow Lib "user32" Alias _
"FindWindowA" (ByVal lpClassName As String, _
ByVal lpWindowName As String) As Long

Declare Function GetWindow Lib "user32" (ByVal hwnd _
As Long, ByVal wCmd As Long) As Long

Declare Function OpenIcon Lib "user32" (ByVal hwnd _
As Long) As Long

Declare Function SetForegroundWindow Lib "user32" _
(ByVal hwnd As Long) As Long
         
Public Const GW_HWNDPREV = 3


