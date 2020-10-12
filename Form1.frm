VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Hanya Satu Aplikasi yang Boleh Tampil (2)"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   6495
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ShowPrevInstance()
    Dim OldTitle As String
    Dim ll_WindowHandle As Long
    'Simpan judul ini ke dalam variabel OldTitle
    OldTitle = App.Title
    'Ganti judul aplikasinya...
    App.Title = "abcde - Aplikasi ini akan ditutup!"
    'Cari program sebelumnya. Jika Anda menggunakan VB
    '5.0, ganti "ThunderRT6Main" menjadi
    '"ThunderRT5Main"
    ll_WindowHandle = FindWindow("ThunderRT6Main", _
                      OldTitle)
    'Jika tidak ada aplikasi sebelumnya dibuka, keluar
    'langsung dari prosedur ini
    If ll_WindowHandle = 0 Then Exit Sub
    ll_WindowHandle = GetWindow(ll_WindowHandle, _
                      GW_HWNDPREV)
    'Sekarang ganti window tersebut...
    Call OpenIcon(ll_WindowHandle)
    'Dan bawa sebagai latar depan (tampil di depan)
    Call SetForegroundWindow(ll_WindowHandle)
    End
End Sub

Private Sub Form_Load()
    If App.PrevInstance Then ShowPrevInstance
End Sub


