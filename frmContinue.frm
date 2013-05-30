VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmContinue 
   Caption         =   "継続"
   ClientHeight    =   720
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   1905
   OleObjectBlob   =   "frmContinue.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmContinue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function ShowWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const SW_SHOWMINIMIZED = 2
Private Const SW_SHOWNORMAL = 1
Private Const SW_MINIMIZE = 6
Private Const SW_RESTORE = 9
Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = (-20)
Private Const WS_SYSMENU = &H80000

Private mlHwnd As Long
Private Sub cmdContinue_Click()
    blFlgEnd = False
    setLayWin
    
End Sub

Private Sub UserForm_Activate()

    Dim dStyle As Long
    
    '自らのハンドルを取得
    mlHwnd = FindWindow("ThunderDFrame", "UserForm1")
    Do While mlHwnd = 0
        mlHwnd = FindWindow("ThunderDFrame", "継続")
        DoEvents
    Loop
    
    '閉じるボタンの消去
    dStyle = GetWindowLong(mlHwnd, GWL_STYLE)
    dStyle = SetWindowLong(mlHwnd, GWL_STYLE, dStyle Xor WS_SYSMENU)
    
    'Set topmost
    SetWindowPos mlHwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE

End Sub
