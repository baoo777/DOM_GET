VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmWindow 
   Caption         =   "IE選択"
   ClientHeight    =   1845
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4425
   OleObjectBlob   =   "frmWindow.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
    
    Dim objShell As Object
    Dim objWindow As Object
    
    Set objShell = CreateObject("Shell.Application")
    For Each objWindow In objShell.Windows
        If TypeName(objWindow.Document) = "HTMLDocument" Then
            If cmbWindow.Text = objWindow.Document.Title Then
                Set objIE = objWindow
                objIE.Visible = True
            End If
        End If
    Next
    If objIE Is Nothing Then
        MsgBox "HTMLDocumentが見つかりません", vbInformation
    End If
    Unload Me
    
End Sub

Private Sub UserForm_Initialize()
    
    
    Dim objShell As Object
    Dim objWindow As Object
    
    Set objShell = CreateObject("Shell.Application")
    For Each objWindow In objShell.Windows
        If TypeName(objWindow.Document) = "HTMLDocument" Then
            cmbWindow.AddItem objWindow.Document.Title
        End If
    Next
    
    
End Sub

Private Sub UserForm_Terminate()
    If Application.Visible = False Then
        Application.Visible = True
    End If
End Sub
