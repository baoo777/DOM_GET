VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSource 
   Caption         =   "Web情報"
   ClientHeight    =   11265
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14565
   OleObjectBlob   =   "frmSource.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "frmSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private strURL As String
Private Declare Function ShowWindow Lib "user32.dll" (ByVal hwnd As _
Long, ByVal nCmdShow As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Const SW_SHOWMINIMIZED = 2
Private Const SW_SHOWNORMAL = 1
Private Const SW_MINIMIZE = 6
Private Const SW_RESTORE = 9
Private lngCurX As Long
Private lngCurLine As Long
Private Declare Function SendMessageW& Lib "user32" _
    (ByVal hwnd&, _
     ByVal uMsg&, _
     ByVal wParam&, _
     ByVal lParam&)
Private Type RECT
    Left As Long
    Top As Long
    Width As Long
    Height As Long
End Type
'Private Const LVM_SUBITEMHITTEST = &H1039   'Wide系
'Private Type LVHITTESTINFO
'    ptX As Long
'    ptY As Long
'    flags As Long
'    iItem As Long
'    iSubItem As Long
'    iGroup As Long
'End Type
'Private ht As LVHITTESTINFO
'Private Type RECT
'    Left As Long
'    Top As Long
'    Right As Long
'    Bottom As Long
'End Type
Private Type POINTAPI
  X As Long
  Y As Long
End Type
'Private Const LVM_FIRST = &H1000
'Private Const LVM_GETSUBITEMRECT = LVM_FIRST + 56
'Private Const LVIR_LABEL = 2
Private cbMenu As CommandBar

Private Sub chkInput_Click()
    If chkInput.Value = True Then
        lstObject.List(lstObject.ListIndex, 3) = "True"
        strctNode(lstObject.ListIndex).Operation.Value = "True"
    Else
        lstObject.List(lstObject.ListIndex, 3) = "False"
        strctNode(lstObject.ListIndex).Operation.Value = "False"
    End If
End Sub
Private Sub cmbNodeSelect_Click()
    strctNode(lstObject.ListIndex).Operation.Object = cmbNodeSelect.Text
    lstObject.List(lstObject.ListIndex, 1) = cmbNodeSelect.Text
'    If cmbNodeSelect.Text = "Form" Or cmbNodeSelect.Text = "FormNumber" Then
'        cmbOperate.Clear
'        'cmbOperate.AddItem "Submit"
'        AddOperation "Submit"
'    End If
    SetOperate strctNode(lstObject.ListIndex)
End Sub

Private Sub cmbOperate_Click()
    Dim i As Long
    If cmbOperate.ListIndex < 0 Then
        Exit Sub
    End If
    strctNode(lstObject.ListIndex).Operation.Operation = cmbOperate.List(cmbOperate.ListIndex)
    lstObject.List(lstObject.ListIndex, 2) = cmbOperate.List(cmbOperate.ListIndex)
    
    Select Case cmbOperate.Text
    Case "Click"
        txtInput.Visible = False
        cmbSelect.Visible = False
        chkInput.Visible = False
        cmdDownload.Visible = False
        lblPath.Visible = False
    Case "Input"
        txtInput.Visible = True
        cmbSelect.Visible = False
        chkInput.Visible = False
        cmdDownload.Visible = False
        lblPath.Visible = False
    Case "Select"
        txtInput.Visible = False
        cmbSelect.Visible = True
        cmbSelect.Clear
        For i = 0 To UBound(strctNode(lstObject.ListIndex).SelectList)
            cmbSelect.AddItem strctNode(lstObject.ListIndex).SelectList(i).Caption
        Next i
        chkInput.Visible = False
        cmdDownload.Visible = False
        lblPath.Visible = False
    Case "GetText"
        txtInput.Visible = False
        cmbSelect.Visible = False
        chkInput.Visible = False
        cmdDownload.Visible = False
        lblPath.Visible = False
    Case "Download"
        txtInput.Visible = False
        cmbSelect.Visible = False
        chkInput.Visible = False
        cmdDownload.Visible = True
        lblPath.Visible = True
    Case "Checked"
        txtInput.Visible = False
        cmbSelect.Visible = False
        For i = 0 To lstAttribute.ListCount - 1
            If lstAttribute.List(i, 0) = "type" Then
                If lstAttribute.List(i, 1) = "checkbox" Then
                    chkInput.Visible = True
                    optInput.Visible = False
                Else
                    chkInput.Visible = False
                    optInput.Visible = True
                End If
            End If
        Next i
        cmdDownload.Visible = False
        lblPath.Visible = False
    End Select
End Sub

Private Sub cmbSelect_Change()

End Sub

Private Sub cmbSelect_Click()
    If lstObject.List(lstObject.ListIndex, 2) = "Select" Then
        lstObject.List(lstObject.ListIndex, 3) = strctNode(lstObject.ListIndex).SelectList(cmbSelect.ListIndex).Value
        strctNode(lstObject.ListIndex).Operation.Value = strctNode(lstObject.ListIndex).SelectList(cmbSelect.ListIndex).Value
    End If
End Sub
' /////////////////////////////////////////////////////////////////////////////////////
' //#名称 cmdAdd_Click
' //
' //#概要 ソースコードを生成しtxtSourceに設定
' //
' //#引数 なし
' //
' //#戻値 なし
' //
' //#解説
' //
' //#履歴 2013/05/30
' //      Coded by YASUTADA OOBA
' //
' /////////////////////////////////////////////////////////////////////////////////////
Private Sub cmdAdd_Click()
    
    Dim strSrc1st As String
    Dim strSrc2nd As String
    Dim strSrc3rd As String
    Dim lngFirstNode As Long
    Dim i As Long
    
    With Application.VBE.ActiveVBProject.VBComponents("modDomAccess").CodeModule
        
        lngFirstNode = lstObject.ListCount
        For i = 0 To lstObject.ListCount - 1
            If lstObject.Selected(i) = True Then
                If lngFirstNode > i Then
                    lngFirstNode = i
                End If
                If strSrc = "" Then
                    Select Case lstObject.List(i, 1)
                    Case "Id"
                        strSrc = "    Set objDOM = objIE.Document.GetElementById(""" & strctNode(i).NodeId & """)"
                    Case "Name"
                        strSrc = "    Set objDOM = objIE.Document.GetElementsByName(""" & strctNode(i).NodeName & """).Item(" & strctNode(i).NodeElementByName & ")"
                    Case "TagName"
                        strSrc = "    Set objDOM = objIE.Document.GetElementsByTagname(""" & strctNode(i).NodeType & """).Item(" & strctNode(i).NodeElementByTag & ")"
                    Case "Form"
                        strSrc = "    Set objDOM = objIE.Document.Forms(""" & strctNode(i).FormNodeName & """)"
                    Case "FormNumber"
                        strSrc = "    Set objDOM = DOMGetDocObjectFromNumber(objIE,""" & strctNode(i).FormNodeNumber & """)"
                    Case "NodeNumber"
                        strSrc = "    Set objDOM = DOMGetDocObjectFromNumber(objIE,""" & strctNode(i).NodeNumber & """)"
                    End Select
                Else
                    Select Case lstObject.List(i, 1)
                    Case "Id"
                        strSrc = strSrc & vbCrLf & "    Set objDOM = objIE.Document.GetElementById(""" & strctNode(i).NodeId & """)"
                    Case "Name"
                        strSrc = strSrc & vbCrLf & "    Set objDOM = objIE.Document.GetElementsByName(""" & strctNode(i).NodeName & """).Item(" & strctNode(i).NodeElementByName & ")"
                    Case "TagName"
                        strSrc = strSrc & vbCrLf & "    Set objDOM = objIE.Document.GetElementsByTagname(""" & strctNode(i).NodeType & """).Item(" & strctNode(i).NodeElementByTag & ")"
                    Case "Form"
                        strSrc = strSrc & vbCrLf & "    Set objDOM = objIE.Document.Forms(""" & strctNode(i).FormNodeName & """)"
                    Case "FormNumber"
                        strSrc = strSrc & vbCrLf & "    Set objDOM = DOMGetDocObjectFromNumber(objIE,""" & strctNode(i).FormNodeNumber & """)"
                    Case "NodeNumber"
                        strSrc = strSrc & vbCrLf & "    Set objDOM = DOMGetDocObjectFromNumber(objIE,""" & strctNode(i).NodeNumber & """)"
                    End Select
                End If
                
                'Select Case lvObject.ListItems(i).SubItems(2)
                Select Case lstObject.List(i, 2)
                Case "Click"
                    strSrc = strSrc & vbCrLf & "    objDOM.Click"
                Case "Input"
                    strSrc = strSrc & vbCrLf & "    objDOM.Value = """ & lstObject.List(i, 3) & """"
                Case "Select"
                    strSrc = strSrc & vbCrLf & "    objDOM.Value = """ & lstObject.List(i, 3) & """"
                Case "GetText"
                    strSrc = strSrc & vbCrLf & "    Debug.Print objDOM.NodeValue"
                Case "Submit"
                    strSrc = strSrc & vbCrLf & "    objDOM.Submit"
                Case "Download"
                    If strctNode(i).href <> "" Then
                        strSrc = strSrc & vbCrLf & "    Ret = URLDownloadToFile(0,""" & strctNode(i).href & """, """ & lstObject.List(i, 3) & """, 0, 0 )"
                    ElseIf strctNode(i).src <> "" Then
                        strSrc = strSrc & vbCrLf & "    Ret = URLDownloadToFile(0,""" & strctNode(i).src & """, """ & lstObject.List(i, 3) & """, 0, 0 )"
                    End If
                    
                Case "Checked"
                    strSrc = strSrc & vbCrLf & "    objDOM.Checked = " & lstObject.List(i, 3)
                End Select
            End If
        Next i
        
        strSrc1st = "Private Declare Function URLDownloadToFile Lib ""urlmon"" Alias _" & vbCrLf & _
                    """URLDownloadToFileA"" (ByVal pCaller As Long, ByVal szURL As String, ByVal _" & vbCrLf & _
                    "szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long"

        strSrc1st = strSrc1st & vbCrLf & "Private Declare Sub Sleep Lib ""kernel32"" (ByVal dwMilliseconds As Long)"
        strSrc1st = strSrc1st & vbCrLf & "Sub Main()"
        strSrc1st = strSrc1st & vbCrLf & "    Dim objIE As Object"
        strSrc1st = strSrc1st & vbCrLf & "    Dim objDOM as Object"
        strSrc1st = strSrc1st & vbCrLf & "    Set objIE = DOMOpenURL(""" & strctNode(lngFirstNode).URL & """)"  ' strURL & """)"
        strSrc1st = strSrc1st & vbCrLf & "    objIE.Visible = True"

        
        
        
        
        strSrc3rd = "End Sub"
        strSrc3rd = strSrc3rd & vbCrLf & .Lines(.ProcStartLine("DOMOpenURL", vbext_pk_Proc), .ProcCountLines("DOMOpenURL", vbext_pk_Proc))
        strSrc3rd = strSrc3rd & vbCrLf & .Lines(.ProcStartLine("DOMGetDocObjectFromNumber", vbext_pk_Proc), .ProcCountLines("DOMGetDocObjectFromNumber", vbext_pk_Proc))
        strSrc3rd = strSrc3rd & vbCrLf & .Lines(.ProcStartLine("DOMSleepWhileBusy", vbext_pk_Proc), .ProcCountLines("DOMSleepWhileBusy", vbext_pk_Proc))
    End With
    
    
    txtSource.Text = strSrc1st & vbCrLf & strSrc & vbCrLf & strSrc3rd
    
    txtSource.SelStart = 1
    

End Sub
Private Sub cmdAddWait_Click()
    
    Dim strTmp As String
    strTmp = Replace(txtSource.Text, vbCrLf, vbCr)
    txtSource.Text = Left(strTmp, txtSource.SelStart) & vbCrLf & _
                    vbTab & "DOMSleepWhileBusy objIE" & vbCrLf & _
                    Mid(strTmp, txtSource.SelStart + 1)
                    
    
    
    
'    txtSource.Text = Left(txtSource.Text, txtSource.SelStart) & vbCrLf & _
'                     vbTab & "DomSleepWhileBusy objIE" & vbCrLf & _
'                     Mid(txtSource.Text, txtSource.SelStart + 1)
                     
                     
    
End Sub

Private Sub cmdDefine_Click()
    lstObject.List(lstObject.ListIndex, 4) = cmbOperate.List(cmbOperate.ListIndex)
    lvObject.ListItems(lvObject.SelectedItem.Index).SubItems(2) = cmbOperate.List(cmbOperate.ListIndex)
    Select Case cmbOperate.List(cmbOperate.ListIndex)
    Case "Select"
        lstObject.List(lstObject.ListIndex, 5) = cmbSelect.Text
        lvObject.ListItems(lvObject.SelectedItem.Index).SubItems(3) = cmbSelect.Text
    Case "Input"
        lstObject.List(lstObject.ListIndex, 5) = txtInput.Text
        lvObject.ListItems(lvObject.SelectedItem.Index).SubItems(3) = txtInput.Text
    Case "Click"
    Case "GetText"
    End Select
    
End Sub

Private Sub cmdHide_Click()
    blFlgEnd = True
    Unload frmLayered

End Sub

Private Sub cmdAddSleep_Click()
    
    Dim lngCuLine As Long
    Dim lngCuChar As Long
    Dim strBefore As String
    Dim strAfter As String
    Dim i As Long
    

    txtSource.SetFocus
    lngCuLine = txtSource.CurLine
    lngCuChar = 1
    Do Until i >= lngCuLine
        lngCuChar = InStr(lngCuChar, txtSource.Text, vbLf) + 1
        i = i + 1
    Loop
    
    txtSource.SelText = "    DomSleepWhileBusy objIE, 500" & vbCrLf
    
End Sub

Private Sub cmdClear_Click()
    ReDim strctNode(0) As STRCT_NODE
    lstObject.Clear
End Sub

Private Sub cmdDownload_Click()
    
    Dim strFile As String
    
    strFile = Application.GetSaveAsFilename("test.htm", "*.htm,*.htm,*.*,*.*")
    If strFile <> "False" Then
        lblPath.Caption = strFile
        strctNode(lstObject.ListIndex).Operation.Value = lblPath.Caption
        lstObject.List(lstObject.ListIndex, 3) = lblPath.Caption
    End If
    
End Sub

Private Sub cmdShow_Click()
    
    'Excelを最小化
    Me.Hide
    ShowWindow Application.hwnd, SW_MINIMIZE
    
    blFlgEnd = False
    strURL = InputBox("Input URL")
    If strURL = "" Then
        Exit Sub
    End If
    'Application.Visible = False
    
    Set objIE = DOMOpenURL(strURL)
    objIE.Visible = True
    
    DOMSleepWhileBusy objIE
    
    setLayWin
    'Application.Visible = True
    

End Sub

Private Sub cmdShow2_Click()
    
    On Error GoTo ErrHDL
    'Excelを最小化
    ShowWindow Application.hwnd, SW_MINIMIZE
    Me.Hide
    
    blFlgEnd = False
    If objIE Is Nothing Then
        strURL = InputBox("Input URL")
        If strURL = "" Then
            Exit Sub
        End If
        'Application.Visible = False
        
        Set objIE = DOMOpenURL(strURL)
    End If
    objIE.Visible = True
    
    DOMSleepWhileBusy objIE
    
    setLayWin
    'Application.Visible = True

ErrHDL:
    If Err.Number = 462 Then
        MsgBox "IEオブジェクトがありません"
    End If

End Sub

Private Sub cmdSrcAdd_Click()

    Dim strWBK As String
    Dim strMOD As String
    Dim blBKExist As Boolean
    Dim strModName As String
    Dim strSrc As String
    Dim wbk As Workbook
    
    strSrc = txtSource.Text
    strModName = InputBox("Input Workbook name and module name" & vbCrLf & "example) WBKNAME,MODNAME", "MODULE NAME")
    If strModName = "" Then
        Exit Sub
    End If
    
    If UBound(Split(strModName, ",")) = 0 Then
        strMOD = a
    ElseIf UBound(Split(strModName, ",")) >= 1 Then
        strWBK = Split(strModName, ",")(0)
        strMOD = Split(strModName, ",")(1)
    End If
    
    blBKExist = False
    For Each wbk In Workbooks
        If wbk.Name = strWBK Then
            wbk.VBProject.VBComponents.Add(vbext_ct_StdModule).Name = strMOD
            wbk.VBProject.VBComponents(strMOD).CodeModule.InsertLines 1, strSrc
            blBKExist = True
            Exit For
        End If
    Next
    
    If blBKExist = False Then
        Ret = MsgBox("指定のブックが見つかりません" & vbCrLf & "新規にブックを作成しますか?", vbInformation + vbYesNo)
        If Ret = vbYes Then
            Set wbk = Workbooks.Add
            wbk.VBProject.VBComponents.Add(vbext_ct_StdModule).Name = strModName
            wbk.VBProject.VBComponents(strModName).CodeModule.InsertLines 1, strSrc
            Set wbk = Nothing
        End If
    End If
    
    If chkClose.Value = True Then
        Unload Me
    End If
    
End Sub

Private Sub cmdWindow_Click()
    
    ShowWindow Application.hwnd, SW_MINIMIZE
    frmSource.Hide
    'Application.Visible = False
    frmWindow.Show
    
    If Not objIE Is Nothing Then
        blFlgEnd = False
        setLayWin
    End If
    
End Sub
Private Sub lstAttribute_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    
    frmAttribute.txtAttribute = lstAttribute.List(lstAttribute.ListIndex, 0)
    frmAttribute.txtValue = lstAttribute.List(lstAttribute.ListIndex, 1)
    frmAttribute.Show
    
End Sub
Private Sub lstObject_Change()


    ApplyNode strctNode(lstObject.ListIndex)
'    Dim strAttribute() As String
'    Dim i As Long
'
'    lblType.Caption = strctNode(lstObject.ListIndex).NodeType
'
'    txtURL.Text = strctNode(lstObject.ListIndex).URL
'    cmbNodeSelect.Clear
'    If strctNode(lstObject.ListIndex).NodeId <> "" Then
'        cmbNodeSelect.AddItem "Id"
'    End If
'    If strctNode(lstObject.ListIndex).NodeName <> "" Then
'        cmbNodeSelect.AddItem "Name"
'    End If
'    If strctNode(lstObject.ListIndex).FormNodeName <> "" Then
'        cmbNodeSelect.AddItem "Form"
'    End If
'    If strctNode(lstObject.ListIndex).FormNodeNumber <> "" Then
'        cmbNodeSelect.AddItem "FormNumber"
'    End If
'    cmbNodeSelect.AddItem "TagName"
'    cmbNodeSelect.AddItem "NodeNumber"
'    If Not IsNull(lstObject.List(lstObject.ListIndex, 1)) Then
'        cmbNodeSelect.Text = lstObject.List(lstObject.ListIndex, 1)
'    End If
'    strctNode(lstObject.ListIndex).Operation.On = lstObject.Selected(lstObject.ListIndex)
'
'
'    cmbOperate.Clear
'    If cmbNodeSelect.Text = "Form" Or cmbNodeSelect.Text = "FormNumber" Then
'        'cmbOperate.AddItem "Submit"
'        AddOperation "Submit"
'    Else
'        Select Case StrConv(strctNode(lstObject.ListIndex).NodeType, vbLowerCase)
'        Case "input"
'            'cmbOperate.AddItem "Input"
'            AddOperation "Input"
'        Case "select"
'            'cmbOperate.AddItem "Select"
'            AddOperation "Select"
'        Case "button"
'        Case "#text"
'        End Select
'        'cmbOperate.AddItem "GetText"
'        AddOperation "GetText"
'        'cmbOperate.AddItem "Click"
'        AddOperation "Click"
'    End If
'
'
'    If lstObject.List(lstObject.ListIndex, 2) <> "" Then
'        cmbOperate.Text = lstObject.List(lstObject.ListIndex, 2)
'        Select Case cmbOperate.Text
'        Case "Click"
'            txtInput.Visible = False
'            cmbSelect.Visible = False
'        Case "Input"
'            txtInput.Visible = True
'            cmbSelect.Visible = False
'        Case "Select"
'            txtInput.Visible = False
'            cmbSelect.Visible = True
'        Case "GetText"
'            txtInput.Visible = False
'            cmbSelect.Visible = False
'        End Select
'    Else
'        cmbOperate.Text = ""
'    End If
'
'    If lstObject.List(lstObject.ListIndex, 3) <> "" Then
'        If lstObject.List(lstObject.ListIndex, 2) = "Input" Then
'            txtInput.Text = lstObject.List(lstObject.ListIndex, 3)
'        ElseIf lstObject.List(lstObject.ListIndex, 1) = "Select" Then
'            cmbSelect.Text = lstObject.List(lstObject.ListIndex, 3)
'        End If
'    Else
'        txtInput.Text = ""
'        cmbSelect.Text = ""
'    End If
'
'    strAttribute = Split(strctNode(lstObject.ListIndex).Attribute, vbCrLf)
'
'
'    lstAttribute.Clear
'    For i = 0 To UBound(strAttribute)
'        lstAttribute.AddItem
'        lstAttribute.List(i, 0) = Split(strAttribute(i), "=")(0)
'        lstAttribute.List(i, 1) = Split(strAttribute(i), "=")(1)
'        If lstAttribute.List(i, 0) = "href" Then
'            'cmbOperate.AddItem "Download"
'            AddOperation "Download"
'            'cmbOperate.AddItem "Click"
'            AddOperation "Click"
'        ElseIf lstAttribute.List(i, 0) = "type" Then
'            Select Case StrConv(lstAttribute.List(i, 1), vbLowerCase)
'            Case "radio"
'                'cmbOperate.AddItem "Checked"
'                AddOperation "Checked"
'            Case "checkbox"
'                'cmbOperate.AddItem "Checked"
'                AddOperation "Checked"
'            Case "button"
'                'cmbOperate.AddItem "Click"
'                AddOperation "Click"
'            Case "text"
'                'cmbOperate.AddItem "Input"
'                AddOperation "Input"
'            End Select
'        End If
'    Next i
    
End Sub
Private Sub lstObject_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 2 Then
        cbMenu.ShowPopup
    End If
End Sub
Private Sub optInput_Click()
    If optInput.Value = True Then
        optInput.Value = False
        lstObject.List(lstObject.ListIndex, 3) = "False"
    Else
        optInput.Value = True
        lstObject.List(lstObject.ListIndex, 3) = "True"
    End If
End Sub

Private Sub txtInput_Change()
    lstObject.List(lstObject.ListIndex, 3) = txtInput.Text
    strctNode(lstObject.ListIndex).Operation.Value = txtInput.Text
End Sub
Private Sub UserForm_Initialize()
    
    For Each cbMenu In Application.CommandBars
        If cbMenu.Name = "DOMメニュー" Then
            cbMenu.Delete
        End If
    Next
        
    Set cbMenu = Application.CommandBars.Add("DOMメニュー", msoBarPopup, , True)
    With cbMenu.Controls.Add(msoControlButton)
        .Caption = "削除"
        .OnAction = "RemoveItemFromlstObject"
    End With

End Sub

Private Sub UserForm_Layout()

    Dim i As Long
    lstObject.Clear
    If Sgn(strctNode) Then
        For i = 0 To UBound(strctNode)
            If strctNode(i).NodeNumber <> "" Then
                lstObject.AddItem
                lstObject.List(i, 0) = strctNode(i).NodeNumber
                lstObject.List(i, 1) = strctNode(i).Operation.Object
                lstObject.List(i, 2) = strctNode(i).Operation.Operation
                lstObject.List(i, 3) = strctNode(i).Operation.Value
                lstObject.Selected(i) = strctNode(i).Operation.On
            End If
        Next i
    End If


End Sub
Private Sub UserForm_Terminate()
    Application.Visible = True
    If Not objIE Is Nothing Then
        objIE.Quit
        Set objIE = Nothing
    End If
End Sub
Private Function SubItemRect(lngRow As Long, lngCol As Long) As POINTAPI()

    Dim rectSubItem As RECT
    Dim Ret As Long
    Dim ptTmp(1) As POINTAPI
    
    rectSubItem.Top = lngCol
    rectSubItem.Left = LVIR_LABEL
    Ret = SendMessageW(lvObject.hwnd, LVM_GETSUBITEMRECT, ByVal lngRow, VarPtr(rectSubItem))
    
    ptTmp(0).X = rectSubItem.Left
    ptTmp(0).Y = rectSubItem.Top
    ptTmp(1).X = rectSubItem.Right
    ptTmp(1).Y = rectSubItem.Bottom

    SubItemRect = ptTmp
    
End Function
Private Sub AddOperation(strOperation As String)

    Dim i As Long
    Dim blFlg As Boolean
    If cmbOperate.ListCount = 0 Then
        cmbOperate.AddItem strOperation
        Exit Sub
    End If
    For i = 0 To cmbOperate.ListCount - 1
        If cmbOperate.List(i) = strOperation Then
            blFlg = True
            Exit Sub
        End If
    Next i
    If blFlg = False Then
        cmbOperate.AddItem strOperation
    End If

End Sub
Private Sub ApplyNode(tagNode As STRCT_NODE)

    Dim strAttribute() As String
    Dim i As Long

    lblType.Caption = tagNode.NodeType
    
    txtURL.Text = tagNode.URL
    cmbNodeSelect.Clear
    If tagNode.NodeId <> "" Then
        cmbNodeSelect.AddItem "Id"
    End If
    If tagNode.NodeName <> "" Then
        cmbNodeSelect.AddItem "Name"
    End If
    If tagNode.FormNodeName <> "" Then
        cmbNodeSelect.AddItem "Form"
    End If
    If tagNode.FormNodeNumber <> "" Then
        cmbNodeSelect.AddItem "FormNumber"
    End If
    cmbNodeSelect.AddItem "TagName"
    cmbNodeSelect.AddItem "NodeNumber"
    If Not IsNull(lstObject.List(lstObject.ListIndex, 1)) Then
        cmbNodeSelect.Text = lstObject.List(lstObject.ListIndex, 1)
    End If
    tagNode.Operation.On = lstObject.Selected(lstObject.ListIndex)
    
    
    SetOperate tagNode
    

    If lstObject.List(lstObject.ListIndex, 2) <> "" Then
        cmbOperate.Text = lstObject.List(lstObject.ListIndex, 2)
        Select Case cmbOperate.Text
        Case "Click"
            txtInput.Visible = False
            cmbSelect.Visible = False
        Case "Input"
            txtInput.Visible = True
            cmbSelect.Visible = False
        Case "Select"
            txtInput.Visible = False
            cmbSelect.Visible = True
        Case "GetText"
            txtInput.Visible = False
            cmbSelect.Visible = False
        End Select
    Else
        cmbOperate.Text = ""
    End If
    
    If lstObject.List(lstObject.ListIndex, 3) <> "" Then
        If lstObject.List(lstObject.ListIndex, 2) = "Input" Then
            txtInput.Text = lstObject.List(lstObject.ListIndex, 3)
        ElseIf lstObject.List(lstObject.ListIndex, 1) = "Select" Then
            cmbSelect.Text = lstObject.List(lstObject.ListIndex, 3)
        End If
    Else
        txtInput.Text = ""
        cmbSelect.Text = ""
    End If
    

End Sub

Private Sub SetOperate(tagNode As STRCT_NODE)
    
    cmbOperate.Clear
    If cmbNodeSelect.Text = "Form" Or cmbNodeSelect.Text = "FormNumber" Then
        'cmbOperate.AddItem "Submit"
        AddOperation "Submit"
    Else
        Select Case StrConv(tagNode.NodeType, vbLowerCase)
        Case "input"
            'cmbOperate.AddItem "Input"
            AddOperation "Input"
        Case "select"
            'cmbOperate.AddItem "Select"
            AddOperation "Select"
        Case "button"
        Case "#text"
        End Select
        'cmbOperate.AddItem "GetText"
        AddOperation "GetText"
        'cmbOperate.AddItem "Click"
        AddOperation "Click"
    End If
    strAttribute = Split(tagNode.Attribute, vbCrLf)
    lstAttribute.Clear
    For i = 0 To UBound(strAttribute)
        lstAttribute.AddItem
        lstAttribute.List(i, 0) = Split(strAttribute(i), "=")(0)
        lstAttribute.List(i, 1) = Split(strAttribute(i), "=")(1)
        If lstAttribute.List(i, 0) = "href" Or lstAttribute.List(i, 0) = "src" Then
            'cmbOperate.AddItem "Download"
            AddOperation "Download"
            'cmbOperate.AddItem "Click"
            AddOperation "Click"
        ElseIf lstAttribute.List(i, 0) = "type" Then
            Select Case StrConv(lstAttribute.List(i, 1), vbLowerCase)
            Case "radio"
                'cmbOperate.AddItem "Checked"
                AddOperation "Checked"
            Case "checkbox"
                'cmbOperate.AddItem "Checked"
                AddOperation "Checked"
            Case "button"
                'cmbOperate.AddItem "Click"
                AddOperation "Click"
            Case "text"
                'cmbOperate.AddItem "Input"
                AddOperation "Input"
            End Select
        End If
    Next i

End Sub
