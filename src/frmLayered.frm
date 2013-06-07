VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLayered 
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   "frmLayered.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmLayered"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShowWindow Lib "user32.dll" (ByVal hwnd As _
Long, ByVal nCmdShow As Long) As Long
Private Const SW_SHOWMINIMIZED = 2
Private Const SW_SHOWNORMAL = 1
Private Const SW_MINIMIZE = 6
Private Const SW_RESTORE = 9

Private Sub Label1_Click()
    
    On Error Resume Next
        
    If Sgn(strctNode) = 0 Then
        ReDim strctNode(0) As STRCT_NODE
    Else
        If UBound(strctNode) = 0 And strctNode(0).NodeNumber = "" Then
            ReDim strctNode(0) As STRCT_NODE
        Else
            ReDim Preserve strctNode(UBound(strctNode) + 1) As STRCT_NODE
        End If
    End If
    strctNode(UBound(strctNode)).URL = objGNode.Document.URL
    strctNode(UBound(strctNode)).Id = UBound(strctNode)
    strctNode(UBound(strctNode)).NodeNumber = GetNodeString(objGNode)
    strctNode(UBound(strctNode)).NodeId = objGNode.getAttribute("id")
    strctNode(UBound(strctNode)).NodeName = objGNode.getAttribute("name")
    strctNode(UBound(strctNode)).NodeType = objGNode.tagname
    If objGNode.tagname = "SELECT" Then
        For i = 0 To objGNode.ChildNodes.Length
            If objGNode.ChildNodes.Item(i).NodeName = "OPTION" Then
                If objGNode.ChildNodes.Item(i).ChildNodes.Item(0).NodeName = "#text" Then
                    If Sgn(strctNode(UBound(strctNode)).SelectList) = 0 Then
                        ReDim strctNode(UBound(strctNode)).SelectList(0) As STRCT_SELECT_NODE
                        strctNode(UBound(strctNode)).SelectList(0).Caption = objGNode.ChildNodes.Item(i).ChildNodes.Item(0).NodeValue
                        If objGNode.ChildNodes.Item(i).Value = 0 Then
                            strctNode(UBound(strctNode)).SelectList(0).Value = objGNode.ChildNodes.Item(i).ChildNodes.Item(0).NodeValue
                        Else
                            strctNode(UBound(strctNode)).SelectList(0).Value = objGNode.ChildNodes.Item(i).Value
                        End If
                    Else
                        ReDim Preserve strctNode(UBound(strctNode)).SelectList(UBound(strctNode(UBound(strctNode)).SelectList) + 1) As STRCT_SELECT_NODE
                        strctNode(UBound(strctNode)).SelectList(UBound(strctNode(UBound(strctNode)).SelectList)).Caption = objGNode.ChildNodes.Item(i).ChildNodes.Item(0).NodeValue
                        If objGNode.ChildNodes.Item(i).Value = 0 Then
                            strctNode(UBound(strctNode)).SelectList(UBound(strctNode(UBound(strctNode)).SelectList)).Value = objGNode.ChildNodes.Item(i).ChildNodes.Item(0).NodeValue
                        Else
                            strctNode(UBound(strctNode)).SelectList(UBound(strctNode(UBound(strctNode)).SelectList)).Value = objGNode.ChildNodes.Item(i).Value
                        End If
                    End If
                End If
            End If
        Next i
    End If
    strctNode(UBound(strctNode)).FormNodeName = objGNode.form.getAttribute("name")
    strctNode(UBound(strctNode)).FormNodeNumber = GetNodeString(objGNode.form)
    strctNode(UBound(strctNode)).href = objGNode.getAttribute("href")
    strctNode(UBound(strctNode)).src = objGNode.getAttribute("src")
    
    strctNode(UBound(strctNode)).Attribute = ""
    For Each objAttr In objGNode.Attributes
        'If objAttr.specified = True Then
            If objAttr.Name = "name" Then
                strctNode(UBound(strctNode)).NodeName = objAttr.Value
            End If
            If objAttr.Name = "id" Then
                strctNode(UBound(strctNode)).NodeId = objAttr.Value
            End If
            
            If objAttr.Value <> "" Then
                If strctNode(UBound(strctNode)).Attribute = "" Then
                    strctNode(UBound(strctNode)).Attribute = objAttr.NodeName & "=" & objAttr.Value
                Else
                    strctNode(UBound(strctNode)).Attribute = strctNode(UBound(strctNode)).Attribute & vbCrLf & objAttr.NodeName & "=" & objAttr.Value
                End If
            End If
        'End If
    Next
    strctNode(UBound(strctNode)).NodeElementByName = GetElementsByName(objGNode)
    strctNode(UBound(strctNode)).NodeElementByTag = GetElementsByTagName(objGNode)
        
End Sub

Private Sub Label1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    blFlgEnd = True
    
    objIE.Visible = False
    
    frmSource.Left = frmSource.Left + 1
    frmSource.Left = frmSource.Left - 1
    
    ShowWindow Application.hwnd, SW_RESTORE
    Unload Me
    If frmSource.Visible = False Then
        frmSource.Show
    End If
    
End Sub
