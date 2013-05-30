Attribute VB_Name = "modBase"
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Long, ByVal dwFlags As Long) As Long
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal uFlags As Long) As Long
Private Declare Function ShowWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Const SW_SHOWMINIMIZED = 2
Private Const SW_SHOWNORMAL = 1
Private Const SW_MINIMIZE = 6
Private Const SW_RESTORE = 9

Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetAsyncKeyState Lib "user32.dll" (ByVal vKey As Long) As Long

Private Const VK_CONTROL As Long = &H11
Private Type RECT
    Left As Long
    Top As Long
    Width As Long
    Height As Long
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = (-20)                   '拡張ウィンドウスタイル
Private Const LWA_COLORKEY = 1                      'crKeyを透明色として使う
Private Const LWA_ALPHA = 2                         'bAlphaをアルファー値として使う

Private Const WS_CLIPSIBLINGS = &H4000000
Private Const WS_VISIBLE = &H10000000
Private Const WS_POPUP = &H80000000
Private Const WS_EX_LAYERED = &H80000                '拡張ウィンドウ属
Private Const WS_EX_TRANSPARENT = &H20
Private Const WS_EX_WINDOWEDGE = &H100
Private Const WS_BORDER = &H800000
Private Const WS_DLGFRAME = &H400000
Private Const WS_SYSMENU = &H80000

Public Type STRCT_SELECT_NODE
    Caption As String
    Value As String
End Type
Public Type OPERATION_NODE
    Object As String 'ENM_OPERATION_OBJECT
    Operation As String 'ENM_OPERATION_OPERATE
    Value As String
    On As Boolean
End Type
Public Enum ENM_OPERATION_OBJECT
    enmNodeNumber = 0
    enmID
    enmName
    enmTagName
    enmForm
    enmFormNumber
End Enum
Public Enum ENM_OPERATION_OPERATE
    enmClick
    enmInput
    enmGetText
    enmChecked
    enmSelect
    enmDownload
    enmSubmit
End Enum
Public Type STRCT_NODE
    Id As Long
    URL As String
    NodeNumber As String
    NodeName As String
    NodeId As String
    NodeType As String
    FormNodeName As String
    FormNodeNumber As String
    Attribute As String
    NodeElementByName As String
    NodeElementByTag As String
    href As String
    src As String
    SelectList() As STRCT_SELECT_NODE
    Operation As OPERATION_NODE
End Type

Public objIE As Object
Public objGNode As Object
Public blFlgEnd As Boolean
Public strctNode() As STRCT_NODE
Sub Auto_Open()
    Dim cb As CommandBar
    For Each cb In Application.CommandBars
        If cb.Name = "DomNew" Then
            cb.Delete
        End If
    Next
    
    Set cb = Application.CommandBars.Add("DomNew", msoBarTop, , True)
    With cb.Controls.Add(msoControlButton)
        .Caption = "URL"
        .style = msoButtonCaption
        .OnAction = "GetIEFromURL"
    End With
    With cb.Controls.Add(msoControlButton)
        .Caption = "Title"
        .style = msoButtonCaption
        .OnAction = "GetIEFromTitle"
    End With
    With cb.Controls.Add(msoControlButton)
        .Caption = "MainForm"
        .style = msoButtonCaption
        .OnAction = "ShowForm"
    End With
    cb.Visible = True
    
End Sub
' /////////////////////////////////////////////////////////////////////////////////////
' //#名称 GetIEFromURL
' //
' //#概要 Windowタイトルを選択してDOM操作を開始する
' //
' //#引数 なし
' //
' //#戻値 なし
' //
' //#解説 URLメニューから呼び出される
' //
' //#履歴 2013/04/27
' //      Coded by YASUTADA OOBA
' //
' /////////////////////////////////////////////////////////////////////////////////////
Sub GetIEFromTitle()
    
    Load frmSource
    ShowWindow Application.hwnd, SW_MINIMIZE
    'Application.Visible = False
    frmWindow.Show
    
    If Not objIE Is Nothing Then
        setLayWin
    End If
    
End Sub
' /////////////////////////////////////////////////////////////////////////////////////
' //#名称 GetIEFromURL
' //
' //#概要 URLを入力してDOM取得操作を開始する
' //
' //#引数 なし
' //
' //#戻値 なし
' //
' //#解説 URLメニューから呼び出される
' //
' //#履歴 2013/04/27
' //      Coded by YASUTADA OOBA
' //
' /////////////////////////////////////////////////////////////////////////////////////
Sub GetIEFromURL()

    Dim strURL As String
    
    Load frmSource
    blFlgEnd = False
    strURL = InputBox("Input URL")
    If strURL = "" Then
        Exit Sub
    End If
    
    ShowWindow Application.hwnd, SW_MINIMIZE
    'Application.Visible = False
    
    Set objIE = DOMOpenURL(strURL)
    objIE.Visible = True
    DOMSleepWhileBusy objIE
    
    setLayWin
    
End Sub
' /////////////////////////////////////////////////////////////////////////////////////
' //#名称 MyMouseEvent
' //
' //#概要 Dom情報取得中はこのプロシージャの無限ロープとなる
' //
' //#引数 なし
' //
' //#戻値 なし
' //
' //#解説 frmLayeredを表示してループに入る。マウス下のDOMオブジェクトを取得し、その座標
' //      を取得してfrmLayeredをその位置に移動する。常にマウス下にfrmLayeredがある状態。
' //      frmLayeredをダブルクリックして終了する。
' //
' //#履歴 2013/04/27
' //      Coded by YASUTADA OOBA
' //
' /////////////////////////////////////////////////////////////////////////////////////
Sub MyMouseEvent()

    Dim pt As POINTAPI
    Dim myHWND As Long
    Dim frmhWnd As Long
    Dim rct As RECT
    Dim strctRECT As RECT
    Dim tmpRECT As RECT
    Dim objNode As Object
    Dim objAttr As Object
    Dim rctDesk As RECT
    Dim Ret As Long
    Dim blKeyOn As Boolean
    Dim blKeyOld As Boolean
    Dim lngKeyOn As Long
    
    On Error Resume Next
    
    'IEのブラウザ枠のハンドルを取得
    myHWND = FindWindowEx(objIE.hwnd, 0, "Frame Tab", vbNullString)
    myHWND = FindWindowEx(myHWND, 0, "TabWindowClass", vbNullString)
    myHWND = FindWindowEx(myHWND, 0, "Shell DocObject View", vbNullString)
    myHWND = FindWindowEx(myHWND, 0, "Internet Explorer_Server", vbNullString)
    
    'frmLayerdのハンドル取得
    frmhWnd = FindWindowEx(GetDesktopWindow(), 0, "ThunderDFrame", "")
    
    GetWindowRect GetDesktopWindow, rctDesk
    GetWindowRect myHWND, rct
    
    Do
        DoEvents
        GetCursorPos pt
        SetForegroundWindow frmhWnd
        Ret = GetAsyncKeyState(VK_CONTROL)
        If blKeyOld = False And Ret <> 0 Then
            If blKeyOn = False Then
                blKeyOn = True
                'SetLayeredWindowAttributes frmhWnd, 0, 0, LWA_ALPHA
                Unload frmLayered
                blFlgEnd = True
                frmContinue.Show
            Else
                blKeyOn = False
                'SetLayeredWindowAttributes frmhWnd, 0, 150, LWA_ALPHA
            End If
        End If
        If Ret <> 0 Then
            blKeyOld = True
        Else
            blKeyOld = False
        End If
        
        If rct.Left < pt.X And pt.X < rct.Left + rct.Width And rct.Top < pt.Y And pt.Y < rct.Top + rct.Height Then
            Set objGNode = objIE.Document.ElementFromPoint(pt.X - rct.Left, pt.Y - rct.Top)
            tmpRECT = GetBoundingClientRectEX(objGNode)
            
            strctRECT.Left = rct.Left + tmpRECT.Left
            strctRECT.Top = rct.Top + tmpRECT.Top
            strctRECT.Width = tmpRECT.Width
            strctRECT.Height = tmpRECT.Height
            frmLayered.Label1.Caption = objGNode.tagname
            ResizeLay strctRECT
        End If
        If blFlgEnd = True Then
            Exit Do
        End If
    Loop
    
End Sub
' /////////////////////////////////////////////////////////////////////////////////////
' //#名称 GetBoundingClientRectEX
' //
' //#概要 指定のNodeオブジェクトの座標を取得する
' //
' //#引数 objNode Nodeオブジェクト
' //
' //#戻値 指定のオブジェクトを示すGetElementsByNameで得られた配列の添え字
' //
' //#解説 DOMのGetBoundingClientRectメソッドの代用(IE8ぐらいからGetBoundingClientRectが
' //      正しく取得できないので作成
' //
' //#履歴 2013/04/27
' //      Coded by YASUTADA OOBA
' //
' /////////////////////////////////////////////////////////////////////////////////////
Function GetBoundingClientRectEX(objElement) As RECT

    Dim tmpRECT As RECT
    Dim objTmp As Object
    Dim lngScrollX As Long
    Dim lngScrollY As Long
    
    Set objTmp = objElement
    Do Until objTmp.NodeName = "BODY"
        tmpRECT.Left = tmpRECT.Left + objTmp.offsetLeft
        tmpRECT.Top = tmpRECT.Top + objTmp.offsetTop
        Set objTmp = objTmp.offsetParent
    Loop
    lngScrollX = objIE.Document.DocumentElement.ScrollLeft
    lngScrollY = objIE.Document.DocumentElement.ScrollTop
    
    tmpRECT.Left = tmpRECT.Left - lngScrollX
    tmpRECT.Top = tmpRECT.Top - lngScrollY
    tmpRECT.Width = objElement.offsetWidth
    tmpRECT.Height = objElement.offsetHeight
    
    GetBoundingClientRectEX = tmpRECT
    
End Function
' /////////////////////////////////////////////////////////////////////////////////////
' //#名称 GetElementsByName
' //
' //#概要 指定のNodeオブジェクトがRoot(Document)オブジェクトから同じ名前で何番目かを
' //      返す。
' //
' //#引数 objNode Nodeオブジェクト
' //
' //#戻値 指定のオブジェクトを示すGetElementsByNameで得られた配列の添え字
' //
' //#解説 NodeオブジェクトをGetElementsByNameで示すソース生成に使用
' //
' //#履歴 2012/06/06
' //      Coded by YASUTADA OOBA
' //
' /////////////////////////////////////////////////////////////////////////////////////
Function GetElementsByName(objNode As Object) As Long
    Dim strName As String
    Dim i As Long
    
    strName = objNode.getAttribute("name")
    For i = 0 To objNode.Document.GetElementsByTagName(strName).Length
        If objNode Is objNode.Document.GetElementsByName(strName).Item(i) Then
            GetElementsByName = i
            Exit For
        End If
    Next i
    
End Function
' /////////////////////////////////////////////////////////////////////////////////////
' //#名称 GetElementsByTagName
' //
' //#概要 指定のNodeオブジェクトがRoot(Document)オブジェクトから同じタグで何番目かを
' //      返す。
' //
' //#引数 objNode Nodeオブジェクト
' //
' //#戻値 指定のオブジェクトを示すGetElementsByTagNameで得られた配列の添え字
' //
' //#解説 NodeオブジェクトをGetElementsByTagNameで示すソース生成に使用
' //
' //#履歴 2012/06/06
' //      Coded by YASUTADA OOBA
' //
' /////////////////////////////////////////////////////////////////////////////////////
Function GetElementsByTagName(objNode As Object) As Long
    Dim strTag As String
    Dim i As Long
    
    strTag = objNode.tagname
    For i = 0 To objNode.Document.GetElementsByTagName(strTag).Length
        If objNode Is objNode.Document.GetElementsByTagName(strTag).Item(i) Then
            GetElementsByTagName = i
            Exit For
        End If
    Next i
    
End Function
' /////////////////////////////////////////////////////////////////////////////////////
' //#名称 GetNodeString
' //
' //#概要 Root(Document)オブジェクトから指定のNodeオブジェクトまでの分岐の番号を積み上
' //      げた番号をカンマ区切りで取得する
' //
' //#引数 objNode Nodeオブジェクト
' //
' //#戻値 なし
' //
' //#解説 DOMOpenURLを実行する
' //
' //#履歴 2013/04/27
' //      Coded by YASUTADA OOBA
' //
' /////////////////////////////////////////////////////////////////////////////////////
Function GetNodeString(objNode As Object) As String
    Dim objTmp As Object
    Dim strNode As String
    Dim i As Long
    i = 0
    
    Set objTmp = objNode
    Do While Not objTmp Is objNode.OwnerDocument
        Do
            If objTmp Is objTmp.ParentNode.ChildNodes.Item(i) Then
                If strNode = "" Then
                    strNode = CStr(i)
                Else
                    strNode = CStr(i) & "," & strNode
                End If
                i = 0
                Exit Do
            End If
            i = i + 1
        Loop
        Set objTmp = objTmp.ParentNode
    Loop
    
    GetNodeString = strNode
    
End Function
' /////////////////////////////////////////////////////////////////////////////////////
' //#名称 setLayWin
' //
' //#概要 DOM取得操作の初期設定を実施する
' //
' //#引数 なし
' //
' //#戻値 なし
' //
' //#解説 frmLayeredウィンドウのスタイルを変更してMyMouseEventを呼び出す。
' //
' //#履歴 2013/04/27
' //      Coded by YASUTADA OOBA
' //
' /////////////////////////////////////////////////////////////////////////////////////
Sub setLayWin()
    Dim dStyle As Long
    Dim dStyleEx As Long
    Dim hwnd As Long
    Dim Ret As Long
    Dim t As Single
    Dim u2Hdc As Long
    Dim strctRECT As RECT
    Dim frmCon As UserForm
    
    For Each frmCon In UserForms
        If TypeName(frmCon) = "frmContinue" Then
            Unload frmCon
        End If
    Next
    
    frmLayered.Show vbModeless
    frmLayered.Width = 240
    frmLayered.Height = 180
    frmLayered.Label1.Left = 0
    frmLayered.Label1.Top = 0
    frmLayered.Label1.Width = frmLayered.Width - 5
    frmLayered.Label1.Height = frmLayered.Height - 21

    hwnd = FindWindowEx(GetDesktopWindow(), 0, "ThunderDFrame", "")
    
    SetForegroundWindow hwnd
    
    dStyle = GetWindowLong(hwnd, GWL_STYLE)
    dStyle = dStyle Or WS_CLIPSIBLINGS
    dStyle = dStyle Or WS_POPUP
    dStyle = dStyle Xor WS_BORDER
    dStyle = dStyle Xor WS_DLGFRAME
    
    dStyleEx = GetWindowLong(hwnd, GWL_EXSTYLE)
    dStyleEx = dStyle Xor WS_EX_WINDOWEDGE
    dStyleEx = dStyleEx Or WS_EX_LAYERED
    
    SetWindowLong hwnd, GWL_STYLE, dStyle
    SetWindowLong hwnd, GWL_EXSTYLE, dStyleEx
    
    SetLayeredWindowAttributes hwnd, 0, 150, LWA_ALPHA
    'SetLayeredWindowAttributes hWnd, &HFFFFFF, 150, LWA_COLORKEY
    
    strctRECT.Left = 10
    strctRECT.Top = 10
    strctRECT.Width = 100 'frmLayered.Width
    strctRECT.Height = 100 ' frmLayered.Height
    
    
    MyMouseEvent
        
    
End Sub
' /////////////////////////////////////////////////////////////////////////////////////
' //#名称 ResizeLay
' //
' //#概要 frmLayeredウィンドウの位置サイズを設定する
' //
' //#引数 なし
' //
' //#戻値 なし
' //
' //#解説 MyMouseEvent内のループでWeb部品の位置、サイズを取得して呼び出される
' //
' //#履歴 2013/04/27
' //      Coded by YASUTADA OOBA
' //
' /////////////////////////////////////////////////////////////////////////////////////
Sub ResizeLay(strctRECT As RECT)
    
    Dim hwnd As Long
    Const DPI_PER_PPI As Single = 0.75
    
    hwnd = FindWindowEx(GetDesktopWindow(), 0, "ThunderDFrame", "")
    SetWindowPos hwnd, HWND_TOP, strctRECT.Left, strctRECT.Top, strctRECT.Width, strctRECT.Height, SWP_SHOWWINDOW
    'SetForegroundWindow hwnd
    
    frmLayered.Label1.Left = 0 'strctRECT.Left
    frmLayered.Label1.Top = 0 'strctRECT.Top
    frmLayered.Label1.Width = strctRECT.Width * DPI_PER_PPI
    frmLayered.Label1.Height = strctRECT.Height * DPI_PER_PPI
    
    
End Sub
' /////////////////////////////////////////////////////////////////////////////////////
' //#名称 RemoveItemFromlstObject
' //
' //#概要
' //
' //#引数 なし
' //
' //#戻値 なし
' //
' //#解説 MyMouseEvent内のループでWeb部品の位置、サイズを取得して呼び出される
' //
' //#履歴 2013/04/27
' //      Coded by YASUTADA OOBA
' //
' /////////////////////////////////////////////////////////////////////////////////////
Public Sub RemoveItemFromlstObject()
    Dim i As Long
    
    If UBound(strctNode) = 0 Then
        ReDim strctNode(0) As STRCT_NODE
    Else
        RemoveItemFromArry strctNode, frmSource.lstObject.ListIndex
    End If
    
    frmSource.lstObject.Clear
    If Sgn(strctNode) Then
        For i = 0 To UBound(strctNode)
            frmSource.lstObject.AddItem
            frmSource.lstObject.List(i, 0) = strctNode(i).NodeNumber
            frmSource.lstObject.List(i, 1) = strctNode(i).Operation.Object
            frmSource.lstObject.List(i, 2) = strctNode(i).Operation.Operation
            frmSource.lstObject.List(i, 3) = strctNode(i).Operation.Value
            frmSource.lstObject.Selected(i) = strctNode(i).Operation.On
        Next i
    End If
End Sub
' /////////////////////////////////////////////////////////////////////////////////////
' //#名称 RemoveItemFromArry
' //
' //#概要 Node情報構造体の配列から指定の添え字の構造体を削除する。
' //
' //#引数 arr():Node情報構造体の配列
' //      lngRemove:削除する構造体の添え字
' //
' //#戻値 なし
' //
' //#解説
' //
' //#履歴 2013/04/27
' //      Coded by YASUTADA OOBA
' //
' /////////////////////////////////////////////////////////////////////////////////////
Sub RemoveItemFromArry(ByRef arr() As STRCT_NODE, lngRemove As Long)
    
    Dim i As Long
    For i = lngRemove To UBound(arr) - 1
        arr(i) = arr(i + 1)
        i = i + 1
    Next i
    ReDim Preserve arr(UBound(arr) - 1)
    
End Sub
' /////////////////////////////////////////////////////////////////////////////////////
' //#名称 ShowForm
' //
' //#概要 メインウィンドウを表示する
' //
' //#引数 なし
' //
' //#戻値 なし
' //
' //#解説
' //
' //#履歴 2013/04/27
' //      Coded by YASUTADA OOBA
' //
' /////////////////////////////////////////////////////////////////////////////////////
Sub ShowForm()
    frmSource.Show
    
End Sub
