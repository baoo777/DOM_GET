Attribute VB_Name = "modDomAccess"
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
' /////////////////////////////////////////////////////////////////////////////////////
' //#名称 TestDOMOpenURL
' //
' //#概要 DOMOpenURLのテスト用
' //
' //#引数
' //
' //#戻値 なし
' //
' //#解説 DOMOpenURLを実行する
' //
' //#履歴 2012/06/06
' //      Coded by YASUTADA OOBA
' //
' /////////////////////////////////////////////////////////////////////////////////////
Sub TestDOMOpenURL()

    Dim objDOM As Object
    Set objDOM = DOMOpenURL("www.yahoo.co.jp")
    
    
End Sub
' /////////////////////////////////////////////////////////////////////////////////////
' //#名称 DOMOpenURL
' //
' //#概要 指定のURLをIEで開く
' //
' //#引数 strURL
' //
' //#戻値 開いたIEオブジェクト
' //
' //#解説
' //
' //#履歴 2012/06/06
' //      Coded by YASUTADA OOBA
' //
' /////////////////////////////////////////////////////////////////////////////////////
Public Function DOMOpenURL(strURL As String) As Object
    
    Dim objIE As Object
    
    Set objIE = CreateObject("InternetExplorer.Application")
    objIE.Navigate2 strURL
    DOMSleepWhileBusy objIE
    'objIE.Visible = True
    Set DOMOpenURL = objIE
    

End Function
' /////////////////////////////////////////////////////////////////////////////////////
' //#名称 TestDOMGetFromURL
' //
' //#概要 DOMGetFromURLのテスト用
' //
' //#引数
' //
' //#戻値 なし
' //
' //#解説 DOMGetFromURLを実施して取得したウィンドウタイトルをデバックウィンドウに表示する。
' //
' //#履歴 2012/06/06
' //      Coded by YASUTADA OOBA
' //
' /////////////////////////////////////////////////////////////////////////////////////
Sub TestDOMGetFromURL()
    
    Dim objDOM As Object
    Set objDOM = DOMGetFromURL("www.yahoo.co.jp")
    Debug.Print objDOM.Document.Title
    
End Sub
' /////////////////////////////////////////////////////////////////////////////////////
' //#名称 DOMGetFromURL
' //
' //#概要 指定のURLのIEウィンドウからIEオブジェクトを取得する
' //
' //#引数 strURL
' //
' //#戻値 指定のURLのIEオブジェクト
' //
' //#解説 すでに開いているウィンドウ群の中から指定のURLのIEオブジェクトを取得する。
' //
' //#履歴 2012/06/06
' //      Coded by YASUTADA OOBA
' //
' /////////////////////////////////////////////////////////////////////////////////////
Public Function DOMGetFromURL(strURL As String) As Object
    
    Dim objShell As Object
    Dim objWindow As Object
    Dim objTmp As Object
    
    Set objShell = CreateObject("Shell.Application")
    
    For Each objWindow In objShell.Windows
        If TypeName(objWindow.Document) = "HTMLDocument" Then
            If InStr(1, objWindow.Document.URL, strURL) <> 0 Then
                Set objTmp = objWindow '.Document
            End If
        End If
    Next
    
    If objTmp Is Nothing Then
        Set DOMGetFromURL = DOMOpenURL(strURL)
    Else
        Set DOMGetFromURL = objTmp
    End If
    
    
End Function
' /////////////////////////////////////////////////////////////////////////////////////
' //#名称 TestDOMGetFromWindowTitle
' //
' //#概要 DOMGetFromWindowTitleを実施する
' //
' //#引数
' //
' //#戻値 なし
' //
' //#解説 DOMGetFromWindowTitleを実施して取得したIEオブジェクトからURLをデバックウィンドウに表示する
' //
' //#履歴 2012/06/06
' //      Coded by YASUTADA OOBA
' //
' /////////////////////////////////////////////////////////////////////////////////////
Sub TestDOMGetFromWindowTitle()
    
    Dim objDOM As Object
    
    Set objDOM = DOMGetFromWindowTitle("Yahoo")
    Debug.Print objDOM.Document.Title
    
End Sub
' /////////////////////////////////////////////////////////////////////////////////////
' //#名称 DOMGetFromWindowTitle
' //
' //#概要 指定のウィンドウタイトルのIEオブジェクトを取得する
' //
' //#引数 strTitle
' //
' //#戻値 IEオブジェクト
' //
' //#解説 全てのウィンドウのなかから指定のURLを含むウィンドウを取得し、IEオブジェクトｌを取得する。
' // 　　 最初に見つかったものを返す。
' //
' //#履歴 2012/06/06
' //      Coded by YASUTADA OOBA
' //
' /////////////////////////////////////////////////////////////////////////////////////
Public Function DOMGetFromWindowTitle(strTitle As String) As Object
    
    Dim objShell As Object
    Dim objWindow As Object
    
    Set objShell = CreateObject("Shell.Application")
    For Each objWindow In objShell.Windows
        If TypeName(objWindow.Document) = "HTMLDocument" Then
            If InStr(1, objWindow.Document.Title, strTitle) <> 0 Then
                Set DOMGetFromWindowTitle = objWindow '.Document
                Exit Function
            End If
        End If
    Next
    
End Function
' /////////////////////////////////////////////////////////////////////////////////////
' //#名称 DOMGetDocObjectFromNumber
' //
' //#概要 ルートからノード番号をカンマ区切りで指定して得られたドキュメントオブジェクトを返す
' //
' //#引数 objIE IEオブジェクト
' //      strDomTree    カンマ区切りのノード番号
' //
' //#戻値 ドキュメントオブジェクト
' //
' //#解説
' //
' //#履歴 2012/06/06
' //      Coded by YASUTADA OOBA
' //
' /////////////////////////////////////////////////////////////////////////////////////
Public Function DOMGetDocObjectFromNumber(objIE As Object, strDOMTree As String) As Object

    On Error GoTo ErrHDL
    
    Dim objTmp As Object
    Dim i As Long
    Dim tmpNodeNumber() As String
    Set objTmp = objIE.Document
    tmpNodeNumber = Split(strDOMTree, ",")
    
    For i = 0 To UBound(tmpNodeNumber)
        Set objTmp = objTmp.ChildNodes.Item(tmpNodeNumber(i))
    Next i
    Set DOMGetDocObjectFromNumber = objTmp
    Exit Function
    
ErrHDL:
    Set DOMGetDocObjectFromNumber = Nothing

End Function
' /////////////////////////////////////////////////////////////////////////////////////
' //#名称 DOMSleepWhileBusy
' //
' //#概要 指定のIEオブジェクトがビジーの間待つ
' //
' //#引数 objIE IEオブジェクト
' //      lngFirstWait 最初に待つ時間(ClickなどBusyを想定していないものはBusyになる前に
' //                   次に遷移する場合がある。
' //
' //#戻値 なし
' //
' //#解説
' //
' //#履歴 2012/06/06
' //      Coded by YASUTADA OOBA
' //
' /////////////////////////////////////////////////////////////////////////////////////
Public Function DOMSleepWhileBusy(objIE As Object, Optional lngFirstWait As Long = -1)

        
    Dim i As Long
    DOMSleepWhileBusy = True
    If lngFirstWait <> -1 Then
        Sleep (lngFirstWait)
    End If
    
    i = 0
    Do Until objIE.Busy = False And objIE.ReadyState = 4 Or i > 10
        Sleep (1000)
        i = i + 1
    Loop
    
    If i > 10 Then
        MsgBox ("タイムアウト")
        DOMSleepWhileBusy = False
    End If

End Function

