Attribute VB_Name = "modDomAccess"
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
' /////////////////////////////////////////////////////////////////////////////////////
' //#���� TestDOMOpenURL
' //
' //#�T�v DOMOpenURL�̃e�X�g�p
' //
' //#����
' //
' //#�ߒl �Ȃ�
' //
' //#��� DOMOpenURL�����s����
' //
' //#���� 2012/06/06
' //      Coded by YASUTADA OOBA
' //
' /////////////////////////////////////////////////////////////////////////////////////
Sub TestDOMOpenURL()

    Dim objDOM As Object
    Set objDOM = DOMOpenURL("www.yahoo.co.jp")
    
    
End Sub
' /////////////////////////////////////////////////////////////////////////////////////
' //#���� DOMOpenURL
' //
' //#�T�v �w���URL��IE�ŊJ��
' //
' //#���� strURL
' //
' //#�ߒl �J����IE�I�u�W�F�N�g
' //
' //#���
' //
' //#���� 2012/06/06
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
' //#���� TestDOMGetFromURL
' //
' //#�T�v DOMGetFromURL�̃e�X�g�p
' //
' //#����
' //
' //#�ߒl �Ȃ�
' //
' //#��� DOMGetFromURL�����{���Ď擾�����E�B���h�E�^�C�g�����f�o�b�N�E�B���h�E�ɕ\������B
' //
' //#���� 2012/06/06
' //      Coded by YASUTADA OOBA
' //
' /////////////////////////////////////////////////////////////////////////////////////
Sub TestDOMGetFromURL()
    
    Dim objDOM As Object
    Set objDOM = DOMGetFromURL("www.yahoo.co.jp")
    Debug.Print objDOM.Document.Title
    
End Sub
' /////////////////////////////////////////////////////////////////////////////////////
' //#���� DOMGetFromURL
' //
' //#�T�v �w���URL��IE�E�B���h�E����IE�I�u�W�F�N�g���擾����
' //
' //#���� strURL
' //
' //#�ߒl �w���URL��IE�I�u�W�F�N�g
' //
' //#��� ���łɊJ���Ă���E�B���h�E�Q�̒�����w���URL��IE�I�u�W�F�N�g���擾����B
' //
' //#���� 2012/06/06
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
' //#���� TestDOMGetFromWindowTitle
' //
' //#�T�v DOMGetFromWindowTitle�����{����
' //
' //#����
' //
' //#�ߒl �Ȃ�
' //
' //#��� DOMGetFromWindowTitle�����{���Ď擾����IE�I�u�W�F�N�g����URL���f�o�b�N�E�B���h�E�ɕ\������
' //
' //#���� 2012/06/06
' //      Coded by YASUTADA OOBA
' //
' /////////////////////////////////////////////////////////////////////////////////////
Sub TestDOMGetFromWindowTitle()
    
    Dim objDOM As Object
    
    Set objDOM = DOMGetFromWindowTitle("Yahoo")
    Debug.Print objDOM.Document.Title
    
End Sub
' /////////////////////////////////////////////////////////////////////////////////////
' //#���� DOMGetFromWindowTitle
' //
' //#�T�v �w��̃E�B���h�E�^�C�g����IE�I�u�W�F�N�g���擾����
' //
' //#���� strTitle
' //
' //#�ߒl IE�I�u�W�F�N�g
' //
' //#��� �S�ẴE�B���h�E�̂Ȃ�����w���URL���܂ރE�B���h�E���擾���AIE�I�u�W�F�N�g�����擾����B
' // �@�@ �ŏ��Ɍ����������̂�Ԃ��B
' //
' //#���� 2012/06/06
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
' //#���� DOMGetDocObjectFromNumber
' //
' //#�T�v ���[�g����m�[�h�ԍ����J���}��؂�Ŏw�肵�ē���ꂽ�h�L�������g�I�u�W�F�N�g��Ԃ�
' //
' //#���� objIE IE�I�u�W�F�N�g
' //      strDomTree    �J���}��؂�̃m�[�h�ԍ�
' //
' //#�ߒl �h�L�������g�I�u�W�F�N�g
' //
' //#���
' //
' //#���� 2012/06/06
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
' //#���� DOMSleepWhileBusy
' //
' //#�T�v �w���IE�I�u�W�F�N�g���r�W�[�̊ԑ҂�
' //
' //#���� objIE IE�I�u�W�F�N�g
' //      lngFirstWait �ŏ��ɑ҂���(Click�Ȃ�Busy��z�肵�Ă��Ȃ����̂�Busy�ɂȂ�O��
' //                   ���ɑJ�ڂ���ꍇ������B
' //
' //#�ߒl �Ȃ�
' //
' //#���
' //
' //#���� 2012/06/06
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
        MsgBox ("�^�C���A�E�g")
        DOMSleepWhileBusy = False
    End If

End Function

