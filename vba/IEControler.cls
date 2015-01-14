VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IEControler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
 
'IEControler
'Copyright (c) 2014 nezuq
'This software is released under the MIT License.
'http://opensource.org/licenses/mit-license.php
 
'[参照設定]
'Microsoft HTML Object Library
'Microsoft Internet Controls
 
'[定数]
Private Const TYPE_VERIFY_VALUE As String = "value"
Private Const TYPE_VERIFY_TITLE As String = "title"
Private Const TYPE_VERIFY_TEXT As String = "text"
Private Const TYPE_OPTIONSELECT_LABEL As String = "label"
Private Const TYPE_OPTIONSELECT_VALUE As String = "value"
Private Const TYPE_OPTIONSELECT_INDEX As String = "index"
Private Const WAITTIME_BEFOREPRINT As Integer = 2
 
'[変数]
Private IEs() As InternetExplorer
Private IdxIEs As Integer
Private ie As InternetExplorer
Private Doc As HTMLDocument
 
'[COM関数]
#If VBA7 Then
'64bit版
Private Declare PtrSafe Sub keybd_event Lib "user32" ( _
    ByVal bVk As Byte, _
    ByVal bScan As Byte, _
    ByVal dwFlags As Long, _
    ByVal dwExtraInfo As Long _
        )
Private Declare PtrSafe Function ShowWindow Lib "user32" ( _
    ByVal hwindow As Long, _
    ByVal cmdshow As Long _
        ) As Long
#Else
'32bit版
Private Declare Sub keybd_event Lib "user32" ( _
    ByVal bVk As Byte, _
    ByVal bScan As Byte, _
    ByVal dwFlags As Long, _
    ByVal dwExtraInfo As Long _
        )
Private Declare Function ShowWindow Lib "user32" ( _
    ByVal hwindow As Long, _
    ByVal cmdshow As Long _
        ) As Long
#End If

'[共通関数]
Private Function exp2ary(ByVal exp As String) As Variant
    Dim prms As Variant
    prms = Split(exp, "=")
    prms(0) = StrConv(Trim(prms(0)), vbLowerCase)
    exp2ary = prms
End Function
 
Private Sub sleep(ByVal sec As Integer)
    Dim ymdTo As Date
    ymdTo = DateAdd("s", sec, Now)
    While Now < ymdTo
       DoEvents
    Wend
End Sub
 
Private Sub clearClipBoard()
    Range("A1").Copy
    Excel.Application.CutCopyMode = False
End Sub
 
Private Sub waitIE(ByRef ie As InternetExplorer)
    Do While ie.Busy = True Or ie.readyState <> 4
        DoEvents
    Loop
End Sub
 
Private Function getElementById(ByRef objHtml As Object, ByVal id As String) As Variant
    Set getElementById = objHtml.getElementById(id)
End Function
 
Private Function getElementsByName(ByRef objHtml As Object, ByVal nmeObj As String) As Variant
    Set getElementsByName = objHtml.getElementsByName(nmeObj)
End Function
 
Private Function getElementsByClassName(ByRef objHtml As Object, ByVal nmeClass As String) As Variant
    Set getElementsByClassName = objHtml.getElementsByClassName(nmeClass)
End Function
 
Private Function getElementsByTagName(ByRef objHtml As Object, ByVal nmeTag As String) As Variant
    Set getElementsByTagName = objHtml.getElementsByTagName(nmeTag)
End Function

Private Function getElementByDomExp(ByRef objHtml As Object, ByRef exp As String) As Variant
    Set getElementByDomExp = Nothing
    
    Dim prms As Variant
    prms = Split(Replace(exp, " ", ""), ",")
    Dim idxPrms As Integer
    idxPrms = 0
    Dim obj As Object
    Set obj = objHtml
    Dim key, val As String
    Do While (idxPrms <= UBound(prms))
        key = prms(idxPrms)
        val = Right(key, Len(key) - 1)
        Select Case Left(key, 1)
            Case "#"
                Set obj = getElementById(obj, val)
                idxPrms = idxPrms + 1
            Case "."
                Set obj = getElementsByClassName(obj, val)(CInt(prms(idxPrms + 1)))
                idxPrms = idxPrms + 2
            Case Else
                Set obj = getElementsByTagName(obj, key)(CInt(prms(idxPrms + 1)))
                idxPrms = idxPrms + 2
        End Select
    Loop
    Set getElementByDomExp = obj
End Function

Private Function verify(ByVal obj As Object, ByVal exp As String) As Boolean
    Dim prms As Variant
    prms = exp2ary(exp)
    Dim isVerify As Boolean
    isVerify = False
    Select Case prms(0)
        Case TYPE_VERIFY_VALUE
            isVerify = (obj.value = prms(1))
        Case TYPE_VERIFY_TEXT
            isVerify = (obj.Text = prms(1))
        Case TYPE_VERIFY_TITLE
            isVerify = (obj.Title = prms(1))
    End Select
End Function

Private Function selectOption(ByVal exp As String, ByRef objSelect As HTMLSelectElement) As HTMLOptionElement
    Dim prms As Variant
    prms = exp2ary(exp)
    Dim typ As String
    typ = prms(0)
    Dim key As String
    key = prms(1)
    Select Case typ
        Case TYPE_OPTIONSELECT_VALUE
            objSelect.Options(key).Click
        Case Else
            Dim opt As Object
            For Each opt In objSelect.Options
                If ((typ = TYPE_OPTIONSELECT_LABEL And opt.Label = key) Or (typ = TYPE_OPTIONSELECT_VALUE And opt.value = key)) Then
                    opt.Click
                    Exit Function
                End If
            Next
    End Select
End Function

Public Sub pasteScreenShot(ByVal nmeCell As String, ByVal nmeSht As String, Optional ByVal nmeWbk As String = "")
    If (nmeWbk = "") Then nmeWbk = ThisWorkbook.name
    
    sleep WAITTIME_BEFOREPRINT
    Dim wbk As Workbook
    For Each wbk In Workbooks
        If (wbk.name = nmeWbk) Then
            Exit For
        End If
    Next
    Set wbk = Workbooks(nmeWbk)
    Dim sht As Worksheet
    Set sht = wbk.Worksheets(nmeSht)
    Dim rng As Range
    Set rng = sht.Range(nmeCell)
    sht.Paste Destination:=rng
End Sub

'[イベント]
Private Sub Class_Initialize()
    IdxIEs = -1
End Sub
 
'[公開関数]
Public Sub OpenIE()
    Set ie = New InternetExplorer
    ie.Visible = True
    IdxIEs = IdxIEs + 1
    ReDim Preserve IEs(IdxIEs)
    Set IEs(IdxIEs) = ie
End Sub

Public Sub CloseIE()
    ie.Quit
    Set ie = Nothing
    IdxIEs = IdxIEs - 1
    If (-1 < IdxIEs) Then
        ReDim Preserve IEs(IdxIEs)
        Set ie = IEs(IdxIEs)
    End If
End Sub

Public Sub ShowWindowNormalSize()
    Dim ret As Variant
    ret = ShowWindow(CLng(ie.Hwnd), 1)
End Sub

Public Sub ShowWindowMinSize()
    Dim ret As Variant
    ret = ShowWindow(CLng(ie.Hwnd), 2)
End Sub

Public Sub ShowWindowMaxSize()
    Dim ret As Variant
    ret = ShowWindow(CLng(ie.Hwnd), 3)
End Sub

Public Sub Wait(ByVal sec As Integer)
    sleep sec
End Sub

Public Sub Transfer(ByVal url As String)
    ie.Navigate url
    waitIE ie
    Set Doc = ie.document
End Sub

Public Sub SelectWindow(ByVal index As Integer)
    Set ie = IEs(index)
    Set Doc = ie.document
End Sub

Public Sub SelectFrame(ByVal index As Integer)
    Set Doc = ie.document.Window.frames(index).document
End Sub
 
Public Sub FocusElementById(ByVal id As String)
    getElementById(Doc, id).Focus
End Sub

Public Sub FocusElementByName(ByVal name As String, Optional ByVal index As Integer = 0)
    getElementsByName(Doc, name)(index).Focus
End Sub

Public Sub FocusElementByTagName(ByVal tagName As String, Optional ByVal index As Integer = 0)
    getElementsByTagName(Doc, tagName)(index).Focus
End Sub

Public Sub FocusElementByClassName(ByVal className As String, Optional ByVal index As Integer = 0)
    getElementsByClassName(Doc, className)(index).Focus
End Sub

Public Sub FocusElementByDomExp(ByVal exp As String)
    getElementByDomExp(Doc, exp).Focus
End Sub

Public Sub InputElementById(ByVal value As String, ByVal id As String)
    Dim elm As HTMLInputElement
    Set elm = getElementById(Doc, id)
    elm.Focus
    elm.value = value
End Sub

Public Sub InputElementByName(ByVal value As String, ByVal name As String, Optional ByVal index As Integer = 0)
    Dim elm As HTMLInputElement
    Set elm = getElementsByName(Doc, name)(index)
    elm.Focus
    elm.value = value
End Sub

Public Sub InputElementByTagName(ByVal value As String, ByVal tagName As String, Optional ByVal index As Integer = 0)
    Dim elm As HTMLInputElement
    Set elm = getElementsByTagName(Doc, tagName)(index)
    elm.Focus
    elm.value = value
End Sub

Public Sub InputElementByClassName(ByVal value As String, ByVal className As String, Optional ByVal index As Integer = 0)
    Dim elm As HTMLInputElement
    Set elm = getElementsByClassName(Doc, className)(index)
    elm.Focus
    elm.value = value
End Sub

Public Sub InputElementByDomExp(ByVal value As String, ByVal exp As String)
    Dim elm As HTMLInputElement
    Set elm = getElementByDomExp(Doc, exp)
    elm.Focus
    elm.value = value
End Sub

Public Sub SelectElementById(ByVal exp As String, ByVal id As String)
    Dim slct As HTMLSelectElement
    Set slct = Doc.getElementById(id)
    selectOption exp, slct
    waitIE ie
End Sub

Public Sub SelectElementByName(ByVal exp As String, ByVal name As String, Optional ByVal index As Integer = 0)
    Dim slct As HTMLSelectElement
    Set slct = Doc.getElementByName(name)(index)
    selectOption exp, slct
    waitIE ie
End Sub

Public Sub SelectElementByTagName(ByVal exp As String, ByVal tagName As String, Optional ByVal index As Integer = 0)
    Dim slct As HTMLSelectElement
    Set slct = Doc.getElementByTagName(tagName)(index)
    selectOption exp, slct
    waitIE ie
End Sub

Public Sub SelectElementByClassName(ByVal exp As String, ByVal className As String, Optional ByVal index As Integer = 0)
    Dim slct As HTMLSelectElement
    Set slct = Doc.getElementByClassName(className)(index)
    selectOption exp, slct
    waitIE ie
End Sub

Public Sub SelectElementByDomExp(ByVal optExp As String, ByVal domExp As String)
    Dim slct As HTMLSelectElement
    Set slct = Doc.getElementByDomExp(domExp)
    selectOption optExp, slct
    waitIE ie
End Sub

Public Sub ClickElementById(ByVal id As String)
    getElementById(Doc, id).Click
    waitIE ie
End Sub

Public Sub ClickElementByName(ByVal name As String, Optional ByVal index As Integer = 0)
    getElementsByName(Doc, name)(index).Click
    waitIE ie
End Sub

Public Sub ClickElementByTagName(ByVal tagName As String, Optional ByVal index As Integer = 0)
    getElementsByTagName(Doc, tagName)(index).Click
    waitIE ie
End Sub

Public Sub ClickElementByClassName(ByVal className As String, Optional ByVal index As Integer = 0)
    getElementsByClassName(Doc, className)(index).Click
    waitIE ie
End Sub

Public Sub ClickElementByDomExp(ByVal exp As String)
    getElementByDomExp(Doc, exp).Click
    waitIE ie
End Sub

Public Sub PushKeys(ByVal keyword As String, Optional ByVal cntRepeat As Integer = 1)
    For i = 1 To cntRepeat
        SendKeys keyword, True
    Next
End Sub

Public Sub PushEnterKeys(Optional ByVal cntRepeat As Integer = 1)
    For i = 1 To cntRepeat
        SendKeys "{enter}", True
    Next
End Sub

Public Sub PushTabKeys(Optional ByVal cntRepeat As Integer = 1)
    For i = 1 To cntRepeat
        SendKeys "{tab}", True
    Next
End Sub

Public Function VerifyAttributeById(ByVal exp As String, ByVal id As String) As Boolean
    VerifyAttributeById = True
    Dim elm As Object
    Set elm = getElementById(Doc, id)
    If (verify(elm, exp)) Then
        VerifyAttributeById = False
        Err.Raise 9999, "VerifyAttributeById", exp
    End If
End Function

Public Function VerifyAttributeByIdNoStop(ByVal exp As String, ByVal id As String) As Boolean
    VerifyAttributeByIdNoStop = True
    Dim elm As Object
    Set elm = getElementById(Doc, id)
    If (verify(elm, exp)) Then
        VerifyAttributeByIdNoStop = False
    End If
End Function

Public Function VerifyAttributeByName(ByVal exp As String, ByVal name As String, Optional ByVal index As Integer = 0) As Boolean
    VerifyAttributeByName = True
    Dim elm As Object
    Set elm = getElementsByName(Doc, name)(index)
    If (verify(elm, exp)) Then
        VerifyAttributeByName = False
        Err.Raise 9999, "VerifyAttributeByName", exp
    End If
End Function

Public Function VerifyAttributeByNameNoStop(ByVal exp As String, ByVal name As String, Optional ByVal index As Integer = 0) As Boolean
    VerifyAttributeByNameNoStop = True
    Dim elm As Object
    Set elm = getElementsByName(Doc, name)(index)
    If (verify(elm, exp)) Then
        VerifyAttributeByNameNoStop = False
    End If
End Function

Public Function VerifyAttributeByTagName(ByVal exp As String, ByVal tagName As String, Optional ByVal index As Integer = 0) As Boolean
    VerifyAttributeByTagName = True
    Dim elm As Object
    Set elm = getElementsByTagName(Doc, tagName)(index)
    If (verify(elm, exp)) Then
        VerifyAttributeByTagName = False
        Err.Raise 9999, "VerifyAttributeByTagName", exp
    End If
End Function

Public Function VerifyAttributeByTagNameNoStop(ByVal exp As String, ByVal tagName As String, Optional ByVal index As Integer = 0) As Boolean
    VerifyAttributeByTagNameNoStop = True
    Dim elm As Object
    Set elm = getElementsByTagName(Doc, tagName)(index)
    If (verify(elm, exp)) Then
        VerifyAttributeByTagNameNoStop = False
    End If
End Function

Public Function VerifyAttributeByClassName(ByVal exp As String, ByVal className As String, Optional ByVal index As Integer = 0) As Boolean
    VerifyAttributeByClassName = True
    Dim elm As Object
    Set elm = getElementsByClassName(Doc, className)(index)
    If (verify(elm, exp)) Then
        VerifyAttributeByClassName = False
        Err.Raise 9999, "VerifyAttributeByClassName", exp
    End If
End Function

Public Function VerifyAttributeByClassNameNoStop(ByVal exp As String, ByVal className As String, Optional ByVal index As Integer = 0) As Boolean
    VerifyAttributeByClassNameNoStop = True
    Dim elm As Object
    Set elm = getElementsByClassName(Doc, className)(index)
    If (verify(elm, exp)) Then
        VerifyAttributeByClassNameNoStop = False
    End If
End Function

Public Function VerifyAttributeByDomExp(ByVal exp As String, ByVal exp As String) As Boolean
    VerifyAttributeByDomExp = True
    Dim elm As Object
    Set elm = getElementByDomExp(Doc, exp)
    If (verify(elm, exp)) Then
        VerifyAttributeByDomExp = False
        Err.Raise 9999, "VerifyAttributeByDomExp", exp
    End If
End Function

Public Function VerifyAttributeByDomExpNoStop(ByVal exp As String, ByVal exp As String) As Boolean
    VerifyAttributeByDomExpNoStop = True
    Dim elm As Object
    Set elm = getElementByDomExp(Doc, exp)
    If (verify(elm, exp)) Then
        VerifyAttributeByDomExpNoStop = False
    End If
End Function

Public Sub ScreenShot(ByVal nmeCell As String, ByVal nmeSht As String, Optional ByVal nmeWbk As String = "")
    'TODO:if scrollbar exist then scroll
    clearClipBoard
    keybd_event vbKeySnapshot, 0&, &H1, 0&
    keybd_event vbKeySnapshot, 0&, &H1 Or &H2, 0&
    pasteScreenShot nmeCell, nmeSht, nmeWbk
End Sub
 
Public Sub ActiveScreenShot(ByVal nmeCell As String, ByVal nmeSht As String, Optional ByVal nmeWbk As String = "")
    'TODO:if scrollbar exist then scroll
    clearClipBoard
    keybd_event &HA4, 0&, &H1, 0&
    keybd_event vbKeySnapshot, 0&, &H1, 0&
    keybd_event vbKeySnapshot, 0&, &H1 Or &H2, 0&
    keybd_event &HA4, 0&, &H1 Or &H2, 0&
    pasteScreenShot nmeCell, nmeSht, nmeWbk
End Sub

Public Function Eval(ByVal nmeCommand As String, Optional ByVal prm1 As String = Empty, Optional ByVal prm2 As String = Empty, Optional ByVal prm3 As String = Empty) As Boolean
    Eval = True

    Dim idxMaxParam As Integer
    idxMaxParam = -1
    If (prm3 <> Empty) Then
        idxMaxParam = 2
    ElseIf (prm2 <> Empty) Then
        idxMaxParam = 1
    ElseIf (prm1 <> Empty) Then
        idxMaxParam = 0
    End If
    
    Select Case nmeCommand
        Case "OpenIE"
            OpenIE
        Case "CloseIE"
            CloseIE
        Case "ShowWindowNormalSize"
            ShowWindowNormalSize
        Case "ShowWindowMinSize"
            ShowWindowMinSize
        Case "ShowWindowMaxSize"
            ShowWindowMaxSize
        Case "Wait"
            Wait prm1
        Case "Transfer"
            Transfer prm1
        Case "SelectWindow"
            SelectWindow prm1
        Case "SelectFrame"
            SelectFrame prm1
        Case "FocusElementById"
            FocusElementById prm1
        Case "FocusElementByName"
            Select Case idxMaxParam
                Case Is > 0
                    FocusElementByName prm1, prm2
                Case Else
                    FocusElementByName prm1
            End Select
        Case "FocusElementByTagName"
            Select Case idxMaxParam
                Case Is > 0
                    FocusElementByTagName prm1, prm2
                Case Else
                    FocusElementByTagName prm1
            End Select
        Case "FocusElementByClassName"
            Select Case idxMaxParam
                Case Is > 0
                    FocusElementByClassName prm1, prm2
                Case Else
                    FocusElementByClassName prm1
            End Select
        Case "FocusElementByDomExp"
            FocusElementByDomExp prm1
        Case "InputElementById"
            InputElementById prm1, prm2
        Case "InputElementByName"
            Select Case idxMaxParam
                Case Is > 1
                    InputElementByName prm1, prm2, prm3
                Case Else
                    InputElementByName prm1, prm2
            End Select
        Case "InputElementByTagName"
            Select Case idxMaxParam
                Case Is > 1
                    InputElementByTagName prm1, prm2, prm3
                Case Else
                    InputElementByTagName prm1, prm2
            End Select
        Case "InputElementByClassName"
            Select Case idxMaxParam
                Case Is > 1
                    InputElementByClassName prm1, prm2, prm3
                Case Else
                    InputElementByClassName prm1, prm2
            End Select
        Case "InputElementByDomExp"
            InputElementByDomExp prm1, prm2
        Case "SelectElementById"
            SelectElementById prm1, prm2
        Case "SelectElementByName"
            Select Case idxMaxParam
                Case Is > 1
                    SelectElementByName prm1, prm2, prm3
                Case Else
                    SelectElementByName prm1, prm2
            End Select
        Case "SelectElementByTagName"
            Select Case idxMaxParam
                Case Is > 1
                    SelectElementByTagName prm1, prm2, prm3
                Case Else
                    SelectElementByTagName prm1, prm2
            End Select
        Case "SelectElementByClassName"
            Select Case idxMaxParam
                Case Is > 1
                    SelectElementByClassName prm1, prm2, prm3
                Case Else
                    SelectElementByClassName prm1, prm2
            End Select
        Case "SelectElementByDomExp"
            SelectElementByDomExp prm1, prm2
        Case "PushKeys"
            Select Case idxMaxParam
                Case Is > 0
                    PushKeys prm1, prm2
                Case Else
                    PushKeys prm1
            End Select
        Case "PushEnterKeys"
            Select Case idxMaxParam
                Case Is > -1
                    PushEnterKeys prm1
                Case Else
                    PushEnterKeys
            End Select
        Case "PushTabKeys"
            Select Case idxMaxParam
                Case Is > -1
                    PushTabKeys prm1
                Case Else
                    PushTabKeys
            End Select
        Case "ClickElementById"
            ClickElementById prm1
        Case "ClickElementByName"
            Select Case idxMaxParam
                Case Is > 0
                    ClickElementByName prm1, prm2
                Case Else
                    ClickElementByName prm1
            End Select
        Case "ClickElementByTagName"
            Select Case idxMaxParam
                Case Is > 0
                    ClickElementByTagName prm1, prm2
                Case Else
                    ClickElementByTagName prm1
            End Select
        Case "ClickElementByClassName"
            Select Case idxMaxParam
                Case Is > 0
                    ClickElementByClassName prm1, prm2
                Case Else
                    ClickElementByClassName prm1
            End Select
        Case "ClickElementByDomExp"
            ClickElementByDomExp prm1
        Case "VerifyAttributeById"
            Eval = VerifyAttributeById(prm1, prm2)
        Case "VerifyAttributeByName"
            Select Case idxMaxParam
                Case Is > 1
                    Eval = VerifyAttributeByName(prm1, prm2, prm3)
                Case Else
                    Eval = VerifyAttributeByName(prm1, prm2)
            End Select
        Case "VerifyAttributeByTagName"
            Select Case idxMaxParam
                Case Is > 1
                    Eval = VerifyAttributeByTagName(prm1, prm2, prm3)
                Case Else
                    Eval = VerifyAttributeByTagName(prm1, prm2)
            End Select
        Case "VerifyAttributeByClassName"
            Select Case idxMaxParam
                Case Is > 1
                    Eval = VerifyAttributeByClassName(prm1, prm2, prm3)
                Case Else
                    Eval = VerifyAttributeByClassName(prm1, prm2)
            End Select
        Case "VerifyAttributeByDomExp"
            Eval = VerifyAttributeByDomExp(prm1, prm2)
        Case "VerifyAttributeByIdNoStop"
            Eval = VerifyAttributeByIdNoStop(prm1, prm2)
        Case "VerifyAttributeByNameNoStop"
            Select Case idxMaxParam
                Case Is > 1
                    Eval = VerifyAttributeByNameNoStop(prm1, prm2, prm3)
                Case Else
                    Eval = VerifyAttributeByNameNoStop(prm1, prm2)
            End Select
        Case "VerifyAttributeByTagNameNoStop"
            Select Case idxMaxParam
                Case Is > 1
                    Eval = VerifyAttributeByTagNameNoStop(prm1, prm2, prm3)
                Case Else
                    Eval = VerifyAttributeByTagNameNoStop(prm1, prm2)
            End Select
        Case "VerifyAttributeByClassNameNoStop"
            Select Case idxMaxParam
                Case Is > 1
                    Eval = VerifyAttributeByClassNameNoStop(prm1, prm2, prm3)
                Case Else
                    Eval = VerifyAttributeByClassNameNoStop(prm1, prm2)
            End Select
        Case "VerifyAttributeByDomExpNoStop"
            Eval = VerifyAttributeByDomExpNoStop(prm1, prm2)
        Case "ScreenShot"
            Select Case idxMaxParam
                Case Is > 1
                    ScreenShot prm1, prm2, prm3
                Case Else
                    ScreenShot prm1, prm2
            End Select
        Case "ActiveScreenShot"
            Select Case idxMaxParam
                Case Is > 1
                    ActiveScreenShot prm1, prm2, prm3
                Case Else
                    ActiveScreenShot prm1, prm2
            End Select
        Case Else
            MsgBox nmeCommand & " is not found.", vbCritical, "Error"
    End Select
End Function


