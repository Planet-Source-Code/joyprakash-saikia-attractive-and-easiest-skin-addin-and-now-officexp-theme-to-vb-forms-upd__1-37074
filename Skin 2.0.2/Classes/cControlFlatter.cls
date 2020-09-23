VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cControlFlater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Module     :   ClsControlFlater
'Author     :   JoyPrakash Saikia
'Created    :   29/07/2002
'Purpose    :
' History :   Based on VB ACCelerator timer Class SSubTmr6.dll
' I've Used this Class from PSC it Self to Put Skin And Office Xp Style Together


Option Explicit
Implements ISubclass

Private Enum EDrawStyle
    FC_DRAWNORMAL = 0
    FC_DRAWRAISED = 1
    FC_DRAWPRESSED = 2
    FC_DRAWDISABLED = 3
End Enum
Private Enum ECmdType
    CT_GENERAL = 0
    CT_COMBOBOX = 1
    CT_COMMANDBUTTON = 2
    CT_SCROLLBAR = 3
End Enum

Private Const WM_COMMAND = &H111
Private Const WM_PAINT = &HF
Private Const WM_SETFOCUS = &H7
Private Const WM_KILLFOCUS = &H8
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_TIMER = &H113
Private Const WM_ENABLE = &HA

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Const SM_CXHSCROLL = 21
Private Const SM_CXHTHUMB = 10
Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Const PS_SOLID = 0
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Const GWL_STYLE = (-16)
Private Const CBS_DROPDOWN = &H2&
Private Const CBS_DROPDOWNLIST = &H3&
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, lpsz2 As Any) As Long
Private Const CBN_CLOSEUP = 8
Private Const CB_GETDROPPEDSTATE = &H157
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Const GW_CHILD = 5
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function CopyRect Lib "user32" (lpDestRect As RECT, lpSourceRect As RECT) As Long

Private m_hWnd As Long
Private m_hWndEdit As Long
Private m_hWndParent As Long
Private m_bSubclass As Boolean
Private m_bMouseOver As Boolean
Private m_bMouseDown As Boolean
Private m_bFocus As Boolean
Private m_bDisabled  As Boolean
Private m_cType As ECmdType

Private m_TrackMouseMove As Boolean

Public Sub Attach(ByRef objthis As Object)
    Dim lhWnd As Long
    Dim lStyle As Long
    pRelease
    
    On Error Resume Next
    lhWnd = objthis.hwnd
    If (Err.Number <> 0) Then
        Err.Raise vbObjectError + 1048 + 1, App.EXEName & "", "InCompatible control type passed to 'Attach' - it must be a control with a hWnd property."
        Exit Sub
    End If
    
    Select Case TypeName(objthis)
    Case "CommandButton"
        m_cType = CT_COMMANDBUTTON
    
    Case "ComboBox"
        m_cType = CT_COMBOBOX
        m_hWndParent = GetParent(lhWnd)
    
    Case "ImageCombo"
        m_cType = CT_COMBOBOX
        m_hWndParent = lhWnd
        lhWnd = FindWindowEx(lhWnd, 0&, "ComboBox", ByVal 0&)
    
    
    Case "HScrollBar"
        m_cType = CT_SCROLLBAR
    Case "OptionButton"
        m_cType = CT_GENERAL
        m_hWndParent = GetParent(lhWnd)
    Case "CheckBox"
        m_cType = CT_GENERAL
        m_hWndParent = GetParent(lhWnd)
    Case Else
        lStyle = GetWindowLong(lhWnd, -20) ' GWL_STYLE)
        If ((lStyle And CBS_DROPDOWN) = CBS_DROPDOWN) Or ((lStyle And CBS_DROPDOWNLIST) = CBS_DROPDOWNLIST) Then
            m_cType = CT_COMBOBOX
            m_hWndParent = lhWnd
        Else
            m_cType = CT_GENERAL
            With objthis
                .Move .Left + 2 * Screen.TwipsPerPixelX, .Top + 2 * Screen.TwipsPerPixelY, .Width - 4 * Screen.TwipsPerPixelX, .Height - 4 * Screen.TwipsPerPixelY
            End With
            m_hWndParent = GetParent(lhWnd)
        End If
    End Select
    
    pAttach lhWnd
End Sub

Private Sub pAttach(ByRef hWndA As Long)
    Dim lStyle As Long
    
    m_hWnd = hWndA
    If (m_hWnd <> 0) Then
        lStyle = GetWindowLong(m_hWnd, GWL_STYLE)
        If (lStyle And CBS_DROPDOWN) = CBS_DROPDOWN Then
            m_hWndEdit = GetWindow(m_hWnd, GW_CHILD)
        End If
        
        AttachMessage Me, m_hWnd, WM_PAINT
        AttachMessage Me, m_hWnd, WM_SETFOCUS
        AttachMessage Me, m_hWnd, WM_KILLFOCUS
        If TrackMouseMove = True Then
            ' Attach Only if MoveMove is Tracked
         AttachMessage Me, m_hWnd, WM_MOUSEMOVE
        End If
        AttachMessage Me, m_hWnd, WM_LBUTTONDOWN
        AttachMessage Me, m_hWnd, WM_LBUTTONUP
        AttachMessage Me, m_hWnd, WM_TIMER
        AttachMessage Me, m_hWnd, WM_ENABLE
        If (m_hWndEdit <> 0) Then
            AttachMessage Me, m_hWndEdit, WM_SETFOCUS
            AttachMessage Me, m_hWndEdit, WM_KILLFOCUS
            AttachMessage Me, m_hWndEdit, WM_MOUSEMOVE
        End If
        If m_cType = CT_COMBOBOX Then
            AttachMessage Me, m_hWndParent, WM_COMMAND
        End If
        m_bSubclass = True
    End If
End Sub

Private Sub pRelease()
    If m_bSubclass Then
        DetachMessage Me, m_hWnd, WM_PAINT
        DetachMessage Me, m_hWnd, WM_SETFOCUS
        DetachMessage Me, m_hWnd, WM_KILLFOCUS
        DetachMessage Me, m_hWnd, WM_MOUSEMOVE
        DetachMessage Me, m_hWnd, WM_LBUTTONDOWN
        DetachMessage Me, m_hWnd, WM_LBUTTONUP
        DetachMessage Me, m_hWnd, WM_TIMER
        DetachMessage Me, m_hWnd, WM_ENABLE
        If (m_hWndEdit <> 0) Then
            DetachMessage Me, m_hWndEdit, WM_SETFOCUS
            DetachMessage Me, m_hWndEdit, WM_KILLFOCUS
            DetachMessage Me, m_hWndEdit, WM_MOUSEMOVE
        End If
        If m_cType = CT_COMBOBOX Then
            DetachMessage Me, m_hWndParent, WM_COMMAND
        End If
    End If
    m_hWnd = 0
    m_hWndEdit = 0
    m_hWndParent = 0
End Sub

Private Sub OnTimer(ByVal bCheckMouse As Boolean)
    Dim bOver As Boolean
   
    Dim PT As POINTAPI
    Dim lhWnd As Long
    
    If bCheckMouse Then
        bOver = True
        GetCursorPos PT
        lhWnd = WindowFromPoint(PT.X, PT.Y)
        If lhWnd <> m_hWnd And lhWnd <> m_hWndEdit Then
            bOver = False
        End If
    End If
    
    If Not bOver Then
        KillTimer m_hWnd, 1
        m_bMouseOver = False
        DrawMe
    End If
End Sub

Private Sub Class_Terminate()
    pRelease
End Sub

Private Property Let ISubclass_MsgResponse(ByVal RHS As SSubTimer6.EMsgResponse)
    '
End Property

Private Property Get ISubclass_MsgResponse() As SSubTimer6.EMsgResponse
    If (CurrentMessage = WM_TIMER) Then
        ISubclass_MsgResponse = emrPostProcess
    Else
        ISubclass_MsgResponse = emrPreprocess
    End If
End Property

Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case iMsg
    Case WM_COMMAND
        If (m_hWnd = lParam) Then
            Select Case wParam \ &H10000
            Case CBN_CLOSEUP
                DrawMe
            End Select
        End If
    Case WM_PAINT
        DrawMe
    Case WM_ENABLE
        m_bDisabled = (IsWindowEnabled(m_hWnd) = 0)
        DrawMe
    Case WM_SETFOCUS
        m_bFocus = True
        DrawMe
    Case WM_KILLFOCUS
        m_bFocus = False
        DrawMe
    Case WM_MOUSEMOVE
        If Not m_bMouseOver Then
            m_bMouseOver = True
            DrawMe
            SetTimer m_hWnd, 1, 10, 0
        End If
    Case WM_LBUTTONDOWN
        m_bMouseDown = True
        DrawMe
    Case WM_LBUTTONUP
        m_bMouseDown = False
        DrawMe
    Case WM_TIMER
        OnTimer True
    End Select
End Function

Private Sub DrawMe()
    Dim dwStyle As EDrawStyle
    Select Case m_cType
    Case CT_GENERAL
        If m_bDisabled Then
            dwStyle = FC_DRAWDISABLED
        ElseIf m_bFocus Or m_bMouseOver Then
            dwStyle = FC_DRAWRAISED
        Else
            dwStyle = FC_DRAWNORMAL
        End If
        DrawEdit dwStyle
    Case CT_COMBOBOX
        If m_bDisabled Then
            dwStyle = FC_DRAWDISABLED
        ElseIf SendMessageLong(m_hWnd, CB_GETDROPPEDSTATE, 0, 0) <> 0 Then
            dwStyle = FC_DRAWPRESSED
        ElseIf m_bFocus Or m_bMouseOver Then
            dwStyle = FC_DRAWRAISED
        Else
            dwStyle = FC_DRAWNORMAL
        End If
        DrawCombo dwStyle
    Case CT_COMMANDBUTTON
        If m_bDisabled Then
            dwStyle = FC_DRAWDISABLED
        ElseIf m_bMouseDown Then
            dwStyle = FC_DRAWPRESSED
        ElseIf m_bMouseOver Then
            dwStyle = FC_DRAWRAISED
        Else
            dwStyle = FC_DRAWNORMAL
        End If
        DrawCommand dwStyle
    Case CT_SCROLLBAR
        If m_bFocus Or m_bMouseOver Then
            DrawScrollBar FC_DRAWRAISED
        Else
            DrawScrollBar FC_DRAWNORMAL
        End If
    End Select
End Sub

Private Sub DrawCommand(ByVal dwStyle As EDrawStyle)
    Dim rcItem As RECT
    Dim pDC As Long
    
    GetClientRect m_hWnd, rcItem
    pDC = GetDC(m_hWnd)
    
    Select Case dwStyle
    Case FC_DRAWNORMAL
        Draw3DRect pDC, rcItem, vb3DLight, vbButtonFace
    Case FC_DRAWDISABLED
        Draw3DRect pDC, rcItem, vb3DDKShadow, vbButtonFace
    Case FC_DRAWRAISED
        Draw3DRect pDC, rcItem, vb3DShadow, vbButtonShadow
    Case FC_DRAWPRESSED
        Draw3DRect pDC, rcItem, vbButtonShadow, vb3DHighlight
    End Select
    
    InflateRect rcItem, -1, -1
    Draw3DRect pDC, rcItem, vbButtonFace, vbButtonFace
    InflateRect rcItem, -1, -1
    Draw3DRect pDC, rcItem, vbButtonFace, vbButtonFace
    
    DeleteDC pDC
End Sub

Private Sub DrawEdit(ByVal dwStyle As EDrawStyle)
    Dim rcItem As RECT
    Dim pDC As Long
    Dim PT As POINTAPI
    
    GetWindowRect m_hWnd, rcItem
    PT.X = rcItem.Left
    PT.Y = rcItem.Top
    ScreenToClient m_hWndParent, PT
    rcItem.Left = PT.X
    rcItem.Top = PT.Y
    PT.X = rcItem.Right
    PT.Y = rcItem.Bottom
    ScreenToClient m_hWndParent, PT
    rcItem.Right = PT.X
    rcItem.Bottom = PT.Y
    
    pDC = GetDC(m_hWndParent)
    
    Select Case dwStyle
    Case FC_DRAWDISABLED
        InflateRect rcItem, 1, 1
        Draw3DRect pDC, rcItem, vb3DHighlight, vb3DHighlight
    Case FC_DRAWNORMAL
        Draw3DRect pDC, rcItem, vbButtonShadow, vb3DHighlight
        InflateRect rcItem, 2, 2
        Draw3DRect pDC, rcItem, vb3DFace, vbButtonFace
    Case FC_DRAWRAISED
        InflateRect rcItem, 2, 2
        Draw3DRect pDC, rcItem, vb3DHighlight, vb3DDKShadow
    End Select
    
    DeleteDC pDC
End Sub

Private Sub DrawCombo(ByVal dwStyle As EDrawStyle)
    Dim rcItem As RECT
    Dim pDC As Long
    
    GetClientRect m_hWnd, rcItem
    pDC = GetDC(m_hWnd)
    
    Select Case dwStyle
    Case FC_DRAWDISABLED
        Draw3DRect pDC, rcItem, vbButtonFace, vbButtonFace
        InflateRect rcItem, -1, -1
        Draw3DRect pDC, rcItem, vb3DHighlight, vb3DHighlight
    Case FC_DRAWNORMAL
        Draw3DRect pDC, rcItem, vbButtonFace, vbButtonFace
        InflateRect rcItem, -1, -1
        Draw3DRect pDC, rcItem, vbButtonFace, vbButtonFace
    Case Else
        Draw3DRect pDC, rcItem, vbButtonShadow, vb3DHighlight
        InflateRect rcItem, -1, -1
        Draw3DRect pDC, rcItem, vbButtonFace, vbButtonFace
    End Select
    
    InflateRect rcItem, -1, -1
   
    rcItem.Left = rcItem.Right - GetSystemMetrics(SM_CXHTHUMB)
    Draw3DRect pDC, rcItem, vbButtonFace, vbButtonFace
    InflateRect rcItem, -1, -1
    Draw3DRect pDC, rcItem, vbButtonFace, vbButtonFace
    
    Select Case dwStyle
    Case FC_DRAWNORMAL
        rcItem.Top = rcItem.Top - 1
        rcItem.Bottom = rcItem.Bottom + 1
        ' Draw3DRect pDC, rcItem, vb3DHighlight, vb3DHighlight
        rcItem.Left = rcItem.Left - 1
        rcItem.Right = rcItem.Left
        Draw3DRect pDC, rcItem, vbWindowBackground, &H0
    Case FC_DRAWRAISED
        rcItem.Top = rcItem.Top - 1
        rcItem.Bottom = rcItem.Bottom + 1
        rcItem.Right = rcItem.Right + 1
        Draw3DRect pDC, rcItem, vb3DShadow, vbButtonShadow
    Case FC_DRAWPRESSED
        rcItem.Left = rcItem.Left - 1
        rcItem.Top = rcItem.Top - 2
        Draw3DRect pDC, rcItem, vbButtonShadow, vb3DHighlight
        OffsetRect rcItem, 1, 1
        Draw3DRect pDC, rcItem, vbButtonShadow, vb3DHighlight
    End Select
    DeleteDC pDC
End Sub

Private Sub DrawScrollBar(ByVal dwStyle As EDrawStyle)
    Dim rcItem As RECT, rcItemA As RECT, rcItemB As RECT
    Dim pDC As Long
    Dim lButtonSize As Long
    
    GetClientRect m_hWnd, rcItem
    pDC = GetDC(m_hWnd)
    
    Select Case dwStyle
    Case FC_DRAWNORMAL
        Draw3DRect pDC, rcItem, vbButtonFace, vbButtonFace
        InflateRect rcItem, -1, -1
        Draw3DRect pDC, rcItem, vb3DHighlight, vb3DHighlight
        InflateRect rcItem, -1, -1
    Case FC_DRAWRAISED
        Draw3DRect pDC, rcItem, vbButtonShadow, vb3DHighlight
        InflateRect rcItem, -1, -1
        Draw3DRect pDC, rcItem, vbButtonFace, vbButtonFace
    End Select
    
    CopyRect rcItemA, rcItem
    CopyRect rcItemB, rcItem
    lButtonSize = GetSystemMetrics(SM_CXHSCROLL)
        
    Select Case dwStyle
    Case FC_DRAWNORMAL
        rcItemA.Right = lButtonSize
        rcItemB.Left = rcItemB.Right - lButtonSize + 2

        Draw3DRect pDC, rcItemA, vbButtonFace, vbButtonFace
        InflateRect rcItemA, -1, -1
        Draw3DRect pDC, rcItemA, vbButtonFace, vbButtonFace

        Draw3DRect pDC, rcItemB, vbButtonFace, vbButtonFace
        InflateRect rcItemB, -1, -1
        Draw3DRect pDC, rcItemB, vbButtonFace, vbButtonFace
    Case FC_DRAWRAISED
        rcItemA.Right = lButtonSize
        rcItemB.Left = rcItemB.Right - lButtonSize + 1

        Draw3DRect pDC, rcItemA, vbButtonFace, vbButtonFace
        InflateRect rcItemA, -1, -1
        Draw3DRect pDC, rcItemA, vb3DHighlight, vbButtonShadow

        Draw3DRect pDC, rcItemB, vbButtonFace, vbButtonFace
        InflateRect rcItemB, -1, -1
        Draw3DRect pDC, rcItemB, vb3DHighlight, vbButtonShadow
    Case FC_DRAWPRESSED
    
    End Select
    
    DeleteDC pDC
End Sub

Private Function Draw3DRect( _
      ByVal hDC As Long, _
      ByRef rcItem As RECT, _
      ByVal oTopLeftColor As OLE_COLOR, _
      ByVal oBottomRightColor As OLE_COLOR)
    Dim hPen As Long
    Dim hPenOld As Long
    Dim tP As POINTAPI
    
    hPen = CreatePen(PS_SOLID, 1, TranslateColor(oTopLeftColor))
    hPenOld = SelectObject(hDC, hPen)
    MoveToEx hDC, rcItem.Left, rcItem.Bottom - 1, tP
    LineTo hDC, rcItem.Left, rcItem.Top
    LineTo hDC, rcItem.Right - 1, rcItem.Top
    SelectObject hDC, hPenOld
    DeleteObject hPen
    If (rcItem.Left <> rcItem.Right) Then
        hPen = CreatePen(PS_SOLID, 1, TranslateColor(oBottomRightColor))
        hPenOld = SelectObject(hDC, hPen)
        LineTo hDC, rcItem.Right - 1, rcItem.Bottom - 1
        LineTo hDC, rcItem.Left, rcItem.Bottom - 1
        SelectObject hDC, hPenOld
        DeleteObject hPen
    End If
End Function

Private Function TranslateColor(ByVal clr As OLE_COLOR, Optional hPal As Long = 0) As Long
    If OleTranslateColor(clr, hPal, TranslateColor) Then
        TranslateColor = -1
    End If
End Function


Public Property Let TrackMouseMove(data As Boolean)
        m_TrackMouseMove = data
        If data = False And m_TrackMouseMove = True Then
            If m_bSubclass Then
                DetachMessage Me, m_hWnd, WM_MOUSEMOVE
                If (m_hWndEdit <> 0) Then
                    DetachMessage Me, m_hWndEdit, WM_MOUSEMOVE
                End If
            End If
        End If
End Property
Public Property Get TrackMouseMove() As Boolean
 TrackMouseMove = m_TrackMouseMove
End Property
