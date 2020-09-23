Attribute VB_Name = "AniModule"
'---------------------------------------------------------------------------------------
'Copyright  :   JoyPrakash Saikia 2002
'Module     :   AniModule
'Author     :   JoyPrakash Saikia
'Created    :   15/06/2002
'Purpose    :   TO Make AnimateWindow in Action
'---------------------------------------------------------------------------------------
Option Explicit



'/*
' *   windows 2000 ,Windows XP windows 98 2nd edition  and Windows ME  has an API function Called
'       AnimateWindow. But there is problem in VB FORMs, when you use this function for a form
' with Frames ,GRID etc. , then it leaves some black spots on it. This is very annoying situation
' So I have used subclassing to animate the windows without Flikering.
'
' -joyprakash Saikia
' */
Private mP_Currentform As Form
Public Enum ESetWindowPosStyles
    SWP_SHOWWINDOW = &H40
    SWP_HIDEWINDOW = &H80
    SWP_FRAMECHANGED = &H20
    SWP_NOACTIVATE = &H10
    SWP_NOCOPYBITS = &H100
    SWP_NOMOVE = &H2
    SWP_NOOWNERZORDER = &H200 ' Don't do owner Z ordering
    SWP_NOREDRAW = &H8
    SWP_NOREPOSITION = SWP_NOOWNERZORDER
    SWP_NOSIZE = &H1
    SWP_NOZORDER = &H4
    SWP_DRAWFRAME = SWP_FRAMECHANGED
End Enum

Public Const AW_HOR_POSITIVE = &H1
Public Const AW_HOR_NEGATIVE = &H2
Public Const AW_VER_POSITIVE = &H4
Public Const AW_VER_NEGATIVE = &H8
Public Const AW_CENTER = &H10
Public Const AW_HIDE = &H10000
Public Const AW_ACTIVATE = &H20000
Public Const AW_SLIDE = &H40000
Public Const AW_BLEND = &H80000
'property VAriable for TransitionType  for the SkinCTL



Public Declare Function AnimateWindow Lib "user32" _
    (ByVal hwnd As Long, _
    ByVal dwTime As Long, ByVal dwFlags As Long) As Long

Public Const WM_PRINTCLIENT = &H318
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (Destination As Any, Source As Any, ByVal Length As Long)

Public Declare Function GetWindowLong Lib "user32" Alias _
    "GetWindowLongA" (ByVal hwnd As Long, _
    ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias _
    "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long
Public Const GWL_WNDPROC = (-4)

Public Declare Function GetProp Lib "user32" Alias "GetPropA" _
    (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Function SetProp Lib "user32" Alias "SetPropA" _
    (ByVal hwnd As Long, ByVal lpString As String, _
    ByVal hData As Long) As Long
Public Declare Function RemoveProp Lib "user32" Alias "RemovePropA" _
    (ByVal hwnd As Long, ByVal lpString As String) As Long

Public Declare Function CallWindowProc Lib "user32" Alias _
    "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
    ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, _
    ByVal lParam As Long) As Long

Public Declare Function OleTranslateColor _
    Lib "oleaut32.dll" _
    (ByVal lOleColor As Long, _
    ByVal lHPalette As Long, _
    lColorRef As Long) As Long

Private Declare Function CreateSolidBrush Lib "gdi32" _
    (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, _
    lpRect As RECT, ByVal hBrush As Long) As Long
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type


 Sub PrintClient(ByVal hDC As Long, ByVal frm As Form, ByVal lParam As Long)

    Dim rct As RECT
    Dim hBr As Long
    'Fill in the hDC with the form's
    'background color. Otherwise the form
    'may appear Totally Garbage.
    rct.Left = 0
    rct.Top = 0
    rct.Right = frm.ScaleX(frm.ScaleWidth, frm.ScaleMode, vbPixels)
    rct.Bottom = frm.ScaleY(frm.ScaleHeight, frm.ScaleMode, vbPixels)
    hBr = CreateSolidBrush(TranslateColor(frm.BackColor))
    FillRect hDC, rct, hBr
    DeleteObject hBr

End Sub

Public Function TranslateColor(inCol As Long) As Long

    Dim retCol As Long
    OleTranslateColor inCol, 0&, retCol
    TranslateColor = retCol
End Function

Public Function AnimWndProc(ByVal hwnd As Long, ByVal wMsg As Long, _
ByVal wParam As Long, ByVal lParam As Long) As Long

    Dim lProc As Long
    Dim lPtr As Long
    Dim frm As Form

    lProc = GetProp(hwnd, "ExAnimWndProc")
    lPtr = GetProp(hwnd, "ExAnimWndPtr")

    'Catch the WM_PRINTCLIENT message so the form
    'won't look like garbage when it appears.
    If wMsg = WM_PRINTCLIENT Then
        CopyMemory frm, lPtr, 4
        PrintClient wParam, mP_Currentform, lParam
        CopyMemory frm, 0&, 4
    End If
    AnimWndProc = CallWindowProc(lProc, hwnd, wMsg, wParam, lParam)

End Function

Public Sub SubclassAnim(frm As Form)

    Dim l As Long

    If GetProp(frm.hwnd, "ExAnimWndProc") <> 0 Then
        'Already subclassed
        Exit Sub
    End If

    l = GetWindowLong(frm.hwnd, GWL_WNDPROC)
    SetProp frm.hwnd, "ExAnimWndProc", l
    SetProp frm.hwnd, "ExAnimWndPtr", ObjPtr(frm)

    SetWindowLong frm.hwnd, GWL_WNDPROC, AddressOf AnimWndProc

End Sub

Public Sub UnSubclassAnim(frm As Form)

    Dim l As Long

    l = GetProp(frm.hwnd, "ExAnimWndProc")
    If l = 0 Then
        'Isn't subclassed anyway
        Exit Sub
    End If

    SetWindowLong frm.hwnd, GWL_WNDPROC, l
    RemoveProp frm.hwnd, "ExAnimWndProc"
    RemoveProp frm.hwnd, "ExAnimWndPtr"

End Sub
'--end block--'

Public Sub AnimateOnLoad(CurrentFrm As Form, ByVal Transition As Long, delay As Long)
  If FindCorrectVersion = True Then
         Set mP_Currentform = CurrentFrm
         SubclassAnim CurrentFrm
        AniModule.AnimateWindow CurrentFrm.hwnd, delay, _
         Transition
        UnSubclassAnim CurrentFrm
         ' Added On 20th July For the Memory Leak
       Set mP_CurrentForm = Nothing
 End If
End Sub
Public Sub ActivateForm(frm As Form)
'Purpose    :   you Can use this Procedure If you Still See Some Part of the Form is not Refreshed
'               
	Dim cnt As Control
	For Each cnt In frm.Controls
		If Not (TypeOf cnt Is Frame) Then cnt.Refresh
		Next
	frm.Refresh
End Sub
Public Sub AnimateOnUnLoad(CurrentFrm As Form, delay As Long, Optional Fade As Boolean = False)
 
 If FindCorrectVersion = True Then
 Set mP_Currentform = CurrentFrm
     SubclassAnim CurrentFrm
     If Fade = True Then
       AnimateWindow CurrentFrm.hwnd, delay, _
        AW_BLEND Or &H10000
      Else
        AnimateWindow CurrentFrm.hwnd, delay, _
         AW_HOR_POSITIVE Or AW_VER_NEGATIVE Or AW_HIDE
      End If
        UnSubclassAnim CurrentFrm
         ' Added On 20th July For the Memory Leak
       Set mP_CurrentForm = Nothing
    End If
End Sub

Function FindCorrectVersion() As Boolean
'Used for Checking OS
Dim OSInfo As OSVERSIONINFO
Dim Ret As Long
OSInfo.dwOSVersionInfoSize = Len(OSInfo)
    'Get the Windows version
    Ret = GetVersionEx(OSInfo)
    If Ret = 0 Then FindCorrectVersion = False: Exit Function
    With OSInfo
    If .dwPlatformId = 1 And .dwBuildNumber >= 22 Then
        'windows 98 2nd Edition or more
      FindCorrectVersion = True
    ElseIf .dwPlatformId = 2 And .dwMajorVersion >= 5 Then
           'Windows 2000 or windowsXP
        FindCorrectVersion = True
    End If
 End With
End Function



