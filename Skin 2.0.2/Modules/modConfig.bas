Attribute VB_Name = "modConfig"
'---------------------------------------------------------------------------------------
'Copyright  :   JoyPrakash Saikia 2002
'Module     :   modConfig
'Author     :   JoyPrakash Saikia
'Created    :   15/06/2002
'Purpose    :  Configuration Module
'
'---------------------------------------------------------------------------------------



Option Explicit
Public Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As Any, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

'Remove System maximize Menu
Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
'Constant Used for Menu
Const MF_BYPOSITION = &H400&
Const MF_REMOVE = &H1000&
'Constants used for Changing the Form Style Or Removing Form Border
Const GWL_STYLE = (-16)
Public Const GWL_EXSTYLE = (-20)
 Const WS_BORDER = &H800000
 Const WS_DLGFRAME = &H400000
 Const WS_CAPTION = &HC00000
 Const WS_THICKFRAME = &H40000

Private Const WS_SIZEBOX = WS_THICKFRAME
' API for Painting
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, _
    ByVal X As Long, ByVal y As Long, _
    ByVal nWidth As Long, ByVal nHeight As Long, _
    ByVal hSrcDC As Long, _
    ByVal xSrc As Long, ByVal ySrc As Long, _
    ByVal dwRop As Long) As Long

' VBA contains a "GetObject" Method
' To remove the Ambiguity , It is being Changed to APIGetObject
Declare Function APIGetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Declare Function GetPrivateProfileString Lib "kernel32" _
    Alias "GetPrivateProfileStringA" (ByVal lpApplicationName _
    As String, lpKeyName As Any, ByVal lpDefault As String, _
    ByVal lpRetunedString As String, ByVal nSize As Long, _
    ByVal lpFilename As String) As Long

Declare Function WritePrivateProfileString Lib "kernel32" _
    Alias "WritePrivateProfileStringA" (ByVal lpApplicationName _
    As String, ByVal lpKeyName As Any, ByVal lpString As Any, _
    ByVal lplFileName As String) As Long
' Current program preferences, especially not skin-specific
Type CurrentPreferncesType
    SkinName As String
    SkinsPath As String
    SkinFullPath As String
    ' Add more as you like here
End Type

' Skin-specific preferences.
'
Type SkinPreferencesType
    BackColor As Long
    ExitButtonX As Long
    ExitButtonY As Long
    MinButtonX As Long
    MinButtonY As Long
    FormHeight As Long
    FormWidth As Long
    ' Add more as you like here
End Type

Public CurrPrefs As CurrentPreferncesType

Public SkinPrefs As SkinPreferencesType
Public m_SkinPathName As String             'added on 12th July for New Property SkinPathName
Private SkinConfigFile As String

Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Public Const IMAGE_BITMAP = 0
Public Const LR_DEFAULTCOLOR = &H0
Public Const LR_LOADMAP3DCOLORS = &H1000
Public Const LR_LOADFROMFILE = &H10
Public Const LR_LOADTRANSPARENT = &H20




Public Const SRCCOPY = &HCC0020
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2

'
' Misc. declarations
'
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long

Public Const NULLHANDLE = 0

Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SPI_GETWORKAREA = 48

Public Sub RemoveParentBorder(lhWnd As Long)
    'Hide the System MenuBar
    Dim lStyle As Long
    lStyle = GetWindowLong(lhWnd, GWL_STYLE)
    lStyle = lStyle And Not (WS_BORDER Or WS_DLGFRAME Or WS_CAPTION Or WS_BORDER Or WS_SIZEBOX Or WS_THICKFRAME)
    SetWindowLong lhWnd, GWL_STYLE, lStyle
    FrameChanged lhWnd
End Sub
Public Sub FrameChanged(hwnd As Long)
    SetWindowPos hwnd, 0, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOZORDER Or SWP_FRAMECHANGED Or SWP_NOSIZE Or SWP_NOMOVE
End Sub

Public Sub RemoveMaximize(Frm As Form)
 Dim hMenu As Long, intCnt As Long

 hMenu = GetSystemMenu(Frm.hwnd, False)

    If hMenu Then
        ' Get System menu's menu count
        intCnt = GetMenuItemCount(hMenu)
        If intCnt Then
            RemoveMenu hMenu, intCnt - 3, MF_BYPOSITION Or MF_REMOVE 'Remove the Maximized
            DrawMenuBar Frm.hwnd
        End If
    End If
End Sub

' Read skin-specific preferences from skin.ini file
Public Sub ReadSkinPreferences()
    Dim strColor As String
    Dim varColorArr As Variant
    
    CurrPrefs.SkinFullPath = SkinPathName + CurrPrefs.SkinName + "\"
    SkinConfigFile = CurrPrefs.SkinFullPath + "skin.cfg"
    
    If Dir(SkinConfigFile) = "" Then
        Err.Raise 1, , "Unable to Locate " & SkinConfigFile
    End If
    
    With SkinPrefs
            
        strColor = ReadSettings("Skin", "backcolor", SkinConfigFile)
        varColorArr = Split(strColor, ",")
        .BackColor = RGB(varColorArr(0), varColorArr(1), varColorArr(2))
        .ExitButtonX = ReadSettings("Skin", "ExitButtonX", SkinConfigFile)
        .ExitButtonY = ReadSettings("Skin", "ExitButtonY", SkinConfigFile)
        .MinButtonX = ReadSettings("Skin", "MinButtonX", SkinConfigFile)
        .MinButtonY = ReadSettings("Skin", "MinButtonY", SkinConfigFile)
        
    End With
    
End Sub

Public Function ReadSettings(Section As String, KeyName As String, Filename As String) As String
Dim Str As String
    
    Str = String(255, Chr(0))
    ReadSettings = VBA.Left(Str, GetPrivateProfileString(Section, ByVal KeyName, 0, Str, Len(Str), Filename))

End Function

Public Function WriteSettings(Section As String, KeyName As String, KeyValue As String, Filename As String) As Boolean
Dim Ret As Long
    
    Ret = WritePrivateProfileString(Section, KeyName, KeyValue, Filename)
    If Ret = 0 Then
        WriteSettings = True
    Else
        WriteSettings = False
    End If
    
End Function
'Following Property is Declared Only for the Module
Public Property Get SkinPathName() As String
If Trim$(m_SkinPathName) = "" Then 'if Skin Path is Not Supplied
    SkinPathName = App.Path + "\skins\"
Else
    SkinPathName = m_SkinPathName
End If
End Property
