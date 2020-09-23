VERSION 5.00
Begin VB.Form FrmSkinSample 
   ClientHeight    =   4695
   ClientLeft      =   2100
   ClientTop       =   2370
   ClientWidth     =   8280
   Icon            =   "frmSample.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4695
   ScaleWidth      =   8280
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   4395
      Left            =   240
      ScaleHeight     =   4395
      ScaleWidth      =   8055
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   90
      Width           =   8055
      Begin VB.VScrollBar VScroll1 
         Height          =   4455
         Left            =   7800
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   0
         Width           =   255
      End
      Begin VB.CheckBox chkTrackMouseMove 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Track Mouse Move"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   150
         TabIndex        =   5
         Top             =   2760
         Width           =   2295
      End
      Begin VB.CheckBox chkApplyThemes 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Apply OfficeXP Theme"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   150
         TabIndex        =   4
         Top             =   2400
         Width           =   2295
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   120
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   4440
         Width           =   6735
      End
      Begin VB.ComboBox cboTransitionType 
         Height          =   315
         ItemData        =   "frmSample.frx":000C
         Left            =   4440
         List            =   "frmSample.frx":0040
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2040
         Width           =   2175
      End
      Begin VB.CheckBox chkEnable 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Always Apply Skin"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   150
         TabIndex        =   3
         Top             =   2040
         Width           =   2295
      End
      Begin VB.TextBox TxtTransitionDelay 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   5640
         TabIndex        =   6
         Top             =   1290
         Width           =   975
      End
      Begin VB.CheckBox chkResizable 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Make This Form Resizable"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   150
         TabIndex        =   2
         Top             =   1650
         Width           =   2295
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&Exit"
         Height          =   375
         Left            =   4800
         TabIndex        =   8
         Top             =   3360
         Width           =   975
      End
      Begin VB.CheckBox chkAnimation 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Show Animation On Load"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   150
         TabIndex        =   1
         Top             =   1260
         Width           =   2295
      End
      Begin prjSKIN.SkinCtl SkinCtl1 
         Height          =   375
         Left            =   2850
         TabIndex        =   0
         Top             =   3330
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   661
         TransitionType  =   393221
         TransitionDelay =   500
         AllowResize     =   0   'False
         ApplyAnimation  =   -1  'True
         Enabled         =   -1  'True
         SkinPathName    =   ""
         ApplyThemes     =   0   'False
         TrackMouseMove  =   0   'False
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Now VB GUI Made Eazy"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5520
         TabIndex        =   16
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label lblTransitionType 
         BackStyle       =   0  'Transparent
         Caption         =   "Transition Type"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3120
         TabIndex        =   14
         Top             =   2040
         Width           =   1410
      End
      Begin VB.Label lblVote 
         BackStyle       =   0  'Transparent
         Caption         =   "Please Vote Me"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   13
         Top             =   4020
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Startup Delay (in Milliseconds)"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3180
         TabIndex        =   12
         Top             =   1410
         Width           =   2370
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "SKIN CONTROL Version 2.0.2"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   11
         Top             =   0
         UseMnemonic     =   0   'False
         Width           =   3600
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Now The Skin Control Has Themes Like OfficeXP, Customize diffrent Controls with XP and Other Styles and 10 Skins"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   10
         Top             =   570
         Width           =   5310
      End
   End
End
Attribute VB_Name = "FrmSkinSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
'Copyright  :   JoyPrakash Saikia 2002
'Module     :   FrmSkinSample
'Author     :   JoyPrakash Saikia
'Created    :   19/07/2002
'Purpose    :   This is Sample Form for Skin Control
'               This is only Avalilable at PSC
'---------------------------------------------------------------------------------------
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Const CB_FINDSTRING = &H14C
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Const LinkPSC = "http://www.pscode.com/vb/scripts/BrowseCategoryOrSearchResults.asp?txtCriteria=Attractive+and+easiest+Skin+&blnWorldDropDownUsed=TRUE&txtMaxNumberOfEntriesPerPage=10&blnResetAllVariables=TRUE&lngWId=1&optSort=Alphabetical"

Private Sub cboTransitionType_Click()

  SaveSetting App.EXEName, "SampleSkin", "TransitionType", cboTransitionType.ItemData(cboTransitionType.ListIndex)
   
End Sub

Private Sub Check1_Click()
 SaveSetting App.EXEName, "SampleSkin", "ApplyThemes", CStr(chkAnimation.Value)
End Sub

Private Sub chkAnimation_Click()
    SaveSetting App.EXEName, "SampleSkin", "ApplyAnimation", CStr(chkAnimation.Value)
End Sub

Private Sub chkAnimation_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ShowHandCursor
    
  
End Sub

Private Sub chkApplyThemes_Click()
        SaveSetting App.EXEName, "SampleSkin", "ApplyThemes", CStr(chkApplyThemes.Value)
        
End Sub

Private Sub chkApplyThemes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ShowHandCursor
End Sub

Private Sub chkEnable_Click()
    SaveSetting App.EXEName, "SampleSkin", "Enabled", CStr(chkEnable.Value)
End Sub

Private Sub chkEnable_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ShowHandCursor

End Sub

Private Sub chkResizable_Click()
    SaveSetting App.EXEName, "SampleSkin", "AllowResize", CStr(chkResizable.Value)
    SkinCtl1.AllowResize = chkResizable.Value
End Sub



Private Sub chkResizable_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   ShowHandCursor

End Sub

Private Sub chkTrackMouseMove_Click()
 SaveSetting App.EXEName, "SampleSkin", "TrackMouseMove", CStr(chkTrackMouseMove.Value)
 SkinCtl1.TrackMouseMove = chkTrackMouseMove.Value
End Sub

Private Sub chkTrackMouseMove_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        ShowHandCursor
End Sub

Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
     Select Case KeyCode
      
        Case vbKeyEscape:
                        Unload Me
     End Select
     
End Sub

Private Sub Form_Load()
On Error GoTo LOAD_ERR
    KeyPreview = True
  Label3 = App.EXEName & " " & App.Major & "." & App.Minor & "." & App.Revision
    
If GetSetting(App.EXEName, "SampleSkin", "Enabled", "1") = "1" Then
        SkinCtl1.DefaultSkinName = GetSetting(App.EXEName, "SampleSkin", "DefaultSkinName", "FadeMac")
        chkEnable.Value = 1
    With SkinCtl1
        .TransitionType = GetSetting(App.EXEName, "SampleSkin", "TransitionType", Zoomout)
         
         'Show Appropriate Transition Type on ComboBox
         ShowTransitiontype
      
        .AllowResize = GetSetting(App.EXEName, "SampleSkin", "AllowResize", 1)
        'Added on 29th July
        .ApplyThemes = GetSetting(App.EXEName, "SampleSkin", "ApplyThemes", 1)
        .TrackMouseMove = GetSetting(App.EXEName, "SampleSkin", "TrackMouseMove", 1)
        .ApplyAnimation = GetSetting(App.EXEName, "SampleSkin", "ApplyAnimation", 1)
        
        TxtTransitionDelay.Text = GetSetting(App.EXEName, "SampleSkin", "TransitionDelay", 500)
        TxtTransitionDelay.Text = IIf(TxtTransitionDelay = "", 500, TxtTransitionDelay)
        .TransitionDelay = TxtTransitionDelay
        chkAnimation.Value = IIf(.ApplyAnimation, 1, 0)
        chkResizable.Value = IIf(.AllowResize, 1, 0)
        chkTrackMouseMove.Value = IIf(.TrackMouseMove, 1, 0)
        chkApplyThemes.Value = IIf(.ApplyThemes, 1, 0)
        
'************************* Following Statment is the can be the only Statment to Apply Skin
'                           NOW THEMES Also
        'if you donot want to use Skin
        Set .ClientObject = Picture1
'****************************************************************************************
    End With
Else
        SkinCtl1.Enabled = False
        chkResizable.Visible = False
        chkAnimation.Visible = False
        TxtTransitionDelay.Visible = False
        Label2.Visible = False
        chkEnable.Value = 0
        lblTransitionType.Visible = False
        cboTransitionType.Visible = False
        chkEnable.Caption = "Enable to Skin On Load"
        chkApplyThemes.Visible = False
        chkTrackMouseMove.Visible = False
End If
    Exit Sub
LOAD_ERR:
        MsgBox Err.Description, vbCritical, "Skin Sample Load Error"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Set FrmSkinSample = Nothing
End Sub

Private Sub lblVote_Click()
    ShellExecute 0, vbNullString, LinkPSC, vbNullString, vbNullString, vbMaximizedFocus
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.MousePointer = vbDefault
End Sub

Private Sub SkinCtl1_SkinSelected()
    SaveSetting App.EXEName, "SampleSkin", "DefaultSkinName", SkinCtl1.DefaultSkinName
End Sub

Private Sub TxtTransitionDelay_KeyPress(KeyAscii As Integer)
    If Not (IsNumeric(Chr$(KeyAscii)) Or KeyAscii = 8) Then KeyAscii = 0
End Sub

Private Sub TxtTransitionDelay_Validate(Cancel As Boolean)
        SaveSetting App.EXEName, "SampleSkin", "TransitionDelay", TxtTransitionDelay.Text
End Sub

Private Sub ShowTransitiontype()
Select Case SkinCtl1.TransitionType
    Case RightBottomTOLeftTop: cboTransitionType.ListIndex = 0
    Case Zoomout:              cboTransitionType.ListIndex = 1
    Case LeftTopToRightBottom: cboTransitionType.ListIndex = 2
    Case LeftBottomToRightTop: cboTransitionType.ListIndex = 3
    Case SlideFromLeft:        cboTransitionType.ListIndex = 4
    Case SlidEFromTop:         cboTransitionType.ListIndex = 5
End Select
End Sub
Sub ShowHandCursor()
        SetCursor LoadCursor(ByVal 0&, 32649&)
End Sub

