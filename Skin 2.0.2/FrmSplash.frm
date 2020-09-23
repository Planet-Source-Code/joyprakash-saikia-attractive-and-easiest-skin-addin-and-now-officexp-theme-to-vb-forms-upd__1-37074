VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   " Marry chirstmas"
   ClientHeight    =   2115
   ClientLeft      =   1920
   ClientTop       =   3135
   ClientWidth     =   7980
   Icon            =   "FrmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   7980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      BackColor       =   &H00000000&
      Caption         =   "Do you Want To see Fade Effect"
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   2700
      TabIndex        =   1
      Top             =   1620
      Width           =   2955
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   750
      Top             =   1680
   End
   Begin VB.PictureBox picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   38.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2010
      Left            =   360
      ScaleHeight     =   2010
      ScaleWidth      =   7590
      TabIndex        =   0
      Top             =   240
      Width           =   7590
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ---------------- My Little Text Animation I made for the Eve of
'                       Christmas
'------------------  joy  24th December 2001
' Now I've added this on this Submisssion on 29th July 2002

'
Private intCount As Long
Private MaxColor As Long
Private Fadeout As Boolean
Private Str1$, Str2$

Dim Timeout As Long
Dim B As Boolean

Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long


Public Function ChangeShape() As Boolean
On Error GoTo SHAPEERR

Dim CreateRgnVal As Long
Dim lngWidth As Long, lngHeight As Long
'This form has Scale Mode as Twips ,
' But All WINDOWS API Recognize Pixels
' So let us Convert it
lngHeight = ScaleY(Height, vbTwips, vbPixels)
lngWidth = ScaleX(Width, vbTwips, vbPixels)
'Now Create the Region with retrieve parameters
lngRetVal1 = CreateEllipticRgn(0, 0, lngWidth, lngHeight)
' Put that region to this window and Redraw it
 SetWindowRgn Me.hwnd, lngRetVal1, True
' Delete the Object
DeleteObject lngRetVal1

ChangeShape = True
Exit Function
SHAPEERR:
ChangeShape = False
End Function



Private Sub Check1_Click()
    If Check1.Value = 0 Then Picture1.Cls
End Sub

Private Sub Form_DblClick()
 Unload Me
 Set frmSplash = Nothing
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
        'Set the MAx Color to 255 i.e, RGB(255,255,255)
        MaxColor = 255
        
        Str1$ = App.EXEName & " " & App.Major & "." & App.Minor & "." & App.Revision
        Fadeout = True
        intCount = 255
        B = ChangeShape
        Picture1.AutoRedraw = True
        Check1.Value = 1
       
End Sub



Private Sub Timer1_Timer()
    If Timeout = 8 Then
            ' EXECUTE THE TIMER FOR 4 ROUNDS
            Timer1.Enabled = False
            Unload Me
            Set frmSplash = Nothing
            Exit Sub
    End If
If Fadeout = False Then
            
            Picture1.ForeColor = RGB(intCount + 1, intCount + 1, intCount + 1)
            Picture1.CurrentX = intCount
            Picture1.CurrentY = intCount
            Picture1.Print Str1$
            
            intCount = intCount + 10
     If intCount >= MaxColor Or (Picture1.CurrentX >= Picture1.ScaleWidth Or _
                    Picture1.CurrentY >= Picture1.ScaleHeight) Then
            Fadeout = True
            Timeout = Timeout + 1
     End If
Else
    If Check1.Value = 0 Then Picture1.Cls
  If intCount > 0 Then
            
            Picture1.ForeColor = RGB(intCount + 1, intCount + 1, 0)
            Picture1.CurrentX = intCount
            Picture1.CurrentY = intCount
            Picture1.Print Str1$
           
            intCount = intCount - 10
            If intCount <= 100 Or (Picture1.CurrentX >= Picture1.ScaleWidth Or _
                    Picture1.CurrentY >= Picture1.ScaleHeight) Then
                    Fadeout = False
                    Timeout = Timeout + 1
            End If
    Else
               Fadeout = False: intCount = 0
  End If
End If
End Sub
