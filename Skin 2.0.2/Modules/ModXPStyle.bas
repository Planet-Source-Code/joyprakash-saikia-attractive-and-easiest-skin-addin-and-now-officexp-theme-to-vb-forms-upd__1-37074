Attribute VB_Name = "ModXPStyle"
Option Explicit
Option Compare Text
Dim i  As Integer
Public K() As cControlFlater
'-------------------Flat Border Const
Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_WINDOWEDGE = &H100
Public Const WS_EX_CLIENTEDGE = &H200
Public Const WS_EX_STATICEDGE = &H20000






Sub FlatAllControls(frm As Form, bolMouseMove As Boolean)
  Dim CTL As Control
  Dim tmpCnt As Integer
  If i = 0 Then
    ReDim Preserve K(0 To frm.Controls.Count)
  Else
    tmpCnt = i + frm.Controls.Count
    ReDim Preserve K(tmpCnt)
  End If
  For Each CTL In frm.Controls '.Count - 1
   
          Select Case TypeName(CTL)
            Case "CommandButton", "TextBox", "ComboBox", "ImageCombo", "HScrollBar", "ListBox", "VScrollBar", "CheckBox", "OptionButton"
              Set K(i) = New cControlFlater
              K(i).TrackMouseMove = bolMouseMove
              K(i).Attach CTL
             
              i = i + 1
             
          End Select
      Next CTL
End Sub
Sub ReleaseMouseMove(frm As Form)

  Dim CTL As Control
  Dim tmpCnt As Integer
    If UBound(K) > 0 Then
    For tmpCnt = 0 To frm.Controls.Count - 1
       
          Select Case TypeName(frm.Controls(tmpCnt))
            Case "CommandButton", "TextBox", "ComboBox", "ImageCombo", "HScrollBar", "ListBox", "VScrollBar", "CheckBox", "OptionButton"
              K(tmpCnt).TrackMouseMove = False
              
              i = tmpCnt
          End Select
      Next tmpCnt
    End If
End Sub
 
 

Sub main()
frmSplash.Show 1

FrmSkinSample.Show
End Sub





