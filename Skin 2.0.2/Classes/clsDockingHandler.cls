VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDockingHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'---------------------------------------------------------------------------------------
'Copyright  :   Free
'Module     :   clsDockingHandler
'Author     :   UnKown
'Created    :   19/07/2002
'Purpose    :   I found this module Somewhere on the Net
'               I found it very Useful , to make Dockable form
'                Just try it.
'---------------------------------------------------------------------------------------
Option Explicit
'
' A class that's an abstract subclass of Form,
' that you can inherit other forms from.
'
'
' Usage:
'
' 1. Include an instance of the class in your form
' 2. Set the the ParentForm property when the form is loaded
' 3. Call StartDockDrag when you want to begin snap-dragging
' 4. Call UpdateDockDrag as the user continues to move the cursor.
'
' The most obvious use is to call the StartDockDrag when
' there's a MouseDown event, and then Call UpdateDockDrag
' on each MouseMove (if the left mouse button is pressed).
' Look at the example form - you may find other ideas!
'
' The class works with Twips, so all methods/properties
' should be passed in that standard. I know, it sucks,
' but VB feels the most comfortable with it...
' Your form can use any scale mode, as long as you
' convert the data before using the class methods!
' Externally set properties


Private mParentForm As Form
Private mAlwaysOnTop As Boolean
Private mSnapDistance As Long

' Internally set variables
Private mStartDragX As Single
Private mStartDragY As Single

Private mWorkAreaRect As RECT

' Are we attached to some edge of the screen?
Private mAttachedToRight As Boolean
Private mAttachedToLeft As Boolean
Private mAttachedToTop As Boolean
Private mAttachedToBottom As Boolean

Private mWindowStyle As Long

Public Sub StartDockDrag(ByVal X As Single, ByVal Y As Single)

    ' Get the WorkArea - the area of the desktop not taken
    ' by the taskbar, using a little known but interesting
    ' API call
    SystemParametersInfo SPI_GETWORKAREA, 0&, mWorkAreaRect, 0&
    
    ' Convert from pixels to twips
    mWorkAreaRect.Top = mWorkAreaRect.Top * Screen.TwipsPerPixelY
    mWorkAreaRect.Left = mWorkAreaRect.Left * Screen.TwipsPerPixelX
    mWorkAreaRect.Bottom = mWorkAreaRect.Bottom * Screen.TwipsPerPixelY
    mWorkAreaRect.Right = mWorkAreaRect.Right * Screen.TwipsPerPixelX
    
    mStartDragX = X
    mStartDragY = Y

End Sub

Public Sub UpdateDockDrag(ByVal X As Single, ByVal Y As Single)
Dim DiffX As Long, DiffY As Long
Dim NewX As Long, NewY As Long
Dim ToLeftDistance As Long
Dim ToRightDistance As Long
Dim ToTopDistance As Long
Dim ToBottomDistance As Long

    If ParentForm Is Nothing Then Exit Sub

    DiffX = X - mStartDragX
    DiffY = Y - mStartDragY
    
    If DiffX = 0 And DiffY = 0 Then Exit Sub
    
    NewX = mParentForm.Left + DiffX
    NewY = mParentForm.Top + DiffY
    
    ' Find the distance to the screen edges
    ToRightDistance = mWorkAreaRect.Right - (NewX + mParentForm.Width)
    ToLeftDistance = NewX - mWorkAreaRect.Left
    ToBottomDistance = mWorkAreaRect.Bottom - (NewY + mParentForm.Height)
    ToTopDistance = NewY - mWorkAreaRect.Top
    
    ' The idea in all the following code is the same:
    ' If wer'e not already attached some specific edge,
    ' find out if we should.
    ' If wer'e already attached, find out whether we should
    ' "break" the attachment, or stay put.
    If Not mAttachedToBottom Then
        If Abs(ToBottomDistance) <= mSnapDistance Then
            ' Attach to edge
            NewY = mWorkAreaRect.Bottom - mParentForm.Height
            mAttachedToBottom = True
        End If
    Else
        If Abs(ToBottomDistance) > mSnapDistance Then
            ' Break the attachement
            mAttachedToBottom = False
        Else
            ' Stay at current position
            NewY = mParentForm.Top
        End If
    End If
    
    If Not mAttachedToTop Then
        If Abs(ToTopDistance) <= mSnapDistance Then
            NewY = mWorkAreaRect.Top
            mAttachedToTop = True
        End If
    Else
        If Abs(ToTopDistance) > mSnapDistance Then
            mAttachedToTop = False
        Else
            NewY = mParentForm.Top
        End If
    End If

    If Not mAttachedToRight Then
        If Abs(ToRightDistance) <= mSnapDistance Then
            NewX = mWorkAreaRect.Right - mParentForm.Width
            mAttachedToRight = True
        End If
    Else
        If Abs(ToRightDistance) > mSnapDistance Then
            mAttachedToRight = False
        Else
            NewX = mParentForm.Left
        End If
    End If

    If Not mAttachedToLeft Then
        If Abs(ToLeftDistance) <= mSnapDistance Then
            NewX = mWorkAreaRect.Left
            mAttachedToLeft = True
        End If
    Else
        If Abs(ToLeftDistance) > mSnapDistance Then
            mAttachedToLeft = False
        Else
            NewX = mParentForm.Left
        End If
    End If
    
    ' Position the window, converting to pixels again
    SetWindowPos mParentForm.hwnd, mWindowStyle, _
        NewX / Screen.TwipsPerPixelX, _
        NewY / Screen.TwipsPerPixelY, _
        mParentForm.Width / Screen.TwipsPerPixelX, _
        mParentForm.Height / Screen.TwipsPerPixelY, 0

End Sub

Public Property Set ParentForm(vData As Form)
    Set mParentForm = vData
End Property

Public Property Get ParentForm() As Form
    Set ParentForm = mParentForm
End Property

' This is just a small bonus, since you may
' want your form to be a sticky one.
Public Property Let AlwaysOnTop(vData As Boolean)
    
    mAlwaysOnTop = vData

    ' Change the window style and update the window
    If mAlwaysOnTop Then
        mWindowStyle = HWND_TOPMOST
    Else
        mWindowStyle = HWND_NOTOPMOST
    End If
    
    If Not ParentForm Is Nothing Then
        SetWindowPos mParentForm.hwnd, mWindowStyle, _
            mParentForm.Left / Screen.TwipsPerPixelX, _
            mParentForm.Top / Screen.TwipsPerPixelY, _
            mParentForm.Width / Screen.TwipsPerPixelX, _
            mParentForm.Height / Screen.TwipsPerPixelY, 0
    End If

End Property

Public Property Get AlwaysOnTop() As Boolean
    AlwaysOnTop = mAlwaysOnTop
End Property

Public Property Let SnapDistance(vData As Long)
    mSnapDistance = vData
End Property

Public Property Get SnapDistance() As Long
    SnapDistance = mSnapDistance
End Property

Private Sub Class_Initialize()
    
    ' Default window style is normal window (not always on top)
    mWindowStyle = HWND_NOTOPMOST
    
    ' Default snap distance is 10 pixels
    mSnapDistance = 10 * Screen.TwipsPerPixelX

End Sub
