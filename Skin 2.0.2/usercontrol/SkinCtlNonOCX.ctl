VERSION 5.00
Begin VB.UserControl SkinCtl 
   BackStyle       =   0  'Transparent
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1845
   KeyPreview      =   -1  'True
   ScaleHeight     =   25
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   123
   Begin VB.ComboBox lstSkins 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   30
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   0
      Width           =   1785
   End
   Begin VB.Menu mnuSkinProp 
      Caption         =   "ShowProperties"
      Visible         =   0   'False
      Begin VB.Menu EdSkin 
         Caption         =   "Edit Skin Properties"
      End
   End
End
Attribute VB_Name = "SkinCtl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'Created By         :   Joyprakash Saikia
'Created on         :   15/6/2002

' Comments:         Just before Couple of days, One of my Friend Show
'                   a Concept of Toggleing the Skins at run time , I find it wonderful.
'                   That was my initial concept for Making this Descent Control.
' Please Vote Me.
'
Option Explicit

Public Enum TrasitionEnum
        'These are the Diffrent Animation TYpe
        RightBottomTOLeftTop = AW_HOR_NEGATIVE + AW_VER_NEGATIVE + AW_ACTIVATE
        Zoomout = AW_CENTER + AW_ACTIVATE
        LeftTopToRightBottom = AW_HOR_POSITIVE + AW_VER_POSITIVE + AW_ACTIVATE
        LeftBottomToRightTop = AW_HOR_POSITIVE + AW_VER_NEGATIVE + AW_ACTIVATE
        SlideFromLeft = AW_SLIDE + AW_HOR_POSITIVE + AW_ACTIVATE
        SlidEFromTop = AW_SLIDE + AW_HOR_POSITIVE + AW_VER_POSITIVE + AW_ACTIVATE
        
End Enum

'Event Declaration
Event SkinSelected()


Dim NumXSlices As Long
Dim NumYSlices As Long

' Minimum needed number of X slices so we don't mess-up
' the button positions
Dim MinXSlices As Long
                        
' Width/height of pad w/o any horizontal segment
Dim BaseXSize As Long
Dim BaseYSize As Long

' Used when resizing the window -
' X/Y distance of the mouse pointer from the form's edge
Dim XDistance As Long
Dim YDistance As Long

' Boolean flags - the current state of the form
Dim InXDrag As Boolean ' In horizontal resize
Dim InYDrag As Boolean ' In vertical resize
Dim InFormDrag As Boolean ' In window drag

Dim NoRedraw As Boolean

' Set to TRUE when in ListSkins(), to prevent lstSkins_Click()
' events from being handled while the list is created
Dim InListSkins As Boolean

' Size of right/bottom segments
Dim XEdgeSize As Single
Dim YEdgeSize As Single

' Handler for window dragging & docking
Dim DockHandler As New clsDockingHandler

' Holds the actual edge skin bitmaps
Dim EdgeImages(FE_LAST) As clsBitmap

' Holds the region data for each of the skin bitmaps
Dim EdgeRegions(FE_LAST) As RegionDataType

Dim WindowRegion As Long ' Current window region

' Custom Exit/Minimize buttons
Dim MyExitButton As New clsButton
Dim MinimizeButton As New clsButton

' Default size of client area. Used to compute the number of
' x/y segments needed when the program is loaded
Dim DEFAULT_CLIENT_HEIGHT As Long '= 500
Dim DEFAULT_CLIENT_WIDTH As Long  '= 500
Dim mvar_picClientArea As Object
Private WithEvents Form As Form
Attribute Form.VB_VarHelpID = -1
'for Animation Effect
Private m_AnimationApplied As Boolean
Private m_AllowResize As Boolean
'For Skin Setting
Private DefaultSkinSupplied As Boolean

Private m_Enabled As Boolean

 ' Added On the Non OCX Version on 12th July
Private m_Delay As Long

Private m_TransitionType As TrasitionEnum
'Added  for office XP like themes
Private m_Applythemes As Boolean
Private m_TrackMouseMove As Boolean
Public mbolAlreadyFlatten As Boolean

Private Sub Form_Load()
        ' if you Modify the Skin control's Enable Property
        ' through  Code , (Not with the Property window) then
        ' Load event will Execute ,
        ' But we don't nedd it if It is Disabled , Right!!!
        If m_Enabled = False Then Set Form = Nothing: Exit Sub
        
        ' Following is the Skin Loading Process
        Form.ScaleMode = 3
        ReadSkinPreferences '
        RemoveParentBorder Form.hwnd
       
       If m_Applythemes = True And mbolAlreadyFlatten = False Then FlatAllControls Form, m_TrackMouseMove
        Form.BackColor = SkinPrefs.BackColor 'to reduce Flikering
                If m_AnimationApplied = True Then
            AnimateOnLoad Form, TransitionType, TransitionDelay
        End If
             
        RemoveMaximize Form
        IntImageFileNames
        DEFAULT_CLIENT_HEIGHT = Form.ScaleHeight
        DEFAULT_CLIENT_WIDTH = Form.ScaleWidth
        LoadAllSkins
        Set DockHandler.ParentForm = Form
        ColorAdjust
        ListSkins


         
End Sub


' A mouse button press may initiate form dragging or resizing
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = vbLeftButton Then
    
        ' Test whether the user has pressed a "button",
        ' and show the 'down button' image if so
        If MyExitButton.HitTest(CLng(X), CLng(Y)) Then
            MyExitButton.PaintDownImage
            Exit Sub
        
        ElseIf MinimizeButton.HitTest(CLng(X), CLng(Y)) Then
            MinimizeButton.PaintDownImage
            Exit Sub
        End If
    
        YDistance = Y - Form.ScaleHeight
        XDistance = X - Form.ScaleWidth
        
        ' If the mouse pointer is on the the bottom edge,
        ' flag Y (vertical) drag
        If Abs(YDistance) < YEdgeSize Then
           If m_AllowResize Then InYDrag = True
        End If
        
        ' If the mouse pointer is on the the right edge,
        ' flag X drag. Don't start drag if wer'e in the window
        ' title area
        If Abs(XDistance) < XEdgeSize And _
           Y > EdgeImages(FE_TOP_RIGHT).Height Then
            If m_AllowResize Then InXDrag = True
        End If
        
        ' If we're in the window title area, start form draggin'
        If (Y <= EdgeImages(FE_TOP_H_SEGMENT).Height) Then
            DockHandler.StartDockDrag X * Screen.TwipsPerPixelX, _
                Y * Screen.TwipsPerPixelY
            InFormDrag = True
        End If
    
    End If

End Sub


Private Sub form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim NewYSlices As Single
Dim NewXSlices As Single
Dim ShowXResizeCursor As Boolean
Dim ShowYResizeCursor As Boolean
Dim ResizingNeeded As Boolean

    If InFormDrag Then
        ' Continue window draggin'
        DockHandler.UpdateDockDrag X * Screen.TwipsPerPixelX, _
            Y * Screen.TwipsPerPixelY
        Exit Sub
    End If
    
    ' Determine what kind of cursor should be shown
    
    If Abs(Y - Form.ScaleHeight) < YEdgeSize Or InYDrag Then
       If m_AllowResize Then ShowYResizeCursor = True
    End If
    
    If (Abs(X - Form.ScaleWidth) < XEdgeSize And _
        Y > EdgeImages(FE_TOP_RIGHT).Height) Or InXDrag Then
       If m_AllowResize Then ShowXResizeCursor = True
    End If
    
    If ShowXResizeCursor And ShowYResizeCursor Then
        Form.MousePointer = vbSizeNWSE
        
    ElseIf ShowXResizeCursor Then
        Form.MousePointer = vbSizeWE
    
    ElseIf ShowYResizeCursor Then
        Form.MousePointer = vbSizeNS
    
    Else
        Form.MousePointer = vbDefault
    End If

    If InXDrag Then
        ' Compute new number of horizontal segments
        NewXSlices = (X - BaseXSize - XDistance) / EdgeImages(FE_TOP_H_SEGMENT).Width
        If NewXSlices < MinXSlices Then NewXSlices = MinXSlices
        
        ' Check if we should actually do the resize. Not every
        ' slightest mouse drag should cause a resize
        If (NewXSlices - NumXSlices >= 0.5) Or _
           (NewXSlices - NumXSlices < -0.5) Then
            
            NumXSlices = NewXSlices
            ResizingNeeded = True
        End If
    End If

    ' Same handling for vertical resize-drag
    If InYDrag Then
        
        NewYSlices = (Y - BaseYSize - YDistance) / EdgeImages(FE_LEFT_V_SEGMENT).Height
        If NewYSlices < 0 Then NewYSlices = 0
        
        If NewYSlices - NumYSlices >= 0.5 Or _
           (NewYSlices - NumYSlices < -0.5) Then
            
            NumYSlices = NewYSlices
            ResizingNeeded = True
        End If
    End If

    If ResizingNeeded Then AdjustParentFormSize
    
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    MyExitButton.PaintUpImage
    MinimizeButton.PaintUpImage

    ' Test whether the user has released a "button",
    ' and commit the appropriate action if so
    If MyExitButton.HitTest(CLng(X), CLng(Y)) Then
       ' End
       Unload Form
    ElseIf MinimizeButton.HitTest(CLng(X), CLng(Y)) Then
        ' This will cause our form too to minimize
      Form.WindowState = vbMinimized
    End If

    ' Clear window dragging/resizing flags
    InXDrag = False
    InYDrag = False
    InFormDrag = False

End Sub



Friend Sub LoadSkin()
Dim i As Long
Dim Filename As String
Dim PrevXSliceSize As Long, PrevYSliceSize As Long

    ' Save for later. You'll see.
    If Not EdgeImages(0) Is Nothing Then
        PrevXSliceSize = EdgeImages(FE_TOP_H_SEGMENT).Width
        PrevYSliceSize = EdgeImages(FE_LEFT_V_SEGMENT).Height
    End If
    
    ' Initialize bitmaps array
    For i = 0 To FE_LAST
        Set EdgeImages(i) = New clsBitmap
    Next
    
    ' Load skin bitmaps. Check that the files actally  exist
    For i = 0 To FE_LAST
        Filename = CurrPrefs.SkinFullPath & arrImageFileNames(i)
        
        If Dir(Filename) = "" Then
            Err.Raise 1, , "Image file " & Filename & " not found!"
                        
        ElseIf EdgeImages(i).LoadFile(Filename) = False Then
            Err.Raise 1, , "Could not load image file: " & Filename
        End If
    Next
    
    ' Set back color according to skin's definition, to match
    ' the skin's "look"
    Form.BackColor = SkinPrefs.BackColor
    TargetPictureObject.BackColor = SkinPrefs.BackColor
    
    ' Prevent the checkbox from flickering when changing back color
    ' See documentation in start of file for all those variables
    BaseXSize = EdgeImages(FE_TOP_LEFT).Width + EdgeImages(FE_TOP_RIGHT).Width
    BaseYSize = EdgeImages(FE_TOP_LEFT).Height + EdgeImages(FE_BOTTOM_LEFT).Height

    XEdgeSize = EdgeImages(FE_RIGHT_V_SEGMENT).Width
    YEdgeSize = EdgeImages(FE_BOTTOM_H_SEGMENT).Height

    ' Here we compute how much horizontal/vertical segments
    ' sould be drawn
    If PrevXSliceSize <> 0 Then
        ' Skin was changed, match number of x/y slices
        ' according to the currect/previous sizes of the slices
        NumXSlices = Round(NumXSlices * PrevXSliceSize / EdgeImages(FE_TOP_H_SEGMENT).Width)
        NumYSlices = Round(NumYSlices * PrevYSliceSize / EdgeImages(FE_LEFT_V_SEGMENT).Height)
    Else
        ' Program was just loaded, match number of x/y slices
        ' to the default client width/height
'        NumXSlices = Round(DEFAULT_CLIENT_SIZE / EdgeImages(FE_TOP_H_SEGMENT).Width)
'        NumYSlices = Round(DEFAULT_CLIENT_SIZE / EdgeImages(FE_LEFT_V_SEGMENT).Height)
        NumXSlices = Round(DEFAULT_CLIENT_WIDTH / EdgeImages(FE_TOP_H_SEGMENT).Width)
        NumYSlices = Round(DEFAULT_CLIENT_HEIGHT / EdgeImages(FE_LEFT_V_SEGMENT).Height)
        
    End If
    
    ' Position Client Area
    TargetPictureObject.Top = EdgeImages(FE_TOP_LEFT).Height
    TargetPictureObject.Left = EdgeImages(FE_LEFT_V_SEGMENT).Width

    ' Initialize exit/minimize buttons
    MyExitButton.Init _
       CurrPrefs.SkinFullPath & "exitbutton_up.bmp", _
       CurrPrefs.SkinFullPath & "exitbutton_down.bmp", _
       SkinPrefs.ExitButtonX, SkinPrefs.ExitButtonY, _
       Form

    MinimizeButton.Init _
       CurrPrefs.SkinFullPath & "minbutton_up.bmp", _
       CurrPrefs.SkinFullPath & "minbutton_down.bmp", _
       SkinPrefs.MinButtonX, SkinPrefs.MinButtonY, _
       Form

    ' Limit minimum number of X slices, in order to allow the
    ' buttons to be drawn correctly
    MinXSlices = FindMinXSlices()
    NumXSlices = IIf(MinXSlices > NumXSlices, MinXSlices, NumXSlices)

    ' Create and store region data for each of the skin bitmaps,
    ' for use whenever creating the window region
    Dim LoadedRegionsFromFile As Boolean
    
    ' If the 'load region data from file' box is checked, try loading region data
    ' from a cache file. if the file does not exist yet, we'll create the regions
    ' and save them - for the next time
   ' If chkUseTransFile.Value Then
        If LoadEdgeRegions(EdgeRegions, CurrPrefs.SkinFullPath & "trans.dat") Then
            LoadedRegionsFromFile = True
        End If
   ' End If
    
    If Not LoadedRegionsFromFile Then
        For i = 0 To FE_LAST
            CreateRegionData EdgeImages(i), EdgeRegions(i)
        Next
    
        SaveEdgeRegions EdgeRegions, CurrPrefs.SkinFullPath & "trans.dat"
    End If

End Sub

Private Sub Form_Paint()
 If Not TargetPictureObject Is Nothing Then
    If Not NoRedraw Then
        DrawEdges Form, EdgeImages, NumXSlices, NumYSlices, False
    
        MyExitButton.PaintUpImage 'Paint the X button
        MinimizeButton.PaintUpImage 'and Minmized Button
        
    End If
 End If
End Sub

Friend Sub AdjustParentFormSize()
Dim NewSize As Long
    
    ' We don't want form redraws when in middle of new size
    ' setting, before the new region was set
    NoRedraw = True
    
    ' Compute width/height of form accodring to the number of
    ' x/y slices
            Form.Width = (EdgeImages(FE_TOP_LEFT).Width + EdgeImages(FE_TOP_H_SEGMENT).Width _
                        * NumXSlices + EdgeImages(FE_TOP_RIGHT).Width) * Screen.TwipsPerPixelX
            Form.Height = (EdgeImages(FE_TOP_LEFT).Height + EdgeImages(FE_LEFT_V_SEGMENT).Height _
                        * NumYSlices + EdgeImages(FE_BOTTOM_LEFT).Height) * Screen.TwipsPerPixelY
            ' Compute size of client area i.e., Picture Box

            NewSize = EdgeImages(FE_LEFT_V_SEGMENT).Height * NumYSlices
            TargetPictureObject.Height = NewSize
            
            NewSize = (Form.Width / Screen.TwipsPerPixelX) - EdgeImages(FE_LEFT_V_SEGMENT).Width - EdgeImages(FE_RIGHT_V_SEGMENT).Width
            TargetPictureObject.Width = NewSize
   
    NoRedraw = False
    
    ' Create new window region. Also triggers a redraw, now that
    ' wer'e done setting the new form shape
    BuildWindowRegion

End Sub

Private Sub BuildWindowRegion()
Dim PrevRegion As Long

    PrevRegion = WindowRegion
    
    ' Create initial region that covers the client area
    WindowRegion = CreateRectRgn(TargetPictureObject.Left, TargetPictureObject.Top, TargetPictureObject.Left + TargetPictureObject.Width, TargetPictureObject.Top + TargetPictureObject.Height)

    ' Add to it the window region of the form edges
    BuildEdgesRegion WindowRegion, EdgeImages, EdgeRegions, NumXSlices, NumYSlices

    ' Finally - set the full region
    SetWindowRgn Form.hwnd, WindowRegion, True
    
    ' Don't forget - delete old window region
    DeleteObject PrevRegion
    
End Sub

' Fill the list of skins.
' Actually it's a list of directories under App.Path
Private Sub ListSkins()
Dim CurrSkinName As String, SkinPos As Long
Dim i As Long

    InListSkins = True
    
    CurrSkinName = Dir(SkinPathName, vbDirectory)
     lstSkins.Clear
    Do While CurrSkinName <> ""
    
        If CurrSkinName <> "." And CurrSkinName <> ".." Then
            If (GetAttr(SkinPathName & CurrSkinName) And vbDirectory) Then
                lstSkins.AddItem CurrSkinName
            
                If CurrSkinName = CurrPrefs.SkinName Then
                    SkinPos = i
                End If
                
                i = i + 1
            End If
        End If
        
        CurrSkinName = Dir()
    Loop
    
    ' Visually select 'default' skin
    InListSkins = False
    lstSkins.ListIndex = SkinPos
    

End Sub

Private Sub Form_Unload(Cancel As Integer)


    Dim i As Long
    Set DockHandler = Nothing
    For i = 0 To FE_LAST
        Set EdgeImages(i) = Nothing
    Next
    


End Sub

Private Sub lstSkins_Click()

    If Not InListSkins Then
        CurrPrefs.SkinName = lstSkins.Text
        ReadSkinPreferences
        LoadAllSkins
        ColorAdjust
        Form_Paint
        RaiseEvent SkinSelected
    End If
 
End Sub
Private Sub ColorAdjust(Optional Visible As Boolean = True)
 Dim CTL As Control
 If Visible = True Then
        For Each CTL In Form.Controls
            If TypeOf CTL Is Frame Or TypeOf CTL Is CheckBox Or TypeOf CTL Is OptionButton Then
               CTL.BackColor = Form.BackColor
               CTL.ForeColor = Form.ForeColor
               
            End If
           
        Next
 Else
        For Each CTL In Form.Controls
            If TypeOf CTL Is Frame Then
               CTL.Visible = Visible
            End If
        Next
End If
End Sub

' Find out the minimum number of horizontal slices
' that allows the buttons to be drawn correctly
Private Function FindMinXSlices() As Long
Dim MinSize As Long
Dim MinButtonSize As Long, ExitButtonSize As Long
    
    If MinimizeButton.X >= 0 Then
        ' Button is attached to top-left corner.
        ' Find out the width of the part of the button that
        ' excceds the top-left part width
        MinButtonSize = MinimizeButton.X + MinimizeButton.Width - _
            EdgeImages(FE_TOP_LEFT).Width
    Else
        ' Button is attached to top-RIGHT corner (its X value
        ' is relative to the right side).
        ' Find out the width of the part of the button that
        ' excceds the top-right part width
        MinButtonSize = Abs(MinimizeButton.X) - _
            EdgeImages(FE_TOP_RIGHT).Width
    End If

    ' Same handling for the exit button
    If MyExitButton.X >= 0 Then
        ExitButtonSize = MyExitButton.X + MyExitButton.Width - _
            EdgeImages(FE_TOP_LEFT).Width
    Else
        ExitButtonSize = Abs(MyExitButton.X) - _
            EdgeImages(FE_TOP_RIGHT).Width
    End If
    
    MinSize = IIf(MinButtonSize > ExitButtonSize, MinButtonSize, ExitButtonSize)

    ' Find out how many slices are needed
    FindMinXSlices = RoundUp(MinSize / EdgeImages(FE_TOP_H_SEGMENT).Width)
    
End Function

' Given a double number, the function always returns a long
' number that is the rounding UP of the double value
Private Function RoundUp(Number As Double) As Long
    RoundUp = IIf(Number - CLng(Number) <> 0, CLng(Number + 0.5), CLng(Number))
End Function





Private Sub UserControl_EnterFocus()
    UserControl.MousePointer = vbDefault
End Sub


Private Sub UserControl_InitProperties()
   
            UserControl.ForeColor = Ambient.BackColor
            Set UserControl.Font = Ambient.Font
            Set mvar_picClientArea = Nothing
            m_AnimationApplied = True
            KeyPreview = True
            m_AllowResize = False
            m_Enabled = True
            m_TransitionType = Zoomout
            m_Delay = 500
            m_Applythemes = True
End Sub




Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 93 Then 'win Menu Key
        EdSkin.Visible = True
        PopupMenu mnuSkinProp, vbLeftButton, UserControl.ScaleLeft, UserControl.ScaleTop
    End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        EdSkin.Visible = True
        PopupMenu mnuSkinProp, vbLeftButton
    End If

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

        UserControl.ForeColor = Ambient.BackColor
        KeyPreview = True
        m_TransitionType = PropBag.ReadProperty("TransitionType", m_TransitionType)
        CurrPrefs.SkinName = PropBag.ReadProperty("DefaultSkinName", "WinDefault")
        m_Delay = PropBag.ReadProperty("TransitionDelay", m_Delay)
      
        'Skin Path is Added on 12th July
        CurrPrefs.SkinsPath = PropBag.ReadProperty("SkinPathName", m_SkinPathName)
          
        m_AllowResize = PropBag.ReadProperty("AllowResize", m_AllowResize)
    Set ClientObject = PropBag.ReadProperty("ClientObject", mvar_picClientArea)
        ApplyAnimation = PropBag.ReadProperty("ApplyAnimation", m_AnimationApplied)
        m_Enabled = PropBag.ReadProperty("Enabled", m_Enabled)
          'Added on 29th July
        m_Applythemes = PropBag.ReadProperty("ApplyThemes", m_Applythemes)
        m_TrackMouseMove = PropBag.ReadProperty("TrackMouseMove", m_TrackMouseMove)
        If Ambient.UserMode = True And m_Enabled Then
            Set Form = Parent
        End If
End Sub

Private Sub UserControl_Resize()
    Size 1845, 375
End Sub


Private Sub UserControl_Terminate()

    NumXSlices = 0
    NumYSlices = 0
    MinXSlices = 0
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
        KeyPreview = True
        UserControl.ForeColor = Ambient.BackColor
        PropBag.WriteProperty "TransitionType", m_TransitionType
        PropBag.WriteProperty "TransitionDelay", m_Delay
        PropBag.WriteProperty "AllowResize", m_AllowResize
        PropBag.WriteProperty "ApplyAnimation", m_AnimationApplied
        PropBag.WriteProperty "ClientObject", mvar_picClientArea
        PropBag.WriteProperty "Enabled", m_Enabled
        PropBag.WriteProperty "SkinPathName", m_SkinPathName
        PropBag.WriteProperty "ApplyThemes", m_Applythemes
        PropBag.WriteProperty "TrackMouseMove", m_TrackMouseMove
End Sub
Private Sub LoadAllSkins()
On Error GoTo LoadAllSkins_ErrHandler
    LoadSkin
    AdjustParentFormSize
    Exit Sub
LoadAllSkins_ErrHandler:
    MsgBox "Unable to load skin. " & vbCrLf & _
        "Reason: " & Err.Description, vbCritical, "Skin Loading Error"
End Sub

Public Property Set ClientObject(data As Object)
Set mvar_picClientArea = data
PropertyChanged "ClientObject"
End Property

Public Property Get ClientObject() As Object
Set ClientObject = mvar_picClientArea
PropertyChanged "ClientObject"
End Property
Private Property Get TargetPictureObject() As Object
    Set TargetPictureObject = mvar_picClientArea
End Property

Public Property Let ApplyAnimation(data As Boolean)
Attribute ApplyAnimation.VB_Description = "Add Animation Effect to  your Form.  \r\nA Like Fadeout,Slide etc. "
 m_AnimationApplied = data
 PropertyChanged "ApplyAnimation"
End Property

Public Property Get ApplyAnimation() As Boolean
    ApplyAnimation = m_AnimationApplied
End Property
Public Property Let AllowResize(data As Boolean)
Attribute AllowResize.VB_Description = "To Make Parent Form  Resizable \r\n set it to True otherwise False  "
 m_AllowResize = data
 PropertyChanged "AllowResize"
End Property

Public Property Get AllowResize() As Boolean
    AllowResize = m_AllowResize
End Property
Public Property Let DefaultSkinName(data As String)
Attribute DefaultSkinName.VB_Description = "This Property can be Set through the Code Only. \r\n It will always Show 'Classic' on property Window"
If Trim$(data) <> "" Then
    CurrPrefs.SkinName = data
    CurrPrefs.SkinsPath = App.Path + "\skins\"
    DefaultSkinSupplied = True
 End If
End Property

Public Property Get DefaultSkinName() As String
 DefaultSkinName = CurrPrefs.SkinName
End Property

Public Property Let Enabled(data As Boolean)
Attribute Enabled.VB_Description = "if you Set this Property to True \r\n Skin will be applied to the Parent Form otherwise The \r\n Orignal form will be Displayed"
    m_Enabled = data
    PropertyChanged "Enabled"
End Property
Public Property Get Enabled() As Boolean
 Enabled = m_Enabled
End Property
Public Property Let SkinPathName(data As String)
    m_SkinPathName = data
    PropertyChanged "SkinPathName"
End Property
Public Property Get SkinPathName() As String
If Trim$(m_SkinPathName) = "" Then 'if Skin Path is Not Supplied
    SkinPathName = App.Path + "\skins\"
Else
    SkinPathName = m_SkinPathName
End If
End Property

Public Property Let TransitionType(data As TrasitionEnum)
Attribute TransitionType.VB_Description = "These are the Transition Types Supported By SkinControl \r\n Default is ZoomOut"
    m_TransitionType = data
    PropertyChanged "TransitionType"
End Property
Public Property Get TransitionType() As TrasitionEnum
 TransitionType = m_TransitionType
End Property

Public Property Get TransitionDelay() As Long
        TransitionDelay = m_Delay
End Property
Public Property Let TransitionDelay(data As Long)
Attribute TransitionDelay.VB_Description = "Transition Delay (in MilliSeconds) \r\n Give a negative Number and See "
    If data < 0 Then
        Err.Raise vbObjectError + 91, "Validation", "Delay Cannot Be negative , Making it to Default Value"
        m_Delay = 500
        Exit Property
    End If
    m_Delay = data
    PropertyChanged "TransitionDelay"
End Property
'Public Sub MakeTransparentBackGround(ControlType As String)
'    Dim cnt As Long
'    For cnt = 1 To frm.Controls.Count
'        If LCase(TypeName(frm.Controls(cnt))) = ControlType Then
'                frm.Controls(cnt).BackColor = frm.BackColor
'        End If
'    End If
'End Sub

Public Property Let ApplyThemes(data As Boolean)

    m_Applythemes = data
    PropertyChanged "ApplyThemes"
End Property

Public Property Get ApplyThemes() As Boolean
    ApplyThemes = m_Applythemes
End Property

Public Property Let TrackMouseMove(data As Boolean)
        If (m_TrackMouseMove = True And data = False) And Ambient.UserMode Then
                ReleaseMouseMove Form
        End If
               
        m_TrackMouseMove = data
        PropertyChanged "TrackMouseMove"

End Property
Public Property Get TrackMouseMove() As Boolean
       TrackMouseMove = m_TrackMouseMove
End Property
