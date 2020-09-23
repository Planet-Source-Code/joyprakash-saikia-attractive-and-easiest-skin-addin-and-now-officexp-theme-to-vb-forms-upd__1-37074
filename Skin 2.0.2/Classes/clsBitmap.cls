VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBitmap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
'Copyright  :   Free
'Module     :   clsBitmap
'Author     :   JoyPrakash Saikia
'Created    :   15/07/2002
'Purpose    :
' Based on  Steve McMahon's work (vbAccelerator.com)
'---------------------------------------------------------------------------------------
Option Explicit

'Class Creats the Bitmaps

Private mDC As Long         ' Memory DC
Private mBitmap As Long     ' Bitmap handle
Private mOldBitmap As Long  ' "Original" Bitmap handle

Private mWidth As Long
Private mHeight As Long

Public Function LoadFile(Filename As String) As Boolean

    ' Clear up previous DC/bitmap
    ClearAll
    
    mBitmap = LoadImage(NULLHANDLE, Filename, IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE Or LR_LOADMAP3DCOLORS)
    
    If (mBitmap <> NULLHANDLE) Then
        LoadFile = LoadBitmapIntoDC
    End If
    
End Function


Private Function LoadBitmapIntoDC() As Boolean
Dim ScreenDC As Long
Dim BitmapData As BITMAP

    ' Create a compatible memory DC to hold the bitmap
    ScreenDC = GetDC(NULLHANDLE)
    mDC = CreateCompatibleDC(ScreenDC)
    ReleaseDC NULLHANDLE, ScreenDC
    
    If (mDC <> NULLHANDLE) Then
        ' If the DC was created successfully,
        ' select the bitmap into it
        mOldBitmap = SelectObject(mDC, mBitmap)
        
        ' Get the dimensions of the bitmap
        APIGetObject mBitmap, Len(BitmapData), BitmapData
        mWidth = BitmapData.bmWidth
        mHeight = BitmapData.bmHeight
        
        LoadBitmapIntoDC = True
    End If

End Function

Property Get Width() As Long
    Width = mWidth
End Property

Property Get Height() As Long
    Height = mHeight
End Property

Property Get hDC() As Long
    hDC = mDC
End Property

Public Sub Paint(DestDC As Long, _
                 DestX As Long, _
                 DestY As Long)
    
    BitBlt DestDC, DestX, DestY, _
        mWidth, mHeight, mDC, _
        0, 0, vbSrcCopy
        
End Sub

Private Sub ClearAll()
    
    If (mDC <> NULLHANDLE) Then
        If (mBitmap <> NULLHANDLE) Then
            ' Select the original bitmap into the DC,
            ' and delete our bitmap
            SelectObject mDC, mOldBitmap
            DeleteObject mBitmap
            mBitmap = NULLHANDLE
        End If
            ' Delete the memory DC
            DeleteObject mDC
            mDC = NULLHANDLE
    End If
End Sub

Private Sub Class_Terminate()
    ClearAll
End Sub
