Attribute VB_Name = "modRectFrmRegion"
'---------------------------------------------------------------------------------------
'Copyright  :   JoyPrakash Saikia 2002
'Module     :   modRectFrmRegion
'Author     :   JoyPrakash Saikia
'Created    :   15/07/2002
'Purpose    :   Make the Form Rectangle
'---------------------------------------------------------------------------------------
Option Explicit

Public Enum FormEdges
    FE_TOP_LEFT
    FE_TOP_RIGHT
    FE_BOTTOM_LEFT
    FE_BOTTOM_RIGHT
    FE_TOP_H_SEGMENT
    FE_BOTTOM_H_SEGMENT
    FE_RIGHT_V_SEGMENT
    FE_LEFT_V_SEGMENT
    FE_LAST = FE_LEFT_V_SEGMENT
End Enum

Type RegionDataType
    RegionData() As Byte
    DataLength As Long
End Type

Public arrImageFileNames(FE_LAST) As String


Public Sub IntImageFileNames()
        '************************************************************
        'Description:
        ' This is the Initialization Routine for Skin Loading
        ' I've used fixed Naming convention for the diffrent part Skins
        

    arrImageFileNames(FE_TOP_LEFT) = "top_left.bmp"
    arrImageFileNames(FE_TOP_RIGHT) = "top_right.bmp"
    arrImageFileNames(FE_BOTTOM_LEFT) = "bottom_left.bmp"
    arrImageFileNames(FE_BOTTOM_RIGHT) = "bottom_right.bmp"
    arrImageFileNames(FE_TOP_H_SEGMENT) = "hsegment_top.bmp"
    arrImageFileNames(FE_BOTTOM_H_SEGMENT) = "hsegment_bottom.bmp"
    arrImageFileNames(FE_RIGHT_V_SEGMENT) = "vsegment_right.bmp"
    arrImageFileNames(FE_LEFT_V_SEGMENT) = "vsegment_left.bmp"

End Sub





Public Sub BuildEdgesRegion(WindowRegion As Long, _
                            EdgeImages() As clsBitmap, _
                            EdgeRegions() As RegionDataType, _
                            NumXSlices As Long, _
                            NumYSlices As Long)
        '************************************************************
        'Description:
        ' This function builds the window region of the form's edges -
        ' the corners and the sides, using the pre-created regions data
        ' Each created region is combined with the full window region
        '************************************************************
Dim i As Long

    ' Make region for top-left corner. That's an easy one
    MakeRegionWithOffset EdgeRegions(FE_TOP_LEFT), 0, 0, WindowRegion

    ' Top-right corner
    MakeRegionWithOffset EdgeRegions(FE_TOP_RIGHT), _
        EdgeImages(FE_TOP_LEFT).Width + (EdgeImages(FE_TOP_H_SEGMENT).Width * NumXSlices), 0, _
        WindowRegion
    
    ' Bottom-left corner
    MakeRegionWithOffset EdgeRegions(FE_BOTTOM_LEFT), 0, _
        EdgeImages(FE_TOP_LEFT).Height + EdgeImages(FE_LEFT_V_SEGMENT).Height * NumYSlices, _
        WindowRegion
    
    ' Bottom-right corner
    MakeRegionWithOffset EdgeRegions(FE_BOTTOM_RIGHT), _
        EdgeImages(FE_TOP_LEFT).Width + (EdgeImages(FE_TOP_H_SEGMENT).Width * NumXSlices), _
        EdgeImages(FE_TOP_LEFT).Height + EdgeImages(FE_LEFT_V_SEGMENT).Height * NumYSlices, _
        WindowRegion

    ' Create the regions for the top and bottom sides,
    ' By the number of X slices.
    For i = 1 To NumXSlices
        MakeRegionWithOffset EdgeRegions(FE_TOP_H_SEGMENT), EdgeImages(FE_TOP_LEFT).Width + EdgeImages(FE_TOP_H_SEGMENT).Width * (i - 1), 0, WindowRegion
        
        MakeRegionWithOffset EdgeRegions(FE_BOTTOM_H_SEGMENT), _
            EdgeImages(FE_TOP_LEFT).Width + EdgeImages(FE_TOP_H_SEGMENT).Width * (i - 1), _
            EdgeImages(FE_TOP_LEFT).Height + EdgeImages(FE_BOTTOM_LEFT).Height + EdgeImages(FE_LEFT_V_SEGMENT).Height * NumYSlices - EdgeImages(FE_BOTTOM_H_SEGMENT).Height, WindowRegion
    Next

    ' Create the regions for the left and right sides,
    ' By the number of Y slices.
    For i = 1 To NumYSlices
        MakeRegionWithOffset EdgeRegions(FE_LEFT_V_SEGMENT), 0, EdgeImages(FE_TOP_LEFT).Height + EdgeImages(FE_LEFT_V_SEGMENT).Height * (i - 1), WindowRegion
        
        MakeRegionWithOffset EdgeRegions(FE_RIGHT_V_SEGMENT), _
            EdgeImages(FE_TOP_LEFT).Width + EdgeImages(FE_TOP_H_SEGMENT).Width * NumXSlices + EdgeImages(FE_TOP_RIGHT).Width - EdgeImages(FE_RIGHT_V_SEGMENT).Width, _
            EdgeImages(FE_TOP_LEFT).Height + EdgeImages(FE_LEFT_V_SEGMENT).Height * (i - 1), _
            WindowRegion
    Next
    
End Sub


Public Sub DrawEdges(DestForm As Form, _
                     EdgeImages() As clsBitmap, _
                     NumXSlices As Long, NumYSlices As Long, Optional Value As Boolean)

Dim i As Long
        'Description:
        ' This fucntion is almost identical to MakeEdgesRegion,
        ' excepts that it actually Puts the edges to the SCreen.
    EdgeImages(FE_TOP_LEFT).Paint DestForm.hDC, 0, 0

    EdgeImages(FE_TOP_RIGHT).Paint DestForm.hDC, _
        EdgeImages(FE_TOP_LEFT).Width + EdgeImages(FE_TOP_H_SEGMENT).Width * NumXSlices, 0
    
    EdgeImages(FE_BOTTOM_LEFT).Paint DestForm.hDC, _
        0, EdgeImages(FE_TOP_LEFT).Height + EdgeImages(FE_LEFT_V_SEGMENT).Height * NumYSlices
    
    EdgeImages(FE_BOTTOM_RIGHT).Paint DestForm.hDC, _
        EdgeImages(FE_TOP_LEFT).Width + EdgeImages(FE_TOP_H_SEGMENT).Width * NumXSlices, _
        EdgeImages(FE_TOP_LEFT).Height + EdgeImages(FE_LEFT_V_SEGMENT).Height * NumYSlices
    
    For i = 1 To NumXSlices
        EdgeImages(FE_TOP_H_SEGMENT).Paint DestForm.hDC, _
            EdgeImages(FE_TOP_LEFT).Width + EdgeImages(FE_TOP_H_SEGMENT).Width * (i - 1), 0
        EdgeImages(FE_BOTTOM_H_SEGMENT).Paint DestForm.hDC, _
            EdgeImages(FE_TOP_LEFT).Width + EdgeImages(FE_TOP_H_SEGMENT).Width * (i - 1), _
            EdgeImages(FE_TOP_LEFT).Height + EdgeImages(FE_BOTTOM_LEFT).Height + EdgeImages(FE_LEFT_V_SEGMENT).Height * NumYSlices - EdgeImages(FE_BOTTOM_H_SEGMENT).Height
    Next

    For i = 1 To NumYSlices
        EdgeImages(FE_LEFT_V_SEGMENT).Paint DestForm.hDC, 0, _
            EdgeImages(FE_TOP_LEFT).Height + EdgeImages(FE_LEFT_V_SEGMENT).Height * (i - 1)
        EdgeImages(FE_RIGHT_V_SEGMENT).Paint DestForm.hDC, _
            DestForm.ScaleWidth - EdgeImages(FE_RIGHT_V_SEGMENT).Width, _
            EdgeImages(FE_TOP_LEFT).Height + EdgeImages(FE_LEFT_V_SEGMENT).Height * (i - 1)
    Next
    
End Sub


Public Sub SaveEdgeRegions(EdgeRegions() As RegionDataType, _
                           FileName As String)

Dim i As Long
'This method typically yields about 50% speed increase in loading skins.
' NOTE: If you change the transparent areas in a bitmap, the file will be outdated.


    Open FileName For Binary As #1

    For i = 0 To FE_LAST
        Put 1, , EdgeRegions(i).DataLength 'put the Data
        Put 1, , EdgeRegions(i).RegionData
    Next

    Close
    
End Sub

' Load the edges' region data from file
Public Function LoadEdgeRegions(EdgeRegions() As RegionDataType, _
                                FileName As String) As Boolean

Dim i As Long
    
    If Dir(FileName) = "" Then Exit Function
    
    Open FileName For Binary As #1
    
    For i = 0 To FE_LAST
        Get 1, , EdgeRegions(i).DataLength
        ReDim EdgeRegions(i).RegionData(EdgeRegions(i).DataLength + 32)
        Get 1, , EdgeRegions(i).RegionData
    Next
    
    Close
    
    LoadEdgeRegions = True

End Function

