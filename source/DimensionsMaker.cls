VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DimensionsMaker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'===============================================================================
' # Fields

Public Result As ShapeRange

Private Shapes As ShapeRange
Private TextSize As Double

'===============================================================================
' # Public

Public Sub MakeDimensions()
    Set Shapes = ActivePage.Shapes.All
    
    Dim IntExtShapes As New ShapeRange
    IntExtShapes.AddRange FindShapesByName(Shapes, BACK_CONTOUR_INT_NAME)
    IntExtShapes.AddRange FindShapesByName(Shapes, BACK_CONTOUR_EXT_NAME)
    Dim MakeupHeight As Double: MakeupHeight = IntExtShapes.SizeHeight
    Dim Bottom As Double: Bottom = IntExtShapes.BottomY
    Dim Color As Color: Set Color = CreateColor(DIMENSION_SHAPES_COLOR)
    
    SetOutlineColor IntExtShapes, CreateColor(DIMENSION_SHAPES_COLOR)
    
    Dim BottomHoles As ShapeRange: Set BottomHoles = _
        FindShapesByName(Shapes, BOTTOM_HOLE_NAME)
    SetNoOutline BottomHoles

    If BottomHoles.Count = 0 Then Exit Sub
    Set Result = CreateShapeRange
    
    BottomHoles.Sort "@Shape1.Left < @Shape2.Left"
    
    Dim Offset As Double: Offset = MakeupHeight * DIMENSION_OFFSET_MULT
    Dim PositionY As Double: PositionY = Bottom - Offset
    TextSize = MakeupHeight * DIMENSION_TEXT_SIZE_MULT
        
    Dim BottomBeamY As Double
    With FindBeams(Shapes)
        If .Some Then
            Result.Add MakeDim( _
                .BottomBeam.SnapPointsOfType(cdrSnapPointBBox)(4), _
                BottomHoles(1).SnapPointsOfType(cdrSnapPointBBox)(5), _
                PositionY _
            )
        End If
    End With
        
    If BottomHoles.Count = 1 Then Exit Sub
    Dim CurrentPos As Double
    Dim Other As Boolean
    Dim Index As Long
    For Index = 2 To BottomHoles.Count
        If Other Then CurrentPos = PositionY Else CurrentPos = PositionY - Offset
        Other = Not Other
        Result.Add MakeDim( _
            BottomHoles(Index - 1).SnapPointsOfType(cdrSnapPointBBox)(5), _
            BottomHoles(Index).SnapPointsOfType(cdrSnapPointBBox)(5), _
            CurrentPos _
        )
    Next Index
    
    ActivePage.SizeHeight = MakeupHeight + Offset * 7
End Sub

'===============================================================================
' # Helpers

Private Function MakeDim( _
                     ByVal Point1 As SnapPoint, _
                     ByVal Point2 As SnapPoint, _
                     ByVal PositionY As Double _
                 ) As Shape
    Set MakeDim = _
        ActiveLayer.CreateLinearDimension( _
            Type:=cdrDimensionHorizontal, _
            Point1:=Point1, _
            Point2:=Point2, _
            TextCentered:=True, _
            TextX:=0, TextY:=PositionY, _
            Style:=cdrDimensionStyleDecimal, _
            Units:=cdrDimensionUnitMM, _
            Placement:=cdrDimensionWithinLine, _
            TextSize:=TextSize _
        )
    SetDimensionPrecision MakeDim, 0
    SetDimensionShowUnits MakeDim, False
    
    'Set MakeDim = MakeDim.BreakApartEx
End Function
