VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BackDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'===============================================================================
' # Fields

Public BeamEdgeOffset As Double

'===============================================================================
' # Public

Public Sub ProcessBackDoc()
    With ActiveDocument
        .Unit = cdrMillimeter
        Dim Source As Shape: Set Source = ActivePage.Shapes.All.Group
        ResizePageToShapes SideAdd:=AverageDim(Source) / 10
        Source.Fill.ApplyNoFill
        
        Dim ContourInt As Shape: Set ContourInt = _
            MakeContour(Source, -BACK_CONTOUR_INT)
        ContourInt.Outline.Color = CreateColor(BACK_COLOR)
        
        Dim ContourExt As Shape: Set ContourExt = _
            MakeContour(Source, BACK_CONTOUR_EXT)
        ContourExt.Outline.Color = CreateColor(BACK_COLOR)
                       
        Dim IntShapes As ShapeRange: Set IntShapes = ContourInt.UngroupAllEx
        NameShapes IntShapes, BACK_CONTOUR_INT_NAME
        
        ContourExt.Name = BACK_CONTOUR_EXT_NAME
        Dim ExtShapes As ShapeRange: Set ExtShapes = ContourExt.BreakApartEx
        
        With New InnerGrooves
            Set .Shapes = IntShapes
            .MakeGrooves
        End With
        With New BottomGrooves
            Set .Shapes = FindShapesNotInside(ExtShapes)
            .MakeBottomGrooves
        End With
        
        Set IntShapes = IntShapes.BreakApartEx
        MakeBeams FindShapesNotInBounds(IntShapes)
        
        .Name = BACK_NAME
        Source.Delete
    End With
End Sub

'===============================================================================
' # Helpers

Private Function MakeContour( _
                     ByVal Shape As Shape, _
                     ByVal Offset As Double _
                 ) As Shape
    Dim Direction As cdrContourDirection
    If Offset > 0 Then
        Direction = cdrContourOutside
    ElseIf Offset < 0 Then
        Direction = cdrContourInside
    Else
        Exit Function
    End If
    With Shape.CreateContour( _
            Direction:=Direction, _
            Offset:=Abs(Offset), _
            Steps:=1 _
        )
        '.Contour.CornerType = cdrContourCornerRound
        .Contour.OutlineColor = CreateColor(BACK_COLOR)
        Set MakeContour = .Separate(1)
    End With
End Function

Private Function MakeBeams(ByVal Shapes As ShapeRange) As ShapeRange
    Shapes.Sort "@Shape1.Top < @Shape2.Top"
    Dim TopBeam As Shape: Set TopBeam = _
        MakeBeam( _
            Shapes.BoundingBox, _
            Shapes.FirstShape.TopY - BEAM_THICKNESS - BeamEdgeOffset _
        )
        
    Shapes.Sort "@Shape1.Bottom > @Shape2.Bottom"
    Dim BottomBeamY As Double: BottomBeamY = _
        Shapes.FirstShape.BottomY + BeamEdgeOffset
    Dim MultDistance As Double: MultDistance = _
        ClosestDividend(TopBeam.BottomY - BottomBeamY, QUANTIZATION_STEP)
    
    Dim BottomBeam As Shape: Set BottomBeam = _
        MakeBeam( _
            Shapes.BoundingBox, _
            TopBeam.BottomY - MultDistance _
        )
    
    Set MakeBeams = CreateShapeRange
    MakeBeams.Add TopBeam
    MakeBeams.Add BottomBeam
End Function

Private Function MakeBeam( _
                     ByVal ShapesRect As Rect, _
                     ByVal Bottom As Double _
                 ) As Shape
    Set MakeBeam = _
        ActiveLayer.CreateRectangle( _
            ShapesRect.Left, Bottom + BEAM_THICKNESS, ShapesRect.Right, Bottom _
        )
    ApplyBeamCommonProps MakeBeam
    MakeBeam.Name = HORIZONTAL_BEAM_NAME
End Function

'шейпы, которые не находятся целиком внутри крайних координат других по оси X
Private Property Get FindShapesNotInBounds( _
                         ByVal Shapes As ShapeRange _
                     ) As ShapeRange
    Set FindShapesNotInBounds = CreateShapeRange
    Dim Shape As Shape
    For Each Shape In Shapes
        If Not ShapeInBoundsOfAny(Shape, Shapes) Then _
            FindShapesNotInBounds.Add Shape
    Next Shape
End Property

Private Property Get ShapeInBoundsOfAny( _
                         ByVal Shape As Shape, _
                         ByVal Shapes As ShapeRange _
                     ) As Boolean
    Dim CurrentShape As Shape
    For Each CurrentShape In Shapes
        If Not CurrentShape Is Shape Then
            If ShapeInBoundsOf(Shape, CurrentShape) Then
                ShapeInBoundsOfAny = True
                Exit Property
            End If
        End If
    Next CurrentShape
End Property

Private Property Get ShapeInBoundsOf( _
                         ByVal Shape1 As Shape, _
                         ByVal Shape2 As Shape _
                     ) As Boolean
    ShapeInBoundsOf = _
        (Shape1.LeftX > Shape2.LeftX) _
    And (Shape1.RightX < Shape2.RightX)
End Property
