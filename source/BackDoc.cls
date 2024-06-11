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
        
        MakeBeam Source.BoundingBox, Source.TopY - BEAM_THICKNESS
        MakeBeam Source.BoundingBox, Source.BottomY
        
        Dim IntShapes As ShapeRange: Set IntShapes = ContourInt.UngroupAllEx
        NameShapes IntShapes, BACK_CONTOUR_INT_NAME
        Dim ExtShapes As ShapeRange: Set ExtShapes = ContourExt.BreakApartEx
        NameShapes ExtShapes, BACK_CONTOUR_EXT_NAME
        
        With New InnerGrooves
            Set .Shapes = IntShapes
            .MakeGrooves
        End With
        With New BottomGrooves
            Set .Shapes = ShapesNotInside(ExtShapes)
            .MakeBottomGrooves
        End With
        
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

Private Function MakeBeam( _
                     ByVal ShapesRect As Rect, _
                     ByVal Bottom As Double _
                 ) As Shape
    Set MakeBeam = _
        ActiveLayer.CreateRectangle( _
            ShapesRect.Left, Bottom + BEAM_THICKNESS, ShapesRect.Right, Bottom _
        )
    MakeBeam.Fill.ApplyNoFill
    MakeBeam.Name = HORIZONTAL_BEAM_NAME
End Function

Private Property Get ShapesNotInside( _
                         ByVal Shapes As ShapeRange _
                     ) As ShapeRange
    Set ShapesNotInside = CreateShapeRange
    Dim Shape As Shape
    For Each Shape In Shapes
        If Not ShapeInsideAny(Shape, Shapes) Then ShapesNotInside.Add Shape
    Next Shape
End Property

Private Property Get ShapeInsideAny( _
                         ByVal Shape As Shape, _
                         ByVal Shapes As ShapeRange _
                     ) As Boolean
    Dim CurrentShape As Shape
    For Each CurrentShape In Shapes
        If Not CurrentShape Is Shape Then
            If IsAllNodesInside(Shape.Curve.Nodes.All, CurrentShape.Curve) Then
                ShapeInsideAny = True
                Exit Property
            End If
        End If
    Next CurrentShape
End Property

Private Property Get IsAllNodesInside( _
                         ByVal Nodes As NodeRange, _
                         ByVal Curve As Curve _
                     ) As Boolean
    Dim Node As Node
    For Each Node In Nodes
        If Not Curve.IsPointInside(Node.PositionX, Node.PositionY) Then _
            Exit Property
    Next Node
    IsAllNodesInside = True
End Property
