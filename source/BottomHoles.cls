VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BottomHoles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'===============================================================================
' # Fields

Public TopBeam As Shape
Public BottomBeam As Shape
Public Shapes As ShapeRange
Public EdgeSecurity As Double

'===============================================================================
' # Public

Public Sub MakeBottomHoles()
    Dim Count As Long: Count = Shapes.Count
    Dim Center As Point
    If Count = 0 Then Exit Sub
    
    Set Center = FindPointForHoleInShape(Shapes.FirstShape)
    TryMakeBottomCircle Center
    TryMakeVerticalBeam Center
    If Count = 1 Then Exit Sub
    
    Set Center = FindPointForHoleInShape(Shapes.LastShape)
    TryMakeBottomCircle Center
    TryMakeVerticalBeam Center
    If Count = 2 Then Exit Sub
    
    Dim StartingX As Double: StartingX = 0
    Dim Index As Long
    For Index = 2 To Count - 1
        Set Center = FindPointForHoleInShape(Shapes(Index))
        TryMakeBottomCircle Center
        If Not Center Is Nothing Then
            If Center.x - StartingX >= VERTICAL_BEAM_STEP Then
                TryMakeVerticalBeam Center
                StartingX = Center.x
            End If
        End If
    Next Index
End Sub

'===============================================================================
' # Helpers

Private Property Get FindPointForHoleInShape(ByVal Shape As Shape) As Point
    Dim ProbeRadius As Double: ProbeRadius = EdgeSecurity + BOTTOM_HOLE_SIZE / 2
    Dim QuantizedDistance As Double: QuantizedDistance = _
        ClosestDividend(Shape.CenterX - BottomBeam.LeftX, QUANTIZATION_STEP)
    Dim StartingPoint As Point: Set StartingPoint = _
        Point.New_(BottomBeam.LeftX + QuantizedDistance, BottomBeam.CenterY)
    With NextValidPoint( _
            BottomBeam, Shape, StartingPoint, QUANTIZATION_STEP, ProbeRadius, _
            BOTTOM_HOLE_EDGE_SPACE _
        )
        If .Some Then
            Set FindPointForHoleInShape = .Point
            Exit Sub
        Else
            With NextValidPoint( _
                BottomBeam, Shape, StartingPoint, -QUANTIZATION_STEP, ProbeRadius, _
                BOTTOM_HOLE_EDGE_SPACE _
            )
                If .Some Then Set FindPointForHoleInShape = .Point
            End With
        End If
    End With
End Property

Private Sub TryMakeBottomCircle(ByVal Center As Point)
    If Center Is Nothing Then Exit Sub
    MakeCircle _
        Center, BOTTOM_HOLE_SIZE / 2, _
        OutlineColor:=CreateColor(BOTTOM_HOLE_COLOR), _
        Name:=BOTTOM_HOLE_NAME
End Sub

Private Function TryMakeVerticalBeam(ByVal Center As Point) As Shape
    If Center Is Nothing Then Exit Function
    Const HalfThickness As Double = BEAM_THICKNESS / 2
    Set TryMakeVerticalBeam = _
        ActiveLayer.CreateRectangle( _
            Center.x - HalfThickness, TopBeam.TopY, _
            Center.x + HalfThickness, BottomBeam.BottomY _
        )
    ApplyBeamCommonProps TryMakeVerticalBeam
    TryMakeVerticalBeam.Name = VERTICAL_BEAM_NAME
End Function
