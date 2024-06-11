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

Public Beam As Shape
Public Shapes As ShapeRange
Public EdgeSecurity As Double

'===============================================================================
' # Public

Public Sub MakeBottomHoles()
    Dim Shape As Shape
    For Each Shape In Shapes
        MakeBottomHoleInShape Shape
    Next Shape
End Sub

'===============================================================================
' # Helpers

Private Sub MakeBottomHoleInShape(ByVal Shape As Shape)
    Dim ProbeRadius As Double: ProbeRadius = EdgeSecurity + BOTTOM_HOLE_SIZE / 2
    Dim StartingPoint As Point: Set StartingPoint = _
        Point.New_(ClosestDividend(Shape.CenterX, HOLES_STEP), Beam.CenterY)
    With NextValidPoint( _
            Beam, Shape, StartingPoint, HOLES_STEP, ProbeRadius, _
            BOTTOM_HOLE_EDGE_SPACE _
        )
        If .Some Then
            MakeBottomCircle .Point
            Exit Sub
        Else
            With NextValidPoint( _
                Beam, Shape, StartingPoint, -HOLES_STEP, ProbeRadius, _
                BOTTOM_HOLE_EDGE_SPACE _
            )
                If .Some Then MakeBottomCircle .Point
            End With
        End If
    End With
End Sub

Private Sub MakeBottomCircle(ByVal Center As Point)
    MakeCircle _
        Center, BOTTOM_HOLE_SIZE / 2, _
        OutlineColor:=CreateColor(BOTTOM_HOLE_COLOR), _
        Name:=BOTTOM_HOLE_NAME
End Sub
