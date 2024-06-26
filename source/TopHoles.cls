VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TopHoles"
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

Public Sub MakeTopHoles()
    Dim Shape As Shape
    For Each Shape In Shapes
        MakeTopHolesInShape Shape
    Next Shape
End Sub

'===============================================================================
' # Helpers

Private Sub MakeTopHolesInShape( _
                ByVal Shape As Shape _
            )
    Dim ProbeRadius As Double: ProbeRadius = EdgeSecurity + TOP_HOLE_SIZE / 2
    With NextValidPoint( _
            Beam, Shape, Point.New_(Shape.LeftX + ProbeRadius, Beam.CenterY), _
            QUANTIZATION_STEP, ProbeRadius, TOP_HOLE_EDGE_SPACE _
        )
        If .Some Then MakeTopCircle .Point
    End With
    If Shape.SizeWidth > ProbeRadius * 4 Then
        With NextValidPoint( _
                Beam, Shape, Point.New_(Shape.RightX - ProbeRadius, Beam.CenterY), _
                -QUANTIZATION_STEP, ProbeRadius, TOP_HOLE_EDGE_SPACE _
            )
            If .Some Then MakeTopCircle .Point
        End With
    End If
End Sub

Private Sub MakeTopCircle(ByVal Center As Point)
    MakeCircle _
        Center, TOP_HOLE_SIZE / 2, _
        OutlineColor:=CreateColor(TOP_HOLE_COLOR), _
        Name:=TOP_HOLE_NAME
End Sub

