VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InnerGrooves"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'===============================================================================
' # Fields

Public Shapes As ShapeRange

'===============================================================================
' # Public

Public Sub MakeGrooves()
    Dim Shape As Shape
    For Each Shape In Shapes
        MakeGroovesOnNodes Shape.Curve.Nodes.All
    Next Shape
End Sub

'===============================================================================
' # Helpers

Private Sub MakeGroovesOnNodes(ByVal Nodes As NodeRange)
    Dim Node As Node
    For Each Node In Nodes
        MakeGroovesOnNode Node
    Next Node
End Sub

Private Sub MakeGroovesOnNode(ByVal Node As Node)
    
    If IsNodeConvex(Node) Then Exit Sub
    
    Dim Angle As Double: Angle = AngleOutside(Node)
    Dim Punch As Shape: Set Punch = MakeGroovePunch
    With Punch
        .LeftX = Node.PositionX - GROOVE_SIZE / 2
        .TopY = Node.PositionY + GROOVE_SIZE / 2
        .Rotate AngleOutside(Node)
    End With
End Sub

Private Function MakeGroovePunch() As Shape
    Set MakeGroovePunch = _
        MakePunch( _
            GROOVE_PUNCH_LENGTH, GROOVE_SIZE, _
            OutlineColor:=CreateCMYKColor(0, 100, 0, 0) _
        )
End Function

Private Property Get AngleOutside(ByVal Node As Node) As Double
    With NodeData.New_(Node)
        AngleOutside = (.Angle1 + .Angle2) / 2
        Dim ProbePoint As Point: Set ProbePoint = _
            Probe(.Position, AngleOutside, GROOVE_PUNCH_LENGTH)
        If Not Node.Parent.IsPointInside(ProbePoint.x, ProbePoint.y) Then _
            Exit Property
        AngleOutside = AngleOutside + 180
    End With
End Property

Private Property Get IsNodeConvex(ByVal Node As Node) As Boolean
    Const MaxHits As Long = PROBE_STEPS
    Dim Hits As Long
    Hits = _
        ProbeHits( _
            Node.Parent, _
            Point.New_(Node.PositionX, Node.PositionY), _
            GROOVE_SIZE / 2, _
            MaxHits _
        )
    If Hits < MaxHits * CONCAVITY_MULT Then IsNodeConvex = True
End Property

Private Sub MakeVectors(ByVal Node As Node)
    Dim n As NodeData: Set n = NodeData.New_(Node)
    With Probe(n.Position, 120, GROOVE_PUNCH_LENGTH)
        ActiveLayer.CreateEllipse2(.x, .y, GROOVE_SIZE / 2).Fill.ApplyUniformFill _
            CreateCMYKColor(100, 0, 0, 0)
    End With
End Sub

Private Property Get Probe( _
                         ByVal StartingPoint As Point, _
                         ByVal Angle As Double, _
                         ByVal Length As Double _
                     ) As Point
    Set Probe = StartingPoint.GetCopy
    Probe.Move Length
    Probe.RotateAroundPoint StartingPoint, Angle
End Property
