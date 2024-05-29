Attribute VB_Name = "Signboard"
'===============================================================================
'   Макрос          : SignboardTest
'   Версия          : 2024.05.29
'   Сайты           : https://vk.com/elvin_macro
'                     https://github.com/elvin-nsk
'   Автор           : elvin-nsk (me@elvin.nsk.ru)
'===============================================================================

Option Explicit

'===============================================================================
' # Manifest

Public Const APP_NAME As String = "SignboardTest"
Public Const APP_DISPLAYNAME As String = APP_NAME
Public Const APP_VERSION As String = "2024.05.29"
Public Const APP_URL As String = "https://vk.com/elvin_macro/" & APP_NAME

'===============================================================================
' # Globals

Private Const GROOVE_SIZE As Double = 3.2
Private Const GROOVE_PUNCH_LENGTH As Double = GROOVE_SIZE * 4

'0.01 - 1, чем больше - тем более вогнутым должен быть угол
'для появления на нём засечек
Private Const CONCAVITY_MULT As Double = 0.65

Private Const HOLES_STEP As Double = 10
Private Const TOP_HOLE_SIZE As Double = 4.2
Private Const TOP_HOLE_STEP As Double = 5
Private Const TOP_HOLE_EDGE_SPACE As Double = HOLES_STEP
Private Const BOTTOM_HOLE_SIZE As Double = 8
Private Const BOTTOM_HOLE_EDGE_SPACE As Double = HOLES_STEP
Private Const PROBE_CIRCLE_MULT As Double = 0.8
'Private Const BEAM_THICKNESS As Double = 20
Private Const BOTTOM_GROOVE_SIZE As Double = 7
Private Const BOTTOM_GROOVE_PUNCH_LENGTH As Double = BOTTOM_GROOVE_SIZE * 3
Private Const BOTTOM_GROOVE_STEP As Double = BOTTOM_GROOVE_SIZE / 3

Private Const PROBE_STEPS As Long = 36

Type Beams
    TopBeam As Shape
    BottomBeam As Shape
    AllBeams As ShapeRange
    Some As Boolean
End Type

Type MaybePoint
    Point As Point
    Some As Boolean
End Type

'===============================================================================
' # Entry points

Sub Start()

    #If DebugMode = 0 Then
    On Error GoTo Catch
    #End If
       
    Dim Shapes As ShapeRange
    With InputData.RequestDocumentOrPage
        If .IsError Then GoTo Finally
        Set Shapes = .Shapes
    End With
    
    Dim Source As ShapeRange
    Set Source = ActiveSelectionRange
    Dim ShapeSize As Double
    ShapeSize = 90
    
    BoostStart "Holes"
    
    MakeHoles Shapes, ShapeSize
    
    Source.CreateSelection
    
Finally:
    BoostFinish
    Exit Sub

Catch:
    VBA.MsgBox VBA.Err.Source & ": " & VBA.Err.Description, vbCritical, "Error"
    Resume Finally

End Sub

'===============================================================================
' # Helpers

Private Sub MakeHoles(ByVal Shapes As ShapeRange, ByVal ShapeSize As Double)
    Dim ShapesToProcess As New ShapeRange
    ShapesToProcess.AddRange Shapes
    Dim Beams As Beams: Beams = FindBeams(Shapes)
    If Not Beams.Some Then Exit Sub
    ShapesToProcess.RemoveRange Beams.AllBeams
    If ShapesToProcess.Count = 0 Then Throw "Нет объектов"
    
    '''
    'Debug.Print Beams.Some
    'Debug.Print ShapesToProcess.Count
    '''
    
    ShapesToProcess.Sort "@Shape1.Left < @Shape2.Left"
    Dim ProbeRadius As Double: ProbeRadius = GetProbeRadius(ShapeSize)
    MakeGrooves ShapesToProcess
    MakeTopHoles Beams.TopBeam, ShapesToProcess, ShapeSize, ProbeRadius
    MakeBottomHoles Beams.BottomBeam, ShapesToProcess, ShapeSize, ProbeRadius
    MakeBottomGrooves ShapesToProcess

End Sub

Private Sub MakeBottomGrooves( _
                ByVal Shapes As ShapeRange _
            )
    Dim Shape As Shape
    For Each Shape In Shapes
        MakeBottomPunch Shape
    Next Shape
End Sub

Private Sub MakeBottomPunch( _
                ByVal Shape As Shape _
            )
    Dim x As Double, x1 As Variant, x2 As Variant
    Dim y As Double: y = Shape.BottomY
    For x = Shape.LeftX To Shape.RightX Step BOTTOM_GROOVE_STEP
        If Shape.Curve.IsOnCurve(x, y, BOTTOM_GROOVE_SIZE) = cdrOnMarginOfShape Then
            If VBA.IsEmpty(x1) Then
                x1 = x
                x2 = x
            Else
                'если далеко убежал, то заканчиваем
                If (x - x2) > BOTTOM_GROOVE_SIZE Then Exit For Else x2 = x
            End If
        End If
    Next x
    x = x1 + (x2 - x1) / 2
    With MakePunch( _
            BOTTOM_GROOVE_PUNCH_LENGTH, BOTTOM_GROOVE_SIZE, _
            OutlineColor:=CreateCMYKColor(0, 100, 100, 0) _
        )
        .Rotate -90
        .CenterX = x
        .TopY = y + BOTTOM_GROOVE_SIZE
    End With
End Sub




Private Sub MakeTopHoles( _
                ByVal Beam As Shape, _
                ByVal ShapesToProcess As ShapeRange, _
                ByVal ShapeSize As Double, _
                ByVal ProbeRadius As Double _
            )
    Dim Shape As Shape
    For Each Shape In ShapesToProcess
        MakeTopHolesInShape Beam, Shape, ProbeRadius
    Next Shape
End Sub

Private Sub MakeTopHolesInShape( _
                ByVal Beam As Shape, _
                ByVal Shape As Shape, _
                ByVal ProbeRadius As Double _
            )
    With NextValidPoint( _
            Beam, Shape, Point.New_(Shape.LeftX + ProbeRadius, Beam.CenterY), _
            HOLES_STEP, ProbeRadius, TOP_HOLE_EDGE_SPACE _
        )
        If .Some Then MakeTopCircle .Point
    End With
    If Shape.SizeWidth > ProbeRadius * 4 Then
        With NextValidPoint( _
                Beam, Shape, Point.New_(Shape.RightX - ProbeRadius, Beam.CenterY), _
                -HOLES_STEP, ProbeRadius, TOP_HOLE_EDGE_SPACE _
            )
            If .Some Then MakeTopCircle .Point
        End With
    End If
End Sub

Private Sub MakeTopCircle(ByVal Center As Point)
    MakeCircle _
        Center, TOP_HOLE_SIZE / 2, _
        OutlineColor:=CreateCMYKColor(100, 0, 0, 0)
End Sub

Private Sub MakeBottomHoles( _
                ByVal Beam As Shape, _
                ByVal ShapesToProcess As ShapeRange, _
                ByVal ShapeSize As Double, _
                ByVal ProbeRadius As Double _
            )
    Dim Shape As Shape
    For Each Shape In ShapesToProcess
        MakeBottomHoleInShape Beam, Shape, ProbeRadius
    Next Shape
End Sub

Private Sub MakeBottomHoleInShape( _
                ByVal Beam As Shape, _
                ByVal Shape As Shape, _
                ByVal ProbeRadius As Double _
            )
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
        OutlineColor:=CreateCMYKColor(100, 0, 100, 0)
End Sub

Private Property Get FindBeams(ByVal Shapes As ShapeRange) As Beams
    Dim Shape As Shape
    With FindBeams
        Set .AllBeams = CreateShapeRange
        For Each Shape In Shapes
            If Shape.Outline.Color.IsSame(CreateCMYKColor(0, 0, 100, 0)) Then
               .AllBeams.Add Shape
            End If
        Next Shape
        If .AllBeams.Count <> 2 Then Exit Property
        If .AllBeams.Shapes(1).TopY > .AllBeams.Shapes(2).TopY Then
            Set .TopBeam = .AllBeams.Shapes(1)
            Set .BottomBeam = .AllBeams.Shapes(2)
        Else
            Set .TopBeam = .AllBeams.Shapes(2)
            Set .BottomBeam = .AllBeams.Shapes(1)
        End If
        .Some = True
    End With
End Property

Private Property Get GetProbeRadius(ByVal ShapeSize As Double) As Double
    GetProbeRadius = ShapeSize * PROBE_CIRCLE_MULT / 2
End Property

'-------------------------------------------------------------------------------

Private Sub MakeGrooves(ByVal Shapes As ShapeRange)
    Dim Shape As Shape
    For Each Shape In Shapes
        MakeGroovesOnNodes Shape.Curve.Nodes.All
    Next Shape
End Sub

Private Sub MakeGroovesOnNodes(ByVal Nodes As NodeRange)
    'MakeVectors Nodes(1)
    Dim Node As Node
    For Each Node In Nodes
        'ActiveLayer.CreateEllipse2 Node.PositionX, Node.PositionY, GROOVE_PUNCH_LENGTH
        'MakeVectors Node
        MakeGroovesOnNode Node
        'MakeVectorsOld Node
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
    'ActiveLayer.CreateLineSegment n.Position.x, n.Position.y, n.ControlPoint1.x, n.ControlPoint1.y
    'ActiveLayer.CreateLineSegment n.Position.x, n.Position.y, n.ControlPoint2.x, n.ControlPoint2.y
    'ActiveLayer.CreateArtisticText n.Position.x, n.Position.y, Node.Type
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

'-------------------------------------------------------------------------------
' # Common

Private Function MakeCircle( _
                     ByVal Center As Point, _
                     ByVal Radius As Double, _
                     Optional ByVal FillColor As Color, _
                     Optional ByVal OutlineColor As Color _
                 ) As Shape
    Set MakeCircle = _
        ActiveLayer.CreateEllipse2(Center.x, Center.y, Radius)
    If IsSome(FillColor) Then MakeCircle.Fill.ApplyUniformFill FillColor
    If IsSome(OutlineColor) Then MakeCircle.Outline.Color.CopyAssign OutlineColor
End Function

Private Function MakePunch( _
                     ByVal Width As Double, _
                     ByVal Diameter As Double, _
                     Optional ByVal FillColor As Color, _
                     Optional ByVal OutlineColor As Color _
                 ) As Shape
    Set MakePunch = _
        ActiveLayer.CreateRectangle2(0, 0, Width, Diameter)
    With MakePunch
        With .Rectangle
            .CornerType = cdrCornerTypeRound
            '.SetRoundness 100 не работает, поэтому так
            .CornerLowerLeft = 100
            .CornerLowerRight = 100
            .CornerUpperLeft = 100
            .CornerUpperRight = 100
        End With
    .RotationCenterX = .LeftX + Diameter / 2
    .RotationCenterY = .BottomY + Diameter / 2
    If IsSome(FillColor) Then .Fill.ApplyUniformFill FillColor
    If IsSome(OutlineColor) Then .Outline.Color.CopyAssign OutlineColor
    End With
End Function

Private Property Get ProbeHits( _
                         ByVal ClosedCurve As Curve, _
                         ByVal Center As Point, _
                         ByVal Radius As Double, _
                         ByVal ProbeSteps As Long _
                     ) As Long
    Dim Step As Double: Step = 360 / ProbeSteps
    Dim Probe As Point: Set Probe = _
        Point.New_(Center.x + Radius, Center.y)
    Dim Angle As Double
    For Angle = Step To 360 Step Step
        Probe.RotateAroundPoint Center, Step
        If PointIsInside(Probe, ClosedCurve) Then ProbeHits = ProbeHits + 1
    Next Angle
End Property

Private Property Get PointIsInside( _
                         ByVal Point As Point, _
                         ByVal Curve As Curve _
                     ) As Boolean
    PointIsInside = Curve.IsPointInside(Point.x, Point.y)
End Property

Private Property Get NextValidPoint( _
                         ByVal Beam As Shape, _
                         ByVal Shape As Shape, _
                         ByVal StartingPoint As Point, _
                         ByVal Step As Double, _
                         ByVal Radius As Double, _
                         ByVal BeamEdgeSpace As Double _
                     ) As MaybePoint
    Const Hits As Long = 36
    Dim LastPoint As Point: Set LastPoint = StartingPoint.GetCopy
    
    Do While LastPoint.x < Shape.RightX _
         And LastPoint.x > Shape.LeftX
         
        If LastPoint.x < Shape.RightX - Radius _
       And LastPoint.x > Shape.LeftX + Radius _
       And LastPoint.x <= Beam.RightX - BeamEdgeSpace _
       And LastPoint.x >= Beam.LeftX + BeamEdgeSpace Then
            If _
                ProbeHits( _
                    Shape.Curve, _
                    LastPoint, Radius, Hits _
                ) = Hits _
            Then
                Set NextValidPoint.Point = LastPoint
                NextValidPoint.Some = True
                Exit Property
            End If
        End If
        LastPoint.Move Step
    Loop
End Property

'===============================================================================
' # Tests

Private Sub testSomething()
    ActiveDocument.Unit = cdrMillimeter
    With Probe(Point.New_(0, 0), 0, GROOVE_PUNCH_LENGTH)
    ActiveLayer.CreateEllipse2(.x, .y, GROOVE_SIZE / 2).Fill.ApplyUniformFill _
        CreateCMYKColor(100, 0, 0, 0)
    End With
End Sub

Private Sub Test2()
    ActiveDocument.Unit = cdrMillimeter
    With Point.New_(10, 0)
        Debug.Print .x, .y
        .RotateAroundPoint Point.New_(0, 0), 45
        Debug.Print .x, .y
    End With
End Sub

'===============================================================================
' # Неактуальное для тестов

Private Sub MakeVectorsOld(ByVal Node As Node)
    On Error Resume Next
    ActiveLayer.CreateLineSegment Node.PositionX, Node.PositionY, Node.Segment.EndingControlPointX, Node.Segment.EndingControlPointY
    ActiveLayer.CreateLineSegment Node.PositionX, Node.PositionY, Node.Segment.Next.StartingControlPointX, Node.Segment.Next.StartingControlPointY
    Dim Angle1 As Double, Angle2 As Double
    Angle1 = Fix(Node.Segment.EndingControlPointAngle)
    Angle2 = Fix(Node.Segment.Next.StartingControlPointAngle)
    ActiveLayer.CreateArtisticText Node.Segment.EndingControlPointX, Node.Segment.EndingControlPointY, Angle1
    ActiveLayer.CreateArtisticText Node.Segment.Next.StartingControlPointX, Node.Segment.Next.StartingControlPointY, Angle2
    On Error GoTo 0
End Sub

Private Sub SetAngleOld(ByVal Node As Node)
    With NodeData.New_(Node)
        Dim ResultAngle As Double: ResultAngle = _
           (.Angle1 + .Angle2) / 2
        Dim ProbePoint1 As Point: Set ProbePoint1 = _
            Probe(.Position, ResultAngle, GROOVE_PUNCH_LENGTH)
            ActiveLayer.CreateEllipse2(ProbePoint1.x, ProbePoint1.y, GROOVE_SIZE / 2).Fill.ApplyUniformFill _
                CreateCMYKColor(100, 0, 0, 0)
            ActiveLayer.CreateArtisticText .Position.x, .Position.y, ResultAngle
            ActiveLayer.CreateArtisticText .ControlPoint1.x, .ControlPoint1.y, Fix(.Angle1)
            ActiveLayer.CreateArtisticText .ControlPoint2.x, .ControlPoint2.y, Fix(.Angle2)
    End With
End Sub

Private Sub MakeBottomHoleInShapeOld( _
                ByVal Beam As Shape, _
                ByVal Shape As Shape, _
                ByVal ProbeRadius As Double, _
                ByVal y As Double, _
                ByRef LastPosition As Double _
            )
    Const Hits As Long = 36
    Do While LastPosition + ProbeRadius < Shape.RightX
        If LastPosition - ProbeRadius > Shape.LeftX _
       And LastPosition > Beam.LeftX _
       And LastPosition <= Beam.RightX - BOTTOM_HOLE_EDGE_SPACE Then
            If _
                ProbeHits( _
                    Shape.Curve, Point.New_(LastPosition, y), ProbeRadius, Hits _
                ) = Hits _
            Then
                MakeCircle _
                    Point.New_(LastPosition, y), _
                    BOTTOM_HOLE_SIZE / 2, _
                    OutlineColor:=CreateCMYKColor(100, 0, 100, 0)
                Exit Sub
            End If
        End If
        LastPosition = LastPosition + HOLES_STEP
    Loop
End Sub
