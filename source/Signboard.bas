Attribute VB_Name = "Signboard"
'===============================================================================
'   Макрос          : SignboardTest
'   Версия          : 2024.01.01
'   Сайты           : https://vk.com/elvin_macro
'                     https://github.com/elvin-nsk
'   Автор           : elvin-nsk (me@elvin.nsk.ru)
'===============================================================================

Option Explicit

Public Const RELEASE As Boolean = False

'===============================================================================
' # Manifest

Public Const APP_NAME As String = "SignboardTest"
Public Const APP_DISPLAYNAME As String = APP_NAME
Public Const APP_VERSION As String = "2024.01.01"
Public Const APP_URL As String = "https://vk.com/elvin_macro/" & APP_NAME

'===============================================================================
' # Globals

Private Const TOP_HOLE_SIZE As Double = 4.2
Private Const BOTTOM_HOLE_SIZE As Double = 8
Private Const GROOVE_SIZE As Double = 3.2
Private Const GROOVE_PROBE_LENGTH As Double = GROOVE_SIZE * 4
Private Const BEAM_THICKNESS As Double = 20
Private Const CONVEXITY_FACTOR As Double = 0.65


'===============================================================================
' # Entry points

Sub Grooves()

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
    
    BoostStart "Grooves"
    
    MakeGrooves Shapes
    
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
        'ActiveLayer.CreateEllipse2 Node.PositionX, Node.PositionY, GROOVE_PROBE_LENGTH
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
        .RotationCenterX = Node.PositionX
        .RotationCenterY = Node.PositionY
        .Rotate AngleOutside(Node)
    End With
End Sub

Private Function MakeGroovePunch() As Shape
    Dim Result As Shape
    Set Result = _
        ActiveLayer.CreateRectangle2(0, 0, GROOVE_PROBE_LENGTH, GROOVE_SIZE)
    With Result.Rectangle
        .CornerType = cdrCornerTypeRound
        '.SetRoundness 100 не работает, поэтому так
        .CornerLowerLeft = 100
        .CornerLowerRight = 100
        .CornerUpperLeft = 100
        .CornerUpperRight = 100
    End With
    Result.Outline.Color.CMYKAssign 0, 100, 0, 0
    Set MakeGroovePunch = Result
End Function

Private Property Get AngleOutside(ByVal Node As Node) As Double
    With NodeData.New_(Node)
        AngleOutside = (.Angle1 + .Angle2) / 2
        Dim ProbePoint As Point: Set ProbePoint = _
            Probe(.Position, AngleOutside, GROOVE_PROBE_LENGTH)
        If Not Node.Parent.IsPointInside(ProbePoint.x, ProbePoint.y) Then _
            Exit Property
        AngleOutside = AngleOutside + 180
    End With
End Property

Private Property Get IsNodeConvex(ByVal Node As Node) As Boolean
    Const Step As Long = 10
    Const MaxHits As Long = 36
    Dim Probe As Point: Set Probe = _
        Point.New_(Node.PositionX + GROOVE_PROBE_LENGTH, Node.PositionY)
    Dim Pivot As Point: Set Pivot = Point.New_(Node.PositionX, Node.PositionY)
    Dim Hits As Long
    Dim a As Double
    For a = Step To MaxHits * Step Step Step 'всего MaxHits итераций
        Probe.RotateAroundPoint Pivot, Step
        If PointIsInside(Probe, Node.Parent) Then Hits = Hits + 1
    Next a
    If Hits < MaxHits * CONVEXITY_FACTOR Then IsNodeConvex = True
End Property

Private Property Get PointIsInside( _
                         ByVal Point As Point, _
                         ByVal Curve As Curve _
                     ) As Boolean
    PointIsInside = Curve.IsPointInside(Point.x, Point.y)
End Property

Private Sub MakeVectors(ByVal Node As Node)
    Dim n As NodeData: Set n = NodeData.New_(Node)
    'ActiveLayer.CreateLineSegment n.Position.x, n.Position.y, n.ControlPoint1.x, n.ControlPoint1.y
    'ActiveLayer.CreateLineSegment n.Position.x, n.Position.y, n.ControlPoint2.x, n.ControlPoint2.y
    'ActiveLayer.CreateArtisticText n.Position.x, n.Position.y, Node.Type
    With Probe(n.Position, 120, GROOVE_PROBE_LENGTH)
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


'===============================================================================
' # Tests

Private Sub testSomething()
    ActiveDocument.Unit = cdrMillimeter
    With Probe(Point.New_(0, 0), 0, GROOVE_PROBE_LENGTH)
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
            Probe(.Position, ResultAngle, GROOVE_PROBE_LENGTH)
            ActiveLayer.CreateEllipse2(ProbePoint1.x, ProbePoint1.y, GROOVE_SIZE / 2).Fill.ApplyUniformFill _
                CreateCMYKColor(100, 0, 0, 0)
            ActiveLayer.CreateArtisticText .Position.x, .Position.y, ResultAngle
            ActiveLayer.CreateArtisticText .ControlPoint1.x, .ControlPoint1.y, Fix(.Angle1)
            ActiveLayer.CreateArtisticText .ControlPoint2.x, .ControlPoint2.y, Fix(.Angle2)
    End With
End Sub
