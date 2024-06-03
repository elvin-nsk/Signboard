Attribute VB_Name = "Signboard"
'===============================================================================
'   Макрос          : Signboard
'   Версия          : 2024.06.03
'   Сайты           : https://vk.com/elvin_macro
'                     https://github.com/elvin-nsk
'   Автор           : elvin-nsk (me@elvin.nsk.ru)
'===============================================================================

Option Explicit

'===============================================================================
' # Manifest

Public Const APP_NAME As String = "Signboard"
Public Const APP_DISPLAYNAME As String = APP_NAME
Public Const APP_VERSION As String = "2024.06.03"
Public Const APP_URL As String = "https://vk.com/elvin_macro/" & APP_NAME

'===============================================================================
' # Globals

Public Const GROOVE_SIZE As Double = 3.2
Public Const GROOVE_PUNCH_LENGTH As Double = GROOVE_SIZE * 4

'0 ... 1, чем больше - тем более вогнутым должен быть угол
'для появления на нём засечек
Public Const CONCAVITY_MULT As Double = 0.6

Public Const HOLES_STEP As Double = 10
Public Const TOP_HOLE_SIZE As Double = 4.2
Public Const TOP_HOLE_STEP As Double = 5
Public Const TOP_HOLE_EDGE_SPACE As Double = HOLES_STEP
Public Const BOTTOM_HOLE_SIZE As Double = 8
Public Const BOTTOM_HOLE_EDGE_SPACE As Double = HOLES_STEP
'Public Const PROBE_CIRCLE_MULT As Double = 0.8
'Public Const BEAM_THICKNESS As Double = 20
Public Const BOTTOM_GROOVE_SIZE As Double = 6
Public Const BOTTOM_GROOVE_PUNCH_LENGTH As Double = BOTTOM_GROOVE_SIZE * 3
Public Const BOTTOM_GROOVE_STEP As Double = BOTTOM_GROOVE_SIZE / 3

Public Const PROBE_STEPS As Long = 36

Public Type Beams
    TopBeam As Shape
    BottomBeam As Shape
    AllBeams As ShapeRange
    Some As Boolean
End Type

Public Type MaybePoint
    Point As Point
    Some As Boolean
End Type

'===============================================================================
' # Entry points

Sub Start()

    #If DebugMode = 0 Then
    On Error GoTo Catch
    #End If
    
    Dim Log As New Logger
       
    Dim Shapes As ShapeRange
    With InputData.RequestDocumentOrPage
        If .IsError Then GoTo Finally
        Set Shapes = .Shapes
    End With
    
    Dim Source As ShapeRange
    Set Source = ActiveSelectionRange
    
    Dim Beams As Beams: Beams = FindBeams(Shapes)
    If Not Beams.Some Then
        Log.Add "Не найдены верняя и/или нижняя части рамы"
        GoTo Finally
    End If
    
    Dim View As New MainView
    Dim Cfg As FormToJsonBinder
    Set Cfg = BindConfig(View)
    View.Show vbModal
    Cfg.RefreshDictionary
    If View.IsCancel Then GoTo Finally
    
    BoostStart APP_DISPLAYNAME
    
    MakeElements Shapes, Beams, Cfg, Log
    
    Source.CreateSelection
    
Finally:
    BoostFinish
    If IsSome(Log) Then Log.Check
    Exit Sub

Catch:
    VBA.MsgBox VBA.Err.Source & ": " & VBA.Err.Description, vbCritical, "Error"
    Resume Finally

End Sub

'===============================================================================
' # Helpers

Private Function BindConfig(ByVal View As MSForms.UserForm) As FormToJsonBinder
    Set BindConfig = FormToJsonBinder.New_( _
        FileBaseName:="elvin_" & APP_NAME, _
        Form:=View, _
        ControlNames:=Collection( _
            "MinEdgeSecurity" _
        ) _
    )
End Function

Private Sub MakeElements( _
                ByVal AllShapes As ShapeRange, _
                ByRef Beams As Beams, _
                ByVal Cfg As FormToJsonBinder, _
                ByVal Log As Logger _
            )
    Dim ShapesToProcess As New ShapeRange
    ShapesToProcess.AddRange AllShapes
    ShapesToProcess.RemoveRange Beams.AllBeams
    If ShapesToProcess.Count = 0 Then Throw "Нет объектов"
    
    ShapesToProcess.Sort "@Shape1.Left < @Shape2.Left"
    With New InnerGrooves
        Set .Shapes = ShapesToProcess
        .MakeGrooves
    End With
    With New TopHoles
        Set .Beam = Beams.TopBeam
        Set .Shapes = ShapesToProcess
        .EdgeSecurity = Cfg("MinEdgeSecurity")
        .MakeTopHoles
    End With
    With New BottomHoles
        Set .Beam = Beams.BottomBeam
        Set .Shapes = ShapesToProcess
        .EdgeSecurity = Cfg("MinEdgeSecurity")
        .MakeBottomHoles
    End With
    With New BottomGrooves
        Set .Shapes = ShapesToProcess
        .MakeBottomGrooves
    End With

End Sub

Private Property Get FindBeams(ByVal Shapes As ShapeRange) As Beams
    Dim Shape As Shape
    With FindBeams
        Set .AllBeams = CreateShapeRange
        For Each Shape In Shapes
            If Shape.Outline.Color.IsSame(CreateCMYKColor(100, 100, 0, 0)) Then
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

'Private Property Get GetProbeRadius(ByVal ShapeThickness As Double) As Double
' GetProbeRadius = ShapeThickness * PROBE_CIRCLE_MULT / 2
'End Property

'===============================================================================
' # Common

Public Function MakeCircle( _
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

Public Function MakePunch( _
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

'поиск ближайшей окружности внутри Shape
'с центром внутри Beam (не ближе BeamEdgeSpace от края),
'с радиусом Radius, с шагом Step
'Step > 0 - вправо от StartingPoint
'Step < 0 - влево
Public Property Get NextValidPoint( _
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

Public Property Get ProbeHits( _
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
        If Probe.Inside(ClosedCurve) Then ProbeHits = ProbeHits + 1
    Next Angle
End Property

'===============================================================================
' # Tests

Private Sub testSomething()
    Dim Form As New MainView
    Dim Cfg As FormToJsonBinder
    Set Cfg = FormToJsonBinder.New_( _
        FileBaseName:="test", _
        Form:=Form, _
        ControlNames:=Collection("TextBox1") _
    )
    Form.Show
    Cfg.RefreshDictionary
    'If Form.IsOk Then
    'Debug.Print Cfg("TextBox1")
    'Cfg("TextBox1") = "xxx"
End Sub
