Attribute VB_Name = "Signboard"
'===============================================================================
'   ������          : Signboard
'   ������          : 2024.06.26
'   �����           : https://vk.com/elvin_macro
'                     https://github.com/elvin-nsk
'   �����           : elvin-nsk (me@elvin.nsk.ru)
'===============================================================================

Option Explicit

'===============================================================================
' # Manifest

Public Const APP_NAME As String = "Signboard"
Public Const APP_DISPLAYNAME As String = APP_NAME
Public Const APP_VERSION As String = "2024.06.26"

'===============================================================================
' # Globals

Public Const QUANTIZATION_STEP As Double = 10

Public Const GROOVE_SIZE As Double = 3.2
Public Const GROOVE_PUNCH_LENGTH As Double = GROOVE_SIZE * 4
Public Const GROOVE_COLOR As String = "CMYK,USER,0,100,0,0"
Public Const GROOVE_NAME As String = "INNER_PUNCH"

'0 ... 1, ��� ������ - ��� ����� �������� ������ ���� ����
'��� ��������� �� �� �������
Public Const CONCAVITY_MULT As Double = 0.6
Public Const PROBE_RADIUS As Double = GROOVE_SIZE / 10

Public Const TOP_HOLE_SIZE As Double = 4.2
Public Const TOP_HOLE_STEP As Double = 5
Public Const TOP_HOLE_EDGE_SPACE As Double = QUANTIZATION_STEP
Public Const TOP_HOLE_COLOR As String = "CMYK,USER,100,0,0,0"
Public Const TOP_HOLE_NAME As String = "TOP_HOLE"
Public Const BOTTOM_HOLE_SIZE As Double = 8
Public Const BOTTOM_HOLE_EDGE_SPACE As Double = QUANTIZATION_STEP
Public Const BOTTOM_HOLE_COLOR As String = TOP_HOLE_COLOR
Public Const BOTTOM_HOLE_NAME As String = "BOTTOM_HOLE"
'Public Const PROBE_CIRCLE_MULT As Double = 0.8
Public Const BEAM_THICKNESS As Double = 20
Public Const VERTICAL_BEAM_NAME As String = "V_BEAM"
Public Const HORIZONTAL_BEAM_NAME As String = "H_BEAM"
Public Const VERTICAL_BEAM_STEP As Double = 1000
Public Const HOLES_DICTIONARY_NAME As String = "Holes"

Public Const BOTTOM_GROOVE_SIZE As Double = 6
Public Const BOTTOM_GROOVE_PUNCH_LENGTH As Double = BOTTOM_GROOVE_SIZE * 3
Public Const BOTTOM_GROOVE_STEP As Double = BOTTOM_GROOVE_SIZE / 3
Public Const BOTTOM_GROOVE_COLOR As String = GROOVE_COLOR
Public Const BOTTOM_GROOVE_NAME As String = "BOTTOM_PUNCH"

Public Const PROBE_STEPS As Long = 36

Public Const FACE_COLOR As String = "CMYK,USER,0,100,100,0"
Public Const FACE_NAME As String = "����"
Public Const BACK_COLOR As String = "CMYK,USER,100,0,100,0"
Public Const BACK_NAME As String = "������"
Public Const BACK_CONTOUR_INT As Double = 0.8
Public Const BACK_CONTOUR_INT_NAME As String = "INT_CONTOUR"
Public Const BACK_CONTOUR_EXT As Double = 0.8
Public Const BACK_CONTOUR_EXT_NAME As String = "EXT_CONTOUR"

Public Const DIMENSION_OFFSET_MULT As Double = 0.1
Public Const DIMENSION_TEXT_SIZE_MULT As Double = 0.5
Public Const DIMENSION_SHAPES_COLOR As String = "CMYK,USER,0,0,0,100"
Public Const DIMENSIONS_NAME As String = "���� � ���������"

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

Sub Part1__PrepareSelected()
    #If DebugMode = 0 Then
    On Error GoTo Catch
    #End If
    
    Dim Log As New Logger
       
    Dim Shapes As ShapeRange
    With InputData.RequestShapes
        If .IsError Then GoTo Finally
        Set Shapes = .Shapes
    End With
    
    Dim Source As ShapeRange
    Set Source = ActiveSelectionRange
    
    If Not CheckShapesHasCurves(Shapes, Log) Then GoTo Finally
    
    Dim Cfg As Dictionary
    If Not ShowPreparationsView(Cfg) Then GoTo Finally
        
    Shapes.CreateDocumentFrom.Activate
    BoostStart "���������� ������� �����"
    ProcessFaceDoc
    BoostFinish
    
    Shapes.CreateDocumentFrom.Activate
    BoostStart "���������� �������"
    With New BackDoc
        .BeamEdgeOffset = Cfg("BeamEdgeOffset")
        .ProcessBackDoc
    End With
    BoostFinish
    
    Source.CreateSelection
    
Finally:
    CheckLog Log
    Exit Sub

Catch:
    VBA.MsgBox VBA.Err.Source & ": " & VBA.Err.Description, vbCritical, "Error"
    Resume Finally
End Sub

Sub Part2__MakeHoles()

    #If DebugMode = 0 Then
    On Error GoTo Catch
    #End If
    
    Dim Log As New Logger
       
    Dim Shapes As ShapeRange
    With InputData.RequestDocumentOrPage
        If .IsError Then GoTo Finally
        Set Shapes = .Shapes
    End With
    
    Dim Beams As Beams: Beams = FindBeams(Shapes)
    If Not Beams.Some Then Log.Add "�� ������� ������ �/��� ������ ����� ����"
    Dim ShapesForHoles As ShapeRange: Set ShapesForHoles = _
        FindShapesByName(Shapes, BACK_CONTOUR_INT_NAME)
    Set ShapesForHoles = FindShapesNotInside(ShapesForHoles)
    If ShapesForHoles.Count = 0 Then Log.Add "�� ������� ��������� ��� ���������"
    If Log.Count > 0 Then GoTo Finally
    
    Dim Cfg As Dictionary
    If Not ShowHolesView(Cfg) Then GoTo Finally
    
    BoostStart "��������� � ������������ ���������"
    
    MakeHoles ShapesForHoles, Beams, Cfg, Log
    
Finally:
    BoostFinish
    CheckLog Log
    Exit Sub

Catch:
    VBA.MsgBox VBA.Err.Source & ": " & VBA.Err.Description, vbCritical, "Error"
    Resume Finally

End Sub

Sub Part3__Finalize()

    #If DebugMode = 0 Then
    On Error GoTo Catch
    #End If
    
    Dim Log As New Logger
       
    Dim Shapes As ShapeRange
    With InputData.RequestDocumentOrPage
        If .IsError Then GoTo Finally
        Set Shapes = .Shapes
    End With
    
    Dim BottomHoles As ShapeRange: Set BottomHoles = _
        FindShapesByName(Shapes, BOTTOM_HOLE_NAME)
        
    Dim VerticalBeams As ShapeRange: Set VerticalBeams = _
        FindShapesByName(Shapes, VERTICAL_BEAM_NAME)

    Dim IntPunches As New ShapeRange
    IntPunches.AddRange FindShapesByName(Shapes, TOP_HOLE_NAME)
    IntPunches.AddRange BottomHoles
    IntPunches.AddRange FindShapesByName(Shapes, GROOVE_NAME)
    
    Dim ExtPunches As New ShapeRange: Set ExtPunches = _
        FindShapesByName(Shapes, BOTTOM_GROOVE_NAME)
        
    If IntPunches.Count = 0 _
   And ExtPunches.Count = 0 Then
        Log.Add "��� ��������� ��� ���������"
    End If
    
    Dim IntShapes As ShapeRange: Set IntShapes = _
        FindShapesByName(Shapes, BACK_CONTOUR_INT_NAME)
    Dim ExtShapes As ShapeRange: Set ExtShapes = _
        FindShapesByName(Shapes, BACK_CONTOUR_EXT_NAME)
        
    If IntShapes.Count = 0 _
   And ExtShapes.Count = 0 Then
        Log.Add "�� ������� ��������� (����), � ������� ������ ������������� ���������"
    End If
    
    If Log.Count > 0 Then GoTo Finally
        
    Dim BackDoc As Document: Set BackDoc = ActiveDocument
    BoostStart "���������"
    
    Dim BottomHolesDup As ShapeRange: Set BottomHolesDup = BottomHoles.Duplicate
    
    If IntPunches.Count > 0 And IntShapes.Count > 0 Then _
        CutShapes IntPunches, IntShapes
    If ExtPunches.Count > 0 And ExtShapes.Count > 0 Then _
        CutShapes ExtPunches, ExtShapes
        
    Set Shapes = ActivePage.Shapes.All
        
    Dim DimensionShapesToDelete As New ShapeRange
    DimensionShapesToDelete.AddRange BottomHolesDup
    DimensionShapesToDelete.AddRange FindShapesByName(Shapes, VERTICAL_BEAM_NAME)
    DimensionShapesToDelete.AddRange FindShapesByName(Shapes, HORIZONTAL_BEAM_NAME)
    
    Dim DimensionShapes As New ShapeRange
    DimensionShapes.AddRange DimensionShapesToDelete
    DimensionShapes.AddRange FindShapesByName(Shapes, BACK_CONTOUR_INT_NAME)
    DimensionShapes.AddRange FindShapesByName(Shapes, BACK_CONTOUR_EXT_NAME)
            
    Dim DimensionsDoc As Document
    Set DimensionsDoc = DimensionShapes.CreateDocumentFrom
    DimensionsDoc.Name = DIMENSIONS_NAME
    
    DimensionsDoc.Activate
    BoostStart "����������� ��������"
    
    With New DimensionsMaker
        .MakeDimensions
    End With
    
    BoostFinish
    
    BackDoc.Activate
    DimensionShapesToDelete.Delete
    BoostFinish
    
    DimensionsDoc.Activate
    
Finally:
    CheckLog Log
    Exit Sub

Catch:
    VBA.MsgBox VBA.Err.Source & ": " & VBA.Err.Description, vbCritical, "Error"
    Resume Finally

End Sub

'===============================================================================
' # Helpers

Private Function CheckSource( _
                     ByVal Shapes As ShapeRange, _
                     ByVal Log As Logger _
                 ) As Boolean
End Function

Private Function CheckShapesHasCurves( _
                     ByVal Shapes As ShapeRange, _
                     ByVal Log As Logger _
                 ) As Boolean
    CheckShapesHasCurves = True
    Dim Shape As Shape
    For Each Shape In Shapes
        If Not HasCurve(Shape) Then
            Log.Add "������ �� � ������", Shape
            CheckShapesHasCurves = False
        End If
    Next Shape
End Function

Private Sub ProcessFaceDoc()
    With ActiveDocument
        .Unit = cdrMillimeter
        ResizePageToShapes SideAdd:=AverageDim(ActivePage.Shapes.All) / 10
        ActivePage.Shapes.All.ApplyNoFill
        SetOutlineColor ActivePage.Shapes.All, CreateColor(FACE_COLOR)
        ActivePage.Shapes.All.Flip cdrFlipHorizontal
        .Name = FACE_NAME
    End With
End Sub

'-------------------------------------------------------------------------------

Private Function ShowPreparationsView(ByRef Cfg As Dictionary) As Boolean
    Dim FileBinder As JsonFileBinder: Set FileBinder = BindConfig
    Set Cfg = FileBinder.GetOrMakeSubDictionary("Preparations")
    Dim View As New PreparationsView
    Dim ViewBinder As ViewToDictionaryBinder: Set ViewBinder = _
        ViewToDictionaryBinder.New_( _
            Dictionary:=Cfg, _
            View:=View, _
            ControlNames:=Pack("BeamEdgeOffset") _
        )
    View.Show vbModal
    ViewBinder.RefreshDictionary
    ShowPreparationsView = View.IsOk
End Function

Private Function ShowHolesView(ByRef Cfg As Dictionary) As Boolean
    Dim FileBinder As JsonFileBinder: Set FileBinder = BindConfig
    Set Cfg = FileBinder.GetOrMakeSubDictionary("Holes")
    Dim View As New HolesView
    Dim ViewBinder As ViewToDictionaryBinder: Set ViewBinder = _
        ViewToDictionaryBinder.New_( _
            Dictionary:=Cfg, _
            View:=View, _
            ControlNames:=Pack("MinEdgeSecurity") _
        )
    View.Show vbModal
    ViewBinder.RefreshDictionary
    ShowHolesView = View.IsOk
End Function

Private Sub MakeHoles( _
                ByVal Shapes As ShapeRange, _
                ByRef Beams As Beams, _
                ByVal Cfg As Dictionary, _
                ByVal Log As Logger _
            )
    Shapes.Sort "@Shape1.Left < @Shape2.Left"
    
    With New TopHoles
        Set .Beam = Beams.TopBeam
        Set .Shapes = Shapes
        .EdgeSecurity = Cfg("MinEdgeSecurity")
        .MakeTopHoles
    End With
    With New BottomHoles
        Set .TopBeam = Beams.TopBeam
        Set .BottomBeam = Beams.BottomBeam
        Set .Shapes = Shapes
        .EdgeSecurity = Cfg("MinEdgeSecurity")
        .MakeBottomHoles
    End With

End Sub

'-------------------------------------------------------------------------------

Private Function CutShapes( _
                     ByVal Punches As ShapeRange, _
                     ByVal Shapes As ShapeRange _
                 )
    Punches.Combine.Trim Shapes.Combine, False, False
End Function

'-------------------------------------------------------------------------------

Private Sub CheckLog(ByVal Log As Logger)
    If IsSome(Log) Then Log.Check
End Sub

'Private Property Get GetProbeRadius(ByVal ShapeThickness As Double) As Double
' GetProbeRadius = ShapeThickness * PROBE_CIRCLE_MULT / 2
'End Property

'===============================================================================
' # Common

Public Property Get FindBeams(ByVal Shapes As ShapeRange) As Beams
    Dim Shape As Shape
    With FindBeams
        Set .AllBeams = CreateShapeRange
        For Each Shape In Shapes
            If Shape.Name = HORIZONTAL_BEAM_NAME Then
               .AllBeams.Add Shape
            End If
        Next Shape
        If .AllBeams.Count <> 2 Then
            Exit Property
        End If
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

Public Function MakeCircle( _
                    ByVal Center As Point, _
                    ByVal Radius As Double, _
                    Optional ByVal FillColor As Color, _
                    Optional ByVal OutlineColor As Color, _
                    Optional ByVal Name As String _
                ) As Shape
    Set MakeCircle = _
        ActiveLayer.CreateEllipse2(Center.x, Center.y, Radius)
    If IsSome(FillColor) Then MakeCircle.Fill.ApplyUniformFill FillColor
    If IsSome(OutlineColor) Then MakeCircle.Outline.Color.CopyAssign OutlineColor
    If Not Name = vbNullString Then MakeCircle.Name = Name
End Function

Public Function MakePunch( _
                    ByVal Width As Double, _
                    ByVal Diameter As Double, _
                    Optional ByVal FillColor As Color, _
                    Optional ByVal OutlineColor As Color, _
                    Optional ByVal Name As String _
                ) As Shape
    Set MakePunch = _
        ActiveLayer.CreateRectangle2(0, 0, Width, Diameter)
    With MakePunch
        With .Rectangle
            .CornerType = cdrCornerTypeRound
            '.SetRoundness 100 �� ��������, ������� ���
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
    If Not Name = vbNullString Then MakePunch.Name = Name
End Function

'����� ��������� ���������� ������ Shape
'� ������� ������ Beam (�� ����� BeamEdgeSpace �� ����),
'� �������� Radius, � ����� Step
'Step > 0 - ������ �� StartingPoint
'Step < 0 - �����
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

Public Sub ApplyBeamCommonProps(ByVal Beam As Shape)
    Beam.Fill.ApplyNoFill
End Sub

Private Function BindConfig() As JsonFileBinder
    Set BindConfig = JsonFileBinder.New_("elvin_" & APP_NAME)
End Function

'===============================================================================
' # Tests

Private Sub testDividend()
    Show ClosestDividend(671, 10) '670
End Sub

Private Sub testSnapPoints()
    With ActivePage.Shapes.First
        Dim Index As Long
        Dim Point As SnapPoint
        For Each Point In .SnapPointsOfType(cdrSnapPointBBox)
            Index = Index + 1
            ActiveLayer.CreateArtisticText _
                Point.PositionX, Point.PositionY, Index
        Next Point
    End With
End Sub
