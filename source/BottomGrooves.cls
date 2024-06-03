VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BottomGrooves"
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

Public Sub MakeBottomGrooves()
    Dim Shape As Shape
    For Each Shape In Shapes
        MakeBottomPunch Shape
    Next Shape
End Sub

'===============================================================================
' # Helpers

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

