VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormToJsonBinder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'===============================================================================
'   Модуль          : FormToJsonBinder
'   Версия          : 2024.06.02
'   Автор           : elvin-nsk (me@elvin.nsk.ru)
'   Сайт            : https://github.com/elvin-nsk/LowCoupledFromCore
'   Назначение      : Привязывает контролы формы к json-файлу
'                     для автоматической загрузки/сохранения состояния
'   Зависимости     : LibCore, JsonConverter
'===============================================================================

'@Interface
'@PredeclaredId
Option Explicit

'===============================================================================
' # Declarations

Private Type This
    ControlNames As Collection
    Dictionary As Scripting.IDictionary
    File As String
    Form As MSForms.UserForm
    ReadOnly As Boolean
End Type
Private This As This
Attribute This.VB_VarHelpID = -1

'===============================================================================
' # Constructors

Friend Function New_( _
                    ByVal FileBaseName As String, _
                    ByVal Form As MSForms.UserForm, _
                    ByVal ControlNames As Collection _
                ) As FormToJsonBinder
    Set New_ = New FormToJsonBinder
    New_.Inject FileBaseName, Form, ControlNames, False
End Function

Friend Function NewReadOnly( _
                    ByVal FileBaseName As String, _
                    ByVal Form As MSForms.UserForm, _
                    ByVal ControlNames As Collection _
                ) As FormToJsonBinder
    Set NewReadOnly = New FormToJsonBinder
    NewReadOnly.Inject FileBaseName, Form, ControlNames, True
End Function

Friend Sub Inject( _
               ByVal FileBaseName As String, _
               ByVal Form As MSForms.UserForm, _
               ByVal ControlNames As Collection, _
               ByVal ReadOnly As Boolean _
           )
    With This
        Set .ControlNames = ControlNames
        Set .Form = Form
        .ReadOnly = ReadOnly
        BindJson FileBaseName
        RefreshForm
    End With
End Sub

Private Sub Class_Terminate()
    TryForceSave
End Sub

'===============================================================================
' # Public

Public Property Get Dictionary() As Scripting.Dictionary
    Set Dictionary = This.Dictionary
End Property

Public Sub ForceSave()
    If TryForceSave = False Then Throw "Read-only"
End Sub

'@DefaultMember
Public Property Get Item(ByVal Key As String) As Variant
    Item = This.Dictionary(Key)
End Property
Public Property Let Item(ByVal Key As String, RHS As Variant)
Attribute Item.VB_UserMemId = 0
    This.Dictionary(Key) = RHS
End Property

Public Function RefreshDictionary() As Scripting.Dictionary
    With This
        Dim ItemName As Variant
        For Each ItemName In This.ControlNames
            .Dictionary(ItemName) = .Form.Controls(ItemName)
        Next ItemName
        Set RefreshDictionary = .Dictionary
    End With
End Function

Public Sub RefreshForm()
    With This
        Dim ItemName As Variant
        For Each ItemName In This.ControlNames
            If .Dictionary.Exists(ItemName) Then
                .Form.Controls(ItemName) = .Dictionary(ItemName)
            End If
        Next ItemName
    End With
End Sub

Public Property Get Self() As FormToJsonBinder
    Set Self = Me
End Property

'===============================================================================
' # Helpers

Private Sub BindJson(ByVal FileBaseName As String)
    This.File = GetConfigFile(FileBaseName)
    With JsonConverter.New_
        If FileExists(This.File) Then
            Set This.Dictionary = _
                JsonConverter.New_.ParseJson(ReadFile(This.File))
        Else
            Set This.Dictionary = New Scripting.Dictionary
        End If
    End With
End Sub

Private Function GetConfigFile(ByVal FileBaseName As String) As String
    GetConfigFile = GMSManager.UserGMSPath & FileBaseName & ".json"
End Function

Private Function TryForceSave() As Boolean
    If This.ReadOnly Then Exit Function
    With JsonConverter.New_
        WriteFile .ConvertToJson(This.Dictionary), This.File
    End With
    TryForceSave = True
End Function

'===============================================================================
' # Examples

'Private Sub Example1()
'    Dim Form As New MainView
'    Dim Cfg As FormToJsonBinder
'    Set Cfg = FormToJsonBinder.New_( _
'        FileBaseName:="Example1", _
'        Form:=Form, _
'        ControlNames:=Collection("TextBox1") _
'    )
'    Form.Show
'    Cfg.RefreshDictionary
'    'If Form.IsOk Then
'    'Debug.Print Cfg("TextBox1")
'    'Cfg("TextBox1") = "xxx"
'End Sub
