VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} HolesView 
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5280
   OleObjectBlob   =   "HolesView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "HolesView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'===============================================================================
' # State

Public IsOk As Boolean
Public IsCancel As Boolean

Private MinEdgeSecurityHandler As TextBoxHandler

'===============================================================================
' # Constructor

Private Sub UserForm_Initialize()
    Caption = APP_DISPLAYNAME & " (v" & APP_VERSION & ")"
    btnOk.Default = True
End Sub

'===============================================================================
' # Handlers

Private Sub UserForm_Activate()
    Set MinEdgeSecurityHandler = _
        TextBoxHandler.New_(MinEdgeSecurity, TextBoxTypeDouble, 0)
End Sub

Private Sub btnOk_Click()
    FormŒ 
End Sub

Private Sub btnCancel_Click()
    FormCancel
End Sub

'===============================================================================
' # Logic

Private Sub FormŒ ()
    Hide
    IsOk = True
End Sub

Private Sub FormCancel()
    Hide
    IsCancel = True
End Sub

'===============================================================================
' # Helpers



'===============================================================================
' # Boilerplate

Private Sub UserForm_QueryClose(—ancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        —ancel = True
        FormCancel
    End If
End Sub
