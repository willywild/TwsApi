Attribute VB_Name = "Main"
Option Explicit

Public db As cXpCapitalDb
Public Api As Api


Public Sub Initialise()
''    Stop
    If Not Api Is Nothing Then Set Api = Nothing
    If Not db Is Nothing Then Set db = Nothing
    If Not MsgBox("Initialize api?", vbYesNo, "CSharpTWSapi Class module") = vbYes Then Exit Sub
    Set Api = New Api
    Set db = New cXpCapitalDb
''    If Not MsgBox("Connect Tws?", vbYesNo, "CSharpTWSapi Class module") = vbYes Then Exit Sub
''    Api.Connect
''    If Not MsgBox("Load / Paste Accounts?", vbYesNo, "CSharpTWSapi Class module") = vbYes Then Exit Sub
''    Api.loadAccounts
''    Api.PasteRecordSets
End Sub
