VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "General"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'=================
' local constants
'=================
' table constants

'[DANO]
Const CELL_FA_ANCHOR = "B2"

'// Xp methods
'[DANO]
Public Sub pasteFAarray()
''    If IsEmpty(Api.arrFA) Then Debug.Print "api.arrFA is Empty - connect?": Stop
    Range(CELL_FA_ANCHOR).Resize(UBound(Api.arrFA), 1).Value = Application.Transpose(Api.arrFA)
End Sub
