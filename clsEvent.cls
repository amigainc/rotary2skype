VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private oDecroche As clsEvent

Private oRaccroche As clsEvent

Private oGet1 As clsEvent

Private oGet2to0 As clsEvent

Private oAction As Object

Public Function executeEvent(sEvent As String) As clsEvent

    Select Case sEvent
        Case "1"
            executeEvent = oGet1
        Case "U"
            executeEvent = oDecroche
        Case "D"
            executeEvent = oRaccroche
        Case Else
            executeEvent = oGet2to0
    End Select

End Function


Public Sub executeAction()

    Eval oAction

End Sub

Public Sub Init(ByRef oDec As clsEvent, _
    ByRef oRac As clsEvent, _
    ByRef o1 As clsEvent, _
    ByRef o2 As clsEvent, _
    Optional ByRef action As Object)

    Set oDecroche = oDec
    Set oRaccroche = oRac
    Set oGet1 = o1
    Set oGet2to0 = o2
    Set oAction = action

End Sub
