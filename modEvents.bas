Attribute VB_Name = "modEvents"
Option Explicit

Public oCurrent As clsEvent

Public oInit As New clsEvent
Public oSearchName As New clsEvent
Public oCompose As New clsEvent
Public oCallReceived As New clsEvent
Public oConversation As New clsEvent
Public oBascule As New clsEvent
Public oConference As New clsEvent
Public oSetWebcam As New clsEvent
Public oUnsetWebcam As New clsEvent
Public oReadMessages As New clsEvent
Public oEndConversation As New clsEvent

Sub initClasses()

    oInit.Init oCompose, Nothing, oSearchName, oSearchName
    oSearchName.Init oConversation, Nothing, oSearchName, oSearchName
    oConversation.Init Nothing, oEndConversation, oSetWebcam, Nothing
    oCallReceived.Init oConversation, Nothing, oBascule, oConference
    oSetWebcam.Init Nothing, oEndConversation, oUnsetWebcam, Nothing
    oUnsetWebcam.Init Nothing, oEndConversation, oSetWebcam, Nothing
    oBascule.Init Nothing, oBascule, oBascule, Nothing
    oCompose.Init Nothing, oEndConversation, oCompose, oCompose
    oConference.Init Nothing, oEndConversation, Nothing, Nothing
    oReadMessages.Init oConversation, Nothing, oReadMessages, Nothing
    oEndConversation.Init oCompose, Nothing, oSearchName, oSearchName

    Set oCurrent = oInit

End Sub

Sub captureEvent(sEvent As String)

    Set oCurrent = oCurrent.executeEvent(sEvent)
    
End Sub


