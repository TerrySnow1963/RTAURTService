﻿Imports RTAURTService
Imports UrtTlbLib
Imports RTAInterfaces

Public Class URTVBQALSimpleScript
    Inherits RTAURTService.URTVBFunctionBlock

    Private Shared _theScript As URT.VBScript = New URT.VBScript

    Public Sub New()
        MyBase.New(_theScript, Nothing)
    End Sub

    Public Sub New(ByVal messageLogger As IRTAUrtMessageLog)
        MyBase.New(_theScript, messageLogger)
    End Sub

    Protected Overrides Sub InternalConnect(ByVal WithInit As Boolean)
        _theScript.ConnectDataItems(WithInit)
    End Sub

    Protected Overrides Sub InternalExecute(ByVal iCause As Integer, ByVal pITMScheduler As IUrtTreeMember)
        _theScript.Execute(iCause, pITMScheduler)
    End Sub

End Class