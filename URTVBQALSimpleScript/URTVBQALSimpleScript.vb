Imports RTAURTService
Imports UrtTlbLib
Imports RTAInterfaces
Imports URT

Public Class ScriptProxyFactory
    Public Sub New()

    End Sub
    Public Function MakeProxy(ByVal messageLogger As IRTAUrtMessageLog, Optional historyLogger As IRTAUrtHistoryLog = Nothing) As URTVBFunctionBlock
        Return New URTVBQALSimpleScript(messageLogger, historyLogger)
    End Function
End Class

Public Class URTVBQALSimpleScript
    Inherits RTAURTService.URTVBFunctionBlock

    Private Shared _theScript As URT.VBScript = New URT.VBScript

    Public Sub New()
        MyBase.New(_theScript, Nothing, Nothing)
    End Sub

    Public Sub New(ByVal messageLogger As IRTAUrtMessageLog, Optional historyLogger As IRTAUrtHistoryLog = Nothing)
        MyBase.New(_theScript, messageLogger, historyLogger)
    End Sub

    Protected Overrides Sub InternalConnect(ByVal WithInit As Boolean)
        _theScript.ConnectDataItems(WithInit)
    End Sub

    Protected Overrides Sub InternalExecute(ByVal iCause As Integer, ByVal pITMScheduler As IUrtTreeMember)
        _theScript.Execute(iCause, pITMScheduler)
    End Sub

End Class