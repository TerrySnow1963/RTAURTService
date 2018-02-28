Imports RTAURTService

Public Class CommandExecutive
    Implements ICommandCallback

    Private _vbfb As URTVBFunctionBlock
    Private _callback As ICommandCallback

    ReadOnly Property TotalCommandsExecuted As Integer
    Private Sub New()

    End Sub
    Public Sub New(vb As URTVBFunctionBlock)
        _vbfb = vb
        _callback = CommandCallbacks.LimitCommandCallsTo(100)
        _TotalCommandsExecuted = 0
    End Sub

    Private Sub PreMessageProcess()

    End Sub
    Private Sub PostMessageProcess()
        _TotalCommandsExecuted += 1
    End Sub
    Public Function Invoke(Of T As IRTAURTCommandParameters)(param As T) As ICommandResult
        Dim cmdResult As ICommandResult
        PreMessageProcess()
        cmdResult = param.GetCommand.Execute(_vbfb, param, Me)
        PostMessageProcess()
        Return cmdResult
    End Function

    Public Function CanContinue() As Boolean Implements ICommandCallback.CanContinue
        Return _callback.CanContinue
    End Function
End Class
