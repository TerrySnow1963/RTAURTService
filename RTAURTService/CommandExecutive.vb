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
        'Todo add callback as argument to Sub New
        _callback = CommandCallbacks.LimitCommandCallsTo(100)
        _TotalCommandsExecuted = 0
    End Sub

    Private Sub PreProcessCommand()

    End Sub
    Private Sub PostProcessCommand()
        _TotalCommandsExecuted += 1
    End Sub
    Public Function Invoke(Of T As IRTAURTCommandParameters)(param As T) As ICommandResult
        Dim cmdResult As ICommandResult
        PreProcessCommand()
        cmdResult = param.GetCommand.Execute(_vbfb, param, Me)
        PostProcessCommand()
        Return cmdResult
    End Function

    Public Function CanContinue() As Boolean Implements ICommandCallback.CanContinue
        Return _callback.CanContinue
    End Function
End Class
