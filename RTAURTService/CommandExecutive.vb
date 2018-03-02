Option Explicit On
Option Strict On
Imports RTAURTService

Public Class CommandExecutive
    Implements ICommandCallback
    Implements ICommandContext

    Private _vbfb As URTVBFunctionBlock
    Private _callback As ICommandCallback
    Private _currentParams As IRTAURTCommandParameters

    ReadOnly Property TotalCommandsExecuted As Integer

    Public ReadOnly Property cmdExec As CommandExecutive Implements ICommandContext.cmdExec
        Get
            Return Me
        End Get
    End Property

    Public ReadOnly Property params As IRTAURTCommandParameters Implements ICommandContext.params
        Get
            Return _currentParams
        End Get
    End Property

    Public ReadOnly Property vbfb As URTVBFunctionBlock Implements ICommandContext.vbfb
        Get
            Return _vbfb
        End Get
    End Property

    Private Sub New()

    End Sub
    Public Sub New(vb As URTVBFunctionBlock)
        _vbfb = vb
        'Todo add callback as argument to Sub New
        _callback = CommandCallbacks.LimitCommandCallsTo(100)
        _TotalCommandsExecuted = 0
        _currentParams = RTAURTCommandNullParameters.Instance
    End Sub

    Private Sub PreProcessCommand()

    End Sub
    Private Sub PostProcessCommand()
        _TotalCommandsExecuted += 1
    End Sub
    Public Function Invoke(Of T As IRTAURTCommandParameters)(param As T) As ICommandResult
        Dim cmdResult As ICommandResult

        _currentParams = param
        PreProcessCommand()

        cmdResult = _currentParams.GetCommand.Execute(Me)
        PostProcessCommand()
        _currentParams = RTAURTCommandNullParameters.Instance
        Return cmdResult
    End Function

    Public Function CanContinue() As Boolean Implements ICommandCallback.CanContinue
        Return _callback.CanContinue
    End Function
End Class
