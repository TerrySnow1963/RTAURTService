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
    Public Sub New(vb As URTVBFunctionBlock, Optional callBack As ICommandCallback = Nothing)
        _vbfb = vb
        If callBack Is Nothing Then callBack = CommandCallbacks.Continue
        _callback = callBack
        _TotalCommandsExecuted = 0
        _currentParams = RTAURTCommandNullParameters.Instance
    End Sub

    Private Sub PreProcessCommand(Of T As IRTAURTCommandParameters)(param As T)
        _currentParams = param
    End Sub
    Private Sub PostProcessCommand(Of T As IRTAURTCommandParameters)(param As T)
        _TotalCommandsExecuted += 1
        _currentParams = RTAURTCommandNullParameters.Instance
    End Sub
    Public Function Invoke(Of T As IRTAURTCommandParameters)(param As T) As ICommandResult
        Dim cmdResult As ICommandResult

        PreProcessCommand(param)
        If CanContinue() Then
            cmdResult = _currentParams.GetCommand.Execute(Me)
        Else
            cmdResult = CommandResults.Stopped
        End If
        PostProcessCommand(param)

        Return cmdResult
    End Function

    Public Function CanContinue() As Boolean Implements ICommandCallback.CanContinue
        Return _callback.CanContinue
    End Function
End Class
