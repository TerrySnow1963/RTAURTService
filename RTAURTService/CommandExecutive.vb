Option Explicit On
Option Strict On
Imports RTAURTService

Public Class CommandExecutive
    Implements ICommandCallback
    Implements ICommandContext

    Private _vbfb As URTVBFunctionBlock
    Private _callback As ICommandCallback
    Private _currentParams As IRTAURTCommandParameters

    Private _commandStack As Stack(Of String)

    Private Const _DEFAULT_RECURSION_DEPTH As Integer = 5
    Private _recursionDepth As Integer
    Private p As ICommandCallback
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

    Public Property RecursionDepth As Integer
        Get
            Return _recursionDepth
        End Get
        Set(value As Integer)
            If value < 1 Then
                Throw New ArgumentOutOfRangeException("value")
            End If
            _recursionDepth = value
        End Set
    End Property

    Private Sub New()

    End Sub
    Public Sub New(vb As URTVBFunctionBlock, Optional callBack As ICommandCallback = Nothing)
        _vbfb = vb
        If callBack Is Nothing Then callBack = CommandCallbacks.Continue
        _callback = callBack
        _TotalCommandsExecuted = 0
        _currentParams = RTAURTCommandNullParameters.Instance
        _commandStack = New Stack(Of String)
        _recursionDepth = _DEFAULT_RECURSION_DEPTH
    End Sub


    Private Sub PreProcessCommand(Of T As IRTAURTCommandParameters)(param As T)
        _currentParams = param
        Dim name As String = param.TryGetRecursiveName
        If name IsNot Nothing Then _commandStack.Push(name)

        If _commandStack.Where(Function(x) x = name).Count > _recursionDepth Then
            _callback = CommandCallbacks.Stop
            RaiseMessage.Raise(_vbfb, "Recursive Limit Reached")
        End If
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
