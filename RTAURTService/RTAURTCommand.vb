Imports System
Imports RTAURTService
Imports UrtTlbLib
Imports RTAInterfaces
Imports URT

Public Interface ICommandContext
    ReadOnly Property cmdExec As CommandExecutive
    ReadOnly Property params As IRTAURTCommandParameters
End Interface

Public Interface RTAURTCommand
    Function Execute(ByVal vbfb As URTVBFunctionBlock,
                     ByVal params As IRTAURTCommandParameters,
                     Optional ByVal callback As ICommandCallback = Nothing) As ICommandResult
    Function Execute(ByVal context As ICommandContext) As ICommandResult
End Interface

Public Interface IRTAURTCommandParameters
    Function GetCommand() As RTAURTCommand
End Interface

Public Enum RTAURTCommandResultCode
    CMD_DONE = 0
    CMD_ERROR = 1
    CMD_STOPPED = 2
End Enum

Public Interface ICommandResult
    Function GetResultCode() As RTAURTCommandResultCode
End Interface

Public Interface ICommandCallback
    Function CanContinue() As Boolean
End Interface

Public Class RaiseMessage
    Public Shared Sub Raise(ByVal vbfb As URTVBFunctionBlock, ByVal msgString As String)
        Dim msg As New ConMessageClass
        msg.AckRequired = False
        msg.Priority = urtMSGPRIORITY.msgHI
        msg.Group = urtMSGGROUP.msgERROR
        msg.text = msgString
        vbfb.Raise(msg, 0)
    End Sub
End Class

Public Class ICommandResults
    Private Shared _Done
    Public Shared ReadOnly Property Done As ICommandResult
        Get
            If _Done Is Nothing Then
                _Done = New DoneResult
            End If
            Return _Done
        End Get
    End Property

    Private Class DoneResult
        Implements ICommandResult

        Public Function GetCommandResult() As RTAURTCommandResultCode Implements ICommandResult.GetResultCode
            Return RTAURTCommandResultCode.CMD_DONE
        End Function
    End Class

    Private Shared _Error
    Public Shared ReadOnly Property [Error] As ICommandResult
        Get
            If _Error Is Nothing Then
                _Error = New ErrorResult
            End If
            Return _Error
        End Get
    End Property

    Private Class ErrorResult
        Implements ICommandResult

        Public Function GetCommandResult() As RTAURTCommandResultCode Implements ICommandResult.GetResultCode
            Return RTAURTCommandResultCode.CMD_ERROR
        End Function
    End Class

    Private Shared _Stopped
    Public Shared ReadOnly Property Stopped As ICommandResult
        Get
            If _Stopped Is Nothing Then
                _Stopped = New StoppedResult
            End If
            Return _Stopped
        End Get
    End Property

    Private Class StoppedResult
        Implements ICommandResult

        Public Function GetCommandResult() As RTAURTCommandResultCode Implements ICommandResult.GetResultCode
            Return RTAURTCommandResultCode.CMD_STOPPED
        End Function
    End Class
End Class


Public Class CommandCallbacks
    Private Shared _Stop As ICommandCallback
    Public Shared ReadOnly Property [Stop] As ICommandCallback
        Get
            If _Stop Is Nothing Then
                _Stop = New StopCallback
            End If
            Return _Stop
        End Get
    End Property

    Private Shared _Continue As ICommandCallback
    Public Shared Property [Continue] As ICommandCallback
        Get
            If _Continue Is Nothing Then
                _Continue = New ContinueCallback
            End If
            Return _Continue
        End Get
        Set(value As ICommandCallback)

        End Set
    End Property

    Public Shared ReadOnly Property LimitCommandCallsTo(commandCallLimit As Integer) As ICommandCallback
        Get
            Return New LimitCallbackCount(commandCallLimit)
        End Get
    End Property

    Private Class StopCallback
        Implements ICommandCallback

        Public Function CanContinue() As Boolean Implements ICommandCallback.CanContinue
            Return False
        End Function

    End Class

    Private Class ContinueCallback
        Implements ICommandCallback

        Public Function CanContinue() As Boolean Implements ICommandCallback.CanContinue
            Return True
        End Function

    End Class

    Private Class LimitCallbackCount
        Implements ICommandCallback

        Private _limit As Integer
        Private _counter As Integer
        Public Sub New(limit As Integer)
            If limit < 0 Then limit = 0
            _limit = limit
            _counter = 0
        End Sub

        Public Function CanContinue() As Boolean Implements ICommandCallback.CanContinue
            _counter += 1
            Return _counter <= _limit
        End Function

    End Class
    Friend Shared Function GetSafeCallbackAndCheckforStop(callback As ICommandCallback) As Boolean
        If callback IsNot Nothing Then Return Not callback.CanContinue
        callback = CommandCallbacks.Continue
        Return False

    End Function
End Class

Public Class CommandFactory

    Private Shared _RTAURTCommandConnect As RTAURTCommandConnect
    Private Shared _RTAURTCommandEnableHistory As RTAURTCommandEnableHistory
    Private Shared _RTAURTCommandExecuteVB As RTAURTCommandExecuteVB
    Private Shared _RTAURTCommandSetValues As RTAURTCommandSetValues
    Private Shared _RTAURTCommandMessage As RTAURTCommandMessage
    Private Shared _RTAURTCommandClearLogs As RTAURTCommandClearLogs
    Private Shared _RTAURTCommandSequence As RTAURTCommandSequence
    Private Shared _RTAURTLCommandLinkedSequence As RTAURTCommandLinkedSequence

    Private Shared Function MakeCommand(Of T As {New})(ByVal var As T) As RTAURTCommand
        If var Is Nothing Then
            var = New T
        End If
        Return var
    End Function

    Public Shared Function MakeConnectCommand() As RTAURTCommand
        Return MakeCommand(_RTAURTCommandConnect)
    End Function

    Public Shared Function MakeEnableHistoryCommand() As RTAURTCommand
        Return MakeCommand(_RTAURTCommandEnableHistory)
    End Function

    Public Shared Function MakeExecuteVBCommand() As RTAURTCommand
        Return MakeCommand(_RTAURTCommandExecuteVB)
    End Function

    Public Shared Function MakeSetValuesCommand() As RTAURTCommand
        Return MakeCommand(_RTAURTCommandSetValues)
    End Function

    Public Shared Function MakeMessageCommand() As RTAURTCommand
        Return MakeCommand(_RTAURTCommandMessage)
    End Function

    Public Shared Function MakeClearLogsCommand() As RTAURTCommand
        Return MakeCommand(_RTAURTCommandClearLogs)
    End Function

    Public Shared Function MakeSequenceCommand() As RTAURTCommand
        Return MakeCommand(_RTAURTCommandSequence)
    End Function

    Public Shared Function MakeLinkedSequenceCommand() As RTAURTCommand
        Return MakeCommand(_RTAURTLCommandLinkedSequence)
    End Function
End Class

Public Class RTAURTCommandConnect
    Implements RTAURTCommand

    Public Function Execute(context As ICommandContext) As ICommandResult Implements RTAURTCommand.Execute
        Throw New NotImplementedException()
    End Function

    Public Function Execute(vbfb As URTVBFunctionBlock,
                            ByVal params As IRTAURTCommandParameters,
                            Optional ByVal callback As ICommandCallback = Nothing) As ICommandResult Implements RTAURTCommand.Execute
        If CommandCallbacks.GetSafeCallbackAndCheckforStop(callback) Then Return ICommandResults.Stopped

        Dim init As Boolean = True
        Dim result As ICommandResult

        Dim connectParams As RTAURTCommandConnectParameters
        If Not params Is Nothing Then
            connectParams = TryCast(params, RTAURTCommandConnectParameters)
            If Not connectParams Is Nothing Then
                vbfb.Connect(connectParams.Init)
                result = ICommandResults.Done
            Else
                RaiseMessage.Raise(vbfb, "Unexpected Type for Command Connect parameter")
                result = ICommandResults.Error
            End If
        Else
            RaiseMessage.Raise(vbfb, "No Parameters")
            result = ICommandResults.Error
        End If

        Return result

    End Function
End Class

Public Class RTAURTCommandEnableHistory
    Implements RTAURTCommand

    Public Function Execute(context As ICommandContext) As ICommandResult Implements RTAURTCommand.Execute
        Throw New NotImplementedException()
    End Function

    Public Function Execute(vbfb As URTVBFunctionBlock,
                            ByVal params As IRTAURTCommandParameters,
                            Optional ByVal callback As ICommandCallback = Nothing) As ICommandResult Implements RTAURTCommand.Execute
        If CommandCallbacks.GetSafeCallbackAndCheckforStop(callback) Then Return ICommandResults.Stopped

        Dim init As Boolean = True

        Dim enableHistParams As RTAURTCommandEnableHistoryParameters
        If Not params Is Nothing Then
            enableHistParams = TryCast(params, RTAURTCommandEnableHistoryParameters)
            If Not enableHistParams Is Nothing Then
                vbfb.HistoryLog = enableHistParams.HistoryLogger

                For Each name In enableHistParams.NamesToHistorize
                    If Not vbfb.HistoryLog.RegisterItem(CType(vbfb, IUrtTreeMember), name) Then
                        RaiseMessage.Raise(vbfb, "History - could not find name")
                    End If
                Next
            Else
                RaiseMessage.Raise(vbfb, "Unexpected Type for Command Connect parameter")
            End If
        End If

        Return ICommandResults.Done
    End Function
End Class

Public Class RTAURTCommandExecuteVB
    Implements RTAURTCommand

    Public Function Execute(context As ICommandContext) As ICommandResult Implements RTAURTCommand.Execute
        Throw New NotImplementedException()
    End Function

    Public Function Execute(vbfb As URTVBFunctionBlock,
                            ByVal params As IRTAURTCommandParameters,
                            Optional ByVal callback As ICommandCallback = Nothing) As ICommandResult Implements RTAURTCommand.Execute
        If CommandCallbacks.GetSafeCallbackAndCheckforStop(callback) Then Return ICommandResults.Stopped

        Dim count As Integer = 1

        Dim executeParams As RTAURTCommandExecuteVBParameters
        If Not params Is Nothing Then
            executeParams = TryCast(params, RTAURTCommandExecuteVBParameters)
            If Not executeParams Is Nothing Then
                count = executeParams.Count
                If count < 1 Then
                    RaiseMessage.Raise(vbfb, "Count of executions < 1, forced to 1")
                End If
            Else

            End If
        End If

        For ii = 0 To count - 1
            vbfb.Execute(0, Nothing)
        Next

        Return ICommandResults.Done
    End Function
End Class

Public Class RTAURTCommandErrorMessages
    Public Const SentValuesErrorCantFindElement As String = "Set Values Error: Cannot Find Element"
    Public Const SequenceErrorNoCommands As String = "Sequence Error: Sequence contains no commands"
    Public Const LinkSequenceErrorNoSuchName As String = "Linked Sequence Error: No Such Sequence"
End Class

Public Class RTAURTCommandSetValues
    Implements RTAURTCommand

    Public Structure ErrorMessagesStruct
        Private Shared _PREFIX As String = "Set Values Error: "
        Public ReadOnly Property CantFindElement As String
            Get
                Return _PREFIX & "Cannot Find Element"
            End Get
        End Property
        Public ReadOnly Property IndexOutOfRange As String
            Get
                Return _PREFIX & "index out of range for array"
            End Get
        End Property
        Public ReadOnly Property ValueFailsToConvert As String
            Get
                Return _PREFIX & "Unable to convert value to type"
            End Get
        End Property


    End Structure
    Public ErrorMessages As ErrorMessagesStruct

    Public Function Execute(vbfb As URTVBFunctionBlock,
                            ByVal params As IRTAURTCommandParameters,
                            Optional ByVal callback As ICommandCallback = Nothing) As ICommandResult Implements RTAURTCommand.Execute
        If CommandCallbacks.GetSafeCallbackAndCheckforStop(callback) Then Return ICommandResults.Stopped

        Dim result As ICommandResult
        Dim setValueParams As RTAURTCommandSetValuesParameters
        Dim idata As IUrtData

        If Not params Is Nothing Then
            setValueParams = TryCast(params, RTAURTCommandSetValuesParameters)
            If Not setValueParams Is Nothing Then
                For Each svItem In setValueParams.NameValueList
                    idata = vbfb.GetElement(svItem.Name)
                    If idata Is Nothing Then
                        RaiseMessage.Raise(vbfb, ErrorMessages.CantFindElement)
                        result = ICommandResults.Error
                        Exit For
                    End If
                    If svItem.IsScalar Then
                        Try
                            idata.PutVariantValue(svItem.Value, "")
                        Catch ex As Exception
                            RaiseMessage.Raise(vbfb, ErrorMessages.ValueFailsToConvert)
                            result = ICommandResults.Error
                        End Try
                    Else
                        Dim iRtaData As IRTAUrtData
                        iRtaData = TryCast(vbfb.GetElement(svItem.Name), IRTAUrtData)
                        If Not iRtaData Is Nothing Then
                            If svItem.Index < 0 Or svItem.Index > idata.Size - 1 Then
                                RaiseMessage.Raise(vbfb, ErrorMessages.IndexOutOfRange)
                                result = ICommandResults.Error
                            Else
                                Try
                                    iRtaData.Item(svItem.Index) = svItem.Value
                                Catch ex As Exception
                                    RaiseMessage.Raise(vbfb, ErrorMessages.ValueFailsToConvert)
                                    result = ICommandResults.Error
                                End Try
                            End If
                        Else
                            RaiseMessage.Raise(vbfb, "SetValues Command: Error get IRTAURTData interface")
                            result = ICommandResults.Error
                        End If
                    End If

                Next

            Else

            End If
        End If
        If result Is Nothing Then result = ICommandResults.Done

        Return result
    End Function

    Public Function Execute(context As ICommandContext) As ICommandResult Implements RTAURTCommand.Execute
        Throw New NotImplementedException()
    End Function
End Class

Public Class RTAURTCommandMessage
    Implements RTAURTCommand

    Public Function Execute(context As ICommandContext) As ICommandResult Implements RTAURTCommand.Execute
        Throw New NotImplementedException()
    End Function

    Public Function Execute(vbfb As URTVBFunctionBlock,
                            ByVal params As IRTAURTCommandParameters,
                            Optional ByVal callback As ICommandCallback = Nothing) As ICommandResult Implements RTAURTCommand.Execute
        If CommandCallbacks.GetSafeCallbackAndCheckforStop(callback) Then Return ICommandResults.Stopped

        Dim msgString As String

        Dim messageParams As RTAURTCommandMessageParameters
        If Not params Is Nothing Then
            messageParams = TryCast(params, RTAURTCommandMessageParameters)
            If Not messageParams Is Nothing Then
                msgString = messageParams.MessageText
                If String.IsNullOrEmpty(msgString) Then
                    RaiseMessage.Raise(vbfb, "MessageCommand writing empty message")
                Else
                    RaiseMessage.Raise(vbfb, msgString)
                End If
            Else

            End If
        End If

        Return ICommandResults.Done

    End Function
End Class

Public Class RTAURTCommandClearLogs
    Implements RTAURTCommand

    Public Enum Logs
        MessageLog = 0
        HistoryLog = 1
        MessageAndHistory = 2
    End Enum

    Public Function Execute(vbfb As URTVBFunctionBlock,
                            ByVal params As IRTAURTCommandParameters,
                            Optional ByVal callback As ICommandCallback = Nothing) As ICommandResult Implements RTAURTCommand.Execute
        If CommandCallbacks.GetSafeCallbackAndCheckforStop(callback) Then Return ICommandResults.Stopped

        Dim clearLogsParams As RTAURTCommandClearLogsParameters
        If Not params Is Nothing Then
            clearLogsParams = TryCast(params, RTAURTCommandClearLogsParameters)
            If Not clearLogsParams Is Nothing Then
                Select Case clearLogsParams.LogsToClear
                    Case Logs.MessageLog
                        vbfb.ClearMessageLog()
                    Case Logs.HistoryLog
                        vbfb.ClearHistoryLog()
                    Case Logs.MessageAndHistory
                        vbfb.ClearMessageLog()
                        vbfb.ClearHistoryLog()
                End Select
            Else

            End If
        End If
        Return ICommandResults.Done
    End Function

    Public Function Execute(context As ICommandContext) As ICommandResult Implements RTAURTCommand.Execute
        Throw New NotImplementedException()
    End Function
End Class

Public MustInherit Class RTAURTCommandParameters
    Implements IRTAURTCommandParameters

    Public MustOverride Function GetCommand() As RTAURTCommand Implements IRTAURTCommandParameters.GetCommand

End Class

Public Class RTAURTCommandConnectParameters
    Inherits RTAURTCommandParameters
    Public Property Init As Boolean
    Private Sub New()

    End Sub
    Public Sub New(bInit As Boolean)
        Init = bInit
    End Sub

    Public Overrides Function GetCommand() As RTAURTCommand
        Return CommandFactory.MakeConnectCommand
    End Function
End Class

Public Class RTAURTCommandExecuteVBParameters
    Inherits RTAURTCommandParameters
    Public Property Count As Integer
    Private Sub New()
    End Sub
    Public Sub New(Optional ByVal cnt As Integer = 1)
        If cnt < 1 Then
            cnt = 1
        End If
        Count = cnt
    End Sub

    Public Overrides Function GetCommand() As RTAURTCommand
        Return CommandFactory.MakeExecuteVBCommand
    End Function
End Class

Public Class SetValueData
    Private Const _SCALAR_ITEM_INDEX As Integer = -1
    Public ReadOnly Property Name As String
    Public ReadOnly Property Value As Object
    Public ReadOnly Property Index As Integer
    Public ReadOnly Property IsScalar As Boolean
        Get
            Return Index = -1
        End Get
    End Property

    Public Sub New(ByVal nam As String, ByVal val As Object, Optional ByVal idx As Integer = _SCALAR_ITEM_INDEX)
        Name = nam
        Value = val
        Index = idx
    End Sub
End Class

Public Class RTAURTCommandSetValuesParameters
    Inherits RTAURTCommandParameters

    Public _list As List(Of SetValueData)
    Public Sub New()
        _list = New List(Of SetValueData)
    End Sub

    Public Sub AddValue(ByVal Name As String, ByVal val As Object, Optional ByVal idx As Integer = -1)
        _list.Add(New SetValueData(Name, val, idx))
    End Sub

    Public Overrides Function GetCommand() As RTAURTCommand
        Return CommandFactory.MakeSetValuesCommand
    End Function

    Public ReadOnly Property NameValueList As List(Of SetValueData)
        Get
            Return _list
        End Get
    End Property
End Class

Public Class RTAURTCommandEnableHistoryParameters
    Inherits RTAURTCommandParameters

    Public ReadOnly Property HistoryLogger As IRTAUrtHistoryLog

    Public NamesToHistorize As List(Of String)
    Public Sub New(ByVal hist As RTAUrtTraceHistory)
        _HistoryLogger = hist
        NamesToHistorize = New List(Of String)
    End Sub

    Public Sub EnableHistoryFor(ByVal Name As String)
        NamesToHistorize.Add(Name)
    End Sub

    Public Overrides Function GetCommand() As RTAURTCommand
        Return CommandFactory.MakeEnableHistoryCommand
    End Function
End Class

Public Class RTAURTCommandMessageParameters
    Inherits RTAURTCommandParameters

    Public Sub New(ByVal msg As String)
        MessageText = msg

    End Sub

    Public ReadOnly Property MessageText As String

    Public Overrides Function GetCommand() As RTAURTCommand
        Return CommandFactory.MakeMessageCommand()
    End Function
End Class

Public Class RTAURTCommandClearLogsParameters
    Inherits RTAURTCommandParameters

    Public Sub New(ByVal lgsToClear As RTAURTCommandClearLogs.Logs)
        LogsToClear = lgsToClear

    End Sub

    Public ReadOnly Property LogsToClear As RTAURTCommandClearLogs.Logs

    Public Overrides Function GetCommand() As RTAURTCommand
        Return CommandFactory.MakeClearLogsCommand
    End Function
End Class

