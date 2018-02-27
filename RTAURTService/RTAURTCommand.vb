Imports System
Imports RTAURTService
Imports UrtTlbLib
Imports RTAInterfaces
Imports URT

Public Interface RTAURTCommand
    Sub Execute(ByVal vbfb As URTVBFunctionBlock, ByRef params As RTAURTCommandParameters)
End Interface

Public Interface IRTAURTCommandParameters
    Function GetCommand() As RTAURTCommand
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

Public Class CommandFactory

    Private Shared _RTAURTCommandConnect As RTAURTCommandConnect
    Private Shared _RTAURTCommandEnableHistory As RTAURTCommandEnableHistory
    Private Shared _RTAURTCommandExecute As RTAURTCommandExecute
    Private Shared _RTAURTCommandSetValues As RTAURTCommandSetValues
    Private Shared _RTAURTCommandMessage As RTAURTCommandMessage
    Private Shared _RTAURTCommandClearLogs As RTAURTCommandClearLogs
    Private Shared _RTAURTCommandSequence As RTAURTCommandSequence

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

    Public Shared Function MakeExecuteCommand() As RTAURTCommand
        Return MakeCommand(_RTAURTCommandExecute)
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
End Class

Public Class RTAURTCommandConnect
    Implements RTAURTCommand

    Public Sub Execute(vbfb As URTVBFunctionBlock, ByRef params As RTAURTCommandParameters) Implements RTAURTCommand.Execute

        Dim init As Boolean = True

        Dim connectParams As RTAURTCommandConnectParameters
        If Not params Is Nothing Then
            connectParams = TryCast(params, RTAURTCommandConnectParameters)
            If Not connectParams Is Nothing Then
                init = connectParams.Init
            Else
                RaiseMessage.Raise(vbfb, "Unexpected Type for Command Connect parameter")
            End If
        End If
        vbfb.Connect(init)
    End Sub
End Class

Public Class RTAURTCommandEnableHistory
    Implements RTAURTCommand

    Public Sub Execute(vbfb As URTVBFunctionBlock, ByRef params As RTAURTCommandParameters) Implements RTAURTCommand.Execute

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

    End Sub
End Class

Public Class RTAURTCommandExecute
    Implements RTAURTCommand

    Public Sub Execute(vbfb As URTVBFunctionBlock, ByRef params As RTAURTCommandParameters) Implements RTAURTCommand.Execute

        Dim count As Integer = 1

        Dim executeParams As RTAURTCommandExecuteParameters
        If Not params Is Nothing Then
            executeParams = TryCast(params, RTAURTCommandExecuteParameters)
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
    End Sub
End Class

Public Class RTAURTCommandErrorMessages
    Public Const SentValuesErrorCantFindElement As String = "Set Values Error: Cannot Find Element"
    Public Const SequenceErrorNoCommands As String = "Sequence Error: Sequence contains no commands"
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

    Public Sub Execute(vbfb As URTVBFunctionBlock, ByRef params As RTAURTCommandParameters) Implements RTAURTCommand.Execute

        Dim setValueParams As RTAURTCommandSetValuesParameters
        Dim idata As IUrtData

        If Not params Is Nothing Then
            setValueParams = TryCast(params, RTAURTCommandSetValuesParameters)
            If Not setValueParams Is Nothing Then
                For Each svItem In setValueParams.NameValueList
                    idata = vbfb.GetElement(svItem.Name)
                    If idata Is Nothing Then
                        RaiseMessage.Raise(vbfb, ErrorMessages.CantFindElement)
                        Continue For
                    End If
                    If svItem.IsScalar Then
                        Try
                            idata.PutVariantValue(svItem.Value, "")
                        Catch ex As Exception
                            RaiseMessage.Raise(vbfb, ErrorMessages.ValueFailsToConvert)
                        End Try
                    Else
                        Dim iRtaData As IRTAUrtData
                        iRtaData = TryCast(vbfb.GetElement(svItem.Name), IRTAUrtData)
                        If Not iRtaData Is Nothing Then
                            If svItem.Index < 0 Or svItem.Index > idata.Size - 1 Then
                                RaiseMessage.Raise(vbfb, ErrorMessages.IndexOutOfRange)
                            Else
                                Try
                                    iRtaData.Item(svItem.Index) = svItem.Value
                                Catch ex As Exception
                                    RaiseMessage.Raise(vbfb, ErrorMessages.ValueFailsToConvert)
                                End Try
                            End If
                        Else
                            RaiseMessage.Raise(vbfb, "SetValues Command: Error get IRTAURTData interface")
                        End If
                    End If

                Next

            Else

            End If
        End If
    End Sub
End Class

Public Class RTAURTCommandMessage
    Implements RTAURTCommand

    Public Sub Execute(vbfb As URTVBFunctionBlock, ByRef params As RTAURTCommandParameters) Implements RTAURTCommand.Execute

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

    End Sub
End Class

Public Class RTAURTCommandClearLogs
    Implements RTAURTCommand

    Public Enum Logs
        MessageLog = 0
        HistoryLog = 1
        MessageAndHistory = 2
    End Enum

    Public Sub Execute(vbfb As URTVBFunctionBlock, ByRef params As RTAURTCommandParameters) Implements RTAURTCommand.Execute

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

    End Sub
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

Public Class RTAURTCommandExecuteParameters
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
        Return CommandFactory.MakeExecuteCommand
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

