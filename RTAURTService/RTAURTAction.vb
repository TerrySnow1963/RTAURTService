Imports System
Imports RTAURTService
Imports UrtTlbLib
Imports RTAInterfaces
Imports URT

Public Interface RTAURTAction
    Sub Execute(ByVal vbfb As URTVBFunctionBlock, ByRef params As RTAURTActionParameters)
End Interface

Public Interface IRTAURTActionParameters
    Function GetCommand() As RTAURTAction
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

    Private Shared _RTAURTActionConnect As RTAURTActionConnect
    Private Shared _RTAURTActionEnableHistory As RTAURTActionEnableHistory
    Private Shared _RTAURTActionExecute As RTAURTActionExecute
    Private Shared _RTAURTActionSetValues As RTAURTActionSetValues
    Private Shared _RTAURTActionMessage As RTAURTActionMessage
    Private Shared _RTAURTActionClearLogs As RTAURTActionClearLogs
    Private Shared _RTAURTActionSequence As RTAURTActionSequence

    Private Shared Function MakeCommand(Of T As {New})(ByVal var As T) As RTAURTAction
        If var Is Nothing Then
            var = New T
        End If
        Return var
    End Function

    Public Shared Function MakeConnectCommand() As RTAURTAction
        Return MakeCommand(_RTAURTActionConnect)
    End Function

    Public Shared Function MakeEnableHistoryCommand() As RTAURTAction
        Return MakeCommand(_RTAURTActionEnableHistory)
    End Function

    Public Shared Function MakeExecuteCommand() As RTAURTAction
        Return MakeCommand(_RTAURTActionExecute)
    End Function

    Public Shared Function MakeSetValuesCommand() As RTAURTAction
        Return MakeCommand(_RTAURTActionSetValues)
    End Function

    Public Shared Function MakeMessageCommand() As RTAURTAction
        Return MakeCommand(_RTAURTActionMessage)
    End Function

    Public Shared Function MakeClearLogsCommand() As RTAURTAction
        Return MakeCommand(_RTAURTActionClearLogs)
    End Function

    Public Shared Function MakeSequenceCommand() As RTAURTAction
        Return MakeCommand(_RTAURTActionSequence)
    End Function
End Class

Public Class RTAURTActionConnect
    Implements RTAURTAction

    Public Sub Execute(vbfb As URTVBFunctionBlock, ByRef params As RTAURTActionParameters) Implements RTAURTAction.Execute

        Dim init As Boolean = True

        Dim connectParams As RTAURTActionConnectParameters
        If Not params Is Nothing Then
            connectParams = TryCast(params, RTAURTActionConnectParameters)
            If Not connectParams Is Nothing Then
                init = connectParams.Init
            Else
                RaiseMessage.Raise(vbfb, "Unexpected Type for action Connect parameter")
            End If
        End If
        vbfb.Connect(init)
    End Sub
End Class

Public Class RTAURTActionEnableHistory
    Implements RTAURTAction

    Public Sub Execute(vbfb As URTVBFunctionBlock, ByRef params As RTAURTActionParameters) Implements RTAURTAction.Execute

        Dim init As Boolean = True

        Dim enableHistParams As RTAURTActionEnableHistoryParameters
        If Not params Is Nothing Then
            enableHistParams = TryCast(params, RTAURTActionEnableHistoryParameters)
            If Not enableHistParams Is Nothing Then
                vbfb.HistoryLog = enableHistParams.HistoryLogger

                For Each name In enableHistParams.NamesToHistorize
                    If Not vbfb.HistoryLog.RegisterItem(CType(vbfb, IUrtTreeMember), name) Then
                        RaiseMessage.Raise(vbfb, "History - could not find name")
                    End If
                Next
            Else
                    RaiseMessage.Raise(vbfb, "Unexpected Type for action Connect parameter")
            End If
        End If

    End Sub
End Class

Public Class RTAURTActionExecute
    Implements RTAURTAction

    Public Sub Execute(vbfb As URTVBFunctionBlock, ByRef params As RTAURTActionParameters) Implements RTAURTAction.Execute

        Dim count As Integer = 1

        Dim executeParams As RTAURTActionExecuteParameters
        If Not params Is Nothing Then
            executeParams = TryCast(params, RTAURTActionExecuteParameters)
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

Public Class RTAURTActionErrorMessages
    Public Const SentValuesErrorCantFindElement As String = "Set Values Error: Cannot Find Element"
    Public Const SequenceErrorNoCommands As String = "Sequence Error: Sequence contains no commands"
End Class

Public Class RTAURTActionSetValues
    Implements RTAURTAction

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

    Public Sub Execute(vbfb As URTVBFunctionBlock, ByRef params As RTAURTActionParameters) Implements RTAURTAction.Execute

        Dim setValueParams As RTAURTActionSetValuesParameters
        Dim idata As IUrtData

        If Not params Is Nothing Then
            setValueParams = TryCast(params, RTAURTActionSetValuesParameters)
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

Public Class RTAURTActionMessage
    Implements RTAURTAction

    Public Sub Execute(vbfb As URTVBFunctionBlock, ByRef params As RTAURTActionParameters) Implements RTAURTAction.Execute

        Dim msgString As String

        Dim messageParams As RTAURTActionMessageParameters
        If Not params Is Nothing Then
            messageParams = TryCast(params, RTAURTActionMessageParameters)
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

Public Class RTAURTActionClearLogs
    Implements RTAURTAction

    Public Enum Logs
        MessageLog = 0
        HistoryLog = 1
        MessageAndHistory = 2
    End Enum

    Public Sub Execute(vbfb As URTVBFunctionBlock, ByRef params As RTAURTActionParameters) Implements RTAURTAction.Execute

        Dim clearLogsParams As RTAURTActionClearLogsParameters
        If Not params Is Nothing Then
            clearLogsParams = TryCast(params, RTAURTActionClearLogsParameters)
            If Not clearLogsParams Is Nothing Then
                Select Case clearLogsParams.LogsToClear
                    Case Logs.MessageLog
                        vbfb.ClearMessageLog
                    Case Logs.HistoryLog
                        vbfb.ClearHistoryLog
                    Case Logs.MessageAndHistory
                        vbfb.ClearMessageLog
                        vbfb.ClearHistoryLog
                End Select
            Else

            End If
        End If

    End Sub
End Class

Public MustInherit Class RTAURTActionParameters
    Implements IRTAURTActionParameters

    Public MustOverride Function GetCommand() As RTAURTAction Implements IRTAURTActionParameters.GetCommand

End Class

Public Class RTAURTActionConnectParameters
    Inherits RTAURTActionParameters
    Public Property Init As Boolean
    Private Sub New()

    End Sub
    Public Sub New(bInit As Boolean)
        Init = bInit
    End Sub

    Public Overrides Function GetCommand() As RTAURTAction
        Return CommandFactory.MakeConnectCommand
    End Function
End Class

Public Class RTAURTActionExecuteParameters
    Inherits RTAURTActionParameters
    Public Property Count As Integer
    Private Sub New()
    End Sub
    Public Sub New(Optional ByVal cnt As Integer = 1)
        If cnt < 1 Then
            cnt = 1
        End If
        Count = cnt
    End Sub

    Public Overrides Function GetCommand() As RTAURTAction
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

Public Class RTAURTActionSetValuesParameters
    Inherits RTAURTActionParameters

    Public _list As List(Of SetValueData)
    Public Sub New()
        _list = New List(Of SetValueData)
    End Sub

    Public Sub AddValue(ByVal Name As String, ByVal val As Object, Optional ByVal idx As Integer = -1)
        _list.Add(New SetValueData(Name, val, idx))
    End Sub

    Public Overrides Function GetCommand() As RTAURTAction
        Return CommandFactory.MakeSetValuesCommand
    End Function

    Public ReadOnly Property NameValueList As List(Of SetValueData)
        Get
            Return _list
        End Get
    End Property
End Class

Public Class RTAURTActionEnableHistoryParameters
    Inherits RTAURTActionParameters

    Public ReadOnly Property HistoryLogger As IRTAUrtHistoryLog

    Public NamesToHistorize As List(Of String)
    Public Sub New(ByVal hist As RTAUrtTraceHistory)
        _HistoryLogger = hist
        NamesToHistorize = New List(Of String)
    End Sub

    Public Sub EnableHistoryFor(ByVal Name As String)
        NamesToHistorize.Add(Name)
    End Sub

    Public Overrides Function GetCommand() As RTAURTAction
        Return CommandFactory.MakeEnableHistoryCommand
    End Function
End Class

Public Class RTAURTActionMessageParameters
    Inherits RTAURTActionParameters

    Public Sub New(ByVal msg As String)
        MessageText = msg

    End Sub

    Public ReadOnly Property MessageText As String

    Public Overrides Function GetCommand() As RTAURTAction
        Return CommandFactory.MakeMessageCommand()
    End Function
End Class

Public Class RTAURTActionClearLogsParameters
    Inherits RTAURTActionParameters

    Public Sub New(ByVal lgsToClear As RTAURTActionClearLogs.Logs)
        LogsToClear = lgsToClear

    End Sub

    Public ReadOnly Property LogsToClear As RTAURTActionClearLogs.Logs

    Public Overrides Function GetCommand() As RTAURTAction
        Return CommandFactory.MakeClearLogsCommand
    End Function
End Class

