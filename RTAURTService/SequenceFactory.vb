Imports RTAURTService
Imports RTAInterfaces
Imports URT


Public Class SequenceFactory

    Private Shared _list As List(Of RTAURTCommandSequenceParameters) = New List(Of RTAURTCommandSequenceParameters)

    Public Shared Function MakeSequence() As RTAURTCommandSequenceParameters

        Dim uniqueName As String = GetNextName()
        Dim params = New RTAURTCommandSequenceParameters(uniqueName)
        _list.Add(params)
        Return params

    End Function

    Public Shared _NameCounter As Integer = 0
    Public Const _BASE_NAME = "Seq"

    Private Shared Function GetNextName() As String
        Dim newName As String = _BASE_NAME & _NameCounter.ToString

        _NameCounter += 1
        Return newName
    End Function

    Public Shared Function GetSequenceByName(linkedSequenceName As String) As RTAURTCommandSequenceParameters
        Dim params As RTAURTCommandSequenceParameters = Nothing
        params = _list.Where(Function(x) x.Name = linkedSequenceName).FirstOrDefault
        Return params
    End Function
End Class


Public Class RTAURTCommandSequence
    Implements RTAURTCommand

    Public Function Execute(context As ICommandContext) As ICommandResult Implements RTAURTCommand.Execute
        Throw New NotImplementedException()
    End Function

    Public Function Execute(vbfb As URTVBFunctionBlock,
                            ByVal params As IRTAURTCommandParameters,
                            Optional ByVal callback As ICommandCallback = Nothing) As ICommandResult Implements RTAURTCommand.Execute
        If CommandCallbacks.GetSafeCallbackAndCheckforStop(callback) Then Return CommandResults.Stopped

        Dim seqParams As RTAURTCommandSequenceParameters
        Dim resultCmd As ICommandResult

        If Not params Is Nothing Then
            seqParams = TryCast(params, RTAURTCommandSequenceParameters)
            If Not seqParams Is Nothing Then
                'ToDo need to check for recursion
                For Each cmdParams In seqParams.GetCommandList
                    resultCmd = cmdParams.GetCommand.Execute(vbfb, cmdParams, callback)
                    If resultCmd.GetResultCode <> RTAURTCommandResultCode.CMD_DONE Then
                        Return resultCmd
                    End If
                Next
                If seqParams.GetCommandList.Count = 0 Then
                    RaiseMessage.Raise(vbfb, RTAURTCommandErrorMessages.SequenceErrorNoCommands)
                    Return CommandResults.Error
                End If
            Else
                RaiseMessage.Raise(vbfb, "Sequence Command Error: passed wrong parameter type")
                Return CommandResults.Error
            End If
        End If

        Return CommandResults.Done
    End Function
End Class

Public Class RTAURTCommandSequenceParameters
    Inherits RTAURTCommandParameters

    Private _name
    Public ReadOnly Property Name As String
        Get
            Return _name
        End Get
    End Property

    Private _list As List(Of IRTAURTCommandParameters)
    Public Sub New(ByVal nm As String)
        _list = New List(Of IRTAURTCommandParameters)
        _name = nm
    End Sub

    Public Overrides Function GetCommand() As RTAURTCommand
        Return CommandFactory.MakeSequenceCommand
    End Function

    Public Sub AddCommandParameters(params As IRTAURTCommandParameters)
        _list.Add(params)
    End Sub

    Public Function GetCommandList() As List(Of IRTAURTCommandParameters)
        Return _list
    End Function
End Class

Public Class RTAURTCommandLinkedSequence
    Implements RTAURTCommand

    Public Function Execute(context As ICommandContext) As ICommandResult Implements RTAURTCommand.Execute
        Dim seqParams As RTAURTCommandLinkedSequenceParameters
        Dim resultCmd As ICommandResult


        seqParams = CType(context.params, RTAURTCommandLinkedSequenceParameters)

        'ToDo need to check for recursion
        Dim linkedSeqParams As RTAURTCommandSequenceParameters = SequenceFactory.GetSequenceByName(seqParams.LinkedSequenceName)
        If linkedSeqParams Is Nothing Then
            RaiseMessage.Raise(context.vbfb, RTAURTCommandErrorMessages.LinkSequenceErrorNoSuchName)
            Return CommandResults.Error
        Else
            'resultCmd = linkedSeqParams.GetCommand.Execute(context.vbfb, linkedSeqParams, callback)
            resultCmd = context.cmdExec.Invoke(linkedSeqParams)
            Return resultCmd
        End If

        Return CommandResults.Done
    End Function

    Public Function Execute(vbfb As URTVBFunctionBlock,
                            ByVal params As IRTAURTCommandParameters,
                            Optional ByVal callback As ICommandCallback = Nothing) As ICommandResult Implements RTAURTCommand.Execute
        If CommandCallbacks.GetSafeCallbackAndCheckforStop(callback) Then Return CommandResults.Stopped

        Dim seqParams As RTAURTCommandLinkedSequenceParameters
        Dim resultCmd As ICommandResult

        If Not params Is Nothing Then
            seqParams = TryCast(params, RTAURTCommandLinkedSequenceParameters)
            If Not seqParams Is Nothing Then
                'ToDo need to check for recursion
                Dim linkedSeqParams As RTAURTCommandSequenceParameters = SequenceFactory.GetSequenceByName(seqParams.LinkedSequenceName)
                If linkedSeqParams Is Nothing Then
                    RaiseMessage.Raise(vbfb, RTAURTCommandErrorMessages.LinkSequenceErrorNoSuchName)
                    Return CommandResults.Error
                Else
                    resultCmd = linkedSeqParams.GetCommand.Execute(vbfb, linkedSeqParams, callback)
                    Return resultCmd
                End If
            Else
                RaiseMessage.Raise(vbfb, "Sequence Command Error: passed wrong parameter type")
                Return CommandResults.Error
            End If
        End If

        Return CommandResults.Done
    End Function
End Class

Public Class RTAURTCommandLinkedSequenceParameters
    Inherits RTAURTCommandParameters

    Private _LinkedSequenceName
    Public Property LinkedSequenceName As String
        Get
            Return _LinkedSequenceName
        End Get
        Set(value As String)
            _LinkedSequenceName = value
        End Set
    End Property

    Public Sub New(ByVal nm As String)
        _LinkedSequenceName = nm
    End Sub

    Public Overrides Function GetCommand() As RTAURTCommand
        Return CommandFactory.MakeLinkedSequenceCommand
    End Function

    Public Overrides Function TryGetRecursiveName() As String
        Return _LinkedSequenceName
    End Function

End Class
