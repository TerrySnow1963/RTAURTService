Imports RTAURTService
Imports RTAInterfaces
Imports URT


Public Class SequenceFactory
    Public Shared Function MakeSequence() As RTAURTCommandSequenceParameters

        Dim uniqueName As String = GetNextName()
        Return New RTAURTCommandSequenceParameters(uniqueName)

    End Function

    Public Shared Function GetNextName() As String
        Return "Seq1"
    End Function

End Class


Public Class RTAURTCommandSequence
    Implements RTAURTCommand

    Public Sub Execute(vbfb As URTVBFunctionBlock, ByRef params As RTAURTCommandParameters) Implements RTAURTCommand.Execute

        Dim seqParams As RTAURTCommandSequenceParameters

        If Not params Is Nothing Then
            seqParams = TryCast(params, RTAURTCommandSequenceParameters)
            If Not seqParams Is Nothing Then
                'ToDo need to check for recursion
                For Each cmdParams In seqParams.GetCommandList
                    cmdParams.GetCommand.Execute(vbfb, cmdParams)
                Next
                If seqParams.GetCommandList.Count = 0 Then
                    RaiseMessage.Raise(vbfb, RTAURTCommandErrorMessages.SequenceErrorNoCommands)
                End If
            Else
                RaiseMessage.Raise(vbfb, "Sequence Command Error: passed wrong parameter type")
            End If
        End If

    End Sub
End Class

Public Class RTAURTCommandSequenceParameters
    Inherits RTAURTCommandParameters

    Private _list As List(Of IRTAURTCommandParameters)
    Public Sub New(ByVal nm As String)
        _list = New List(Of IRTAURTCommandParameters)
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


