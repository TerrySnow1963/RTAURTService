Imports RTAURTService
Imports RTAInterfaces
Imports URT


Public Class SequenceFactory
    Public Shared Function MakeSequence() As RTAURTActionSequenceParameters

        Dim uniqueName As String = GetNextName()
        Return New RTAURTActionSequenceParameters(uniqueName)

    End Function

    Public Shared Function GetNextName() As String
        Return "Seq1"
    End Function

End Class


Public Class RTAURTActionSequence
    Implements RTAURTAction

    Public Sub Execute(vbfb As URTVBFunctionBlock, ByRef params As RTAURTActionParameters) Implements RTAURTAction.Execute

        Dim seqParams As RTAURTActionSequenceParameters

        If Not params Is Nothing Then
            seqParams = TryCast(params, RTAURTActionSequenceParameters)
            If Not seqParams Is Nothing Then
                RaiseMessage.Raise(vbfb, RTAURTActionErrorMessages.SequenceErrorNoCommands)
            Else

            End If
        End If

    End Sub
End Class

Public Class RTAURTActionSequenceParameters
    Inherits RTAURTActionParameters

    Private _list As List(Of RTAURTAction)
    Public Sub New(ByVal nm As String)
        _list = New List(Of RTAURTAction)
    End Sub

    Public Overrides Function GetCommand() As RTAURTAction
        Return CommandFactory.MakeSequenceCommand
    End Function

    Public Sub AddCommandParameters(params As RTAURTActionMessageParameters)
        _list.Add(params)
    End Sub
End Class


