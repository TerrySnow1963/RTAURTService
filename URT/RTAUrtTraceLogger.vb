'Imports RTAInterfaces
Imports URT

Public Interface IRTAUrtMessageLog
    Sub Write(ByVal message As String)
    Sub ClearLog()
End Interface

Public Class RTAURTNullLogger
    Implements IRTAUrtMessageLog

    Public Sub ClearLog() Implements IRTAUrtMessageLog.ClearLog
    End Sub

    Public Sub Write(message As String) Implements IRTAUrtMessageLog.Write
    End Sub
End Class

Public Class RTAUrtTraceLogger
    Implements IRTAUrtMessageLog

    Private _messageList As List(Of String)

    Public Sub New()
        _messageList = New List(Of String)
    End Sub

    Public Sub Write(message As String) Implements IRTAUrtMessageLog.Write
        If String.IsNullOrEmpty(message) Then Return

        _messageList.Add(message)
        Trace.WriteLine(message)
    End Sub

    Public ReadOnly Property Count As Integer
        Get
            Return _messageList.Count
        End Get
    End Property


    Public Function GetLastMessage() As String
        Dim retString As String = String.Empty

        If _messageList.Count > 0 Then
            retString = _messageList.Last
        End If

        Return retString

    End Function

    Public Sub ClearLog() Implements IRTAUrtMessageLog.ClearLog
        _messageList.Clear()
    End Sub
End Class
