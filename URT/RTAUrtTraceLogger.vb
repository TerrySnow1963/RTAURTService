'Imports RTAInterfaces
Imports URT

Public Interface IRTAUrtMessageLog
    Sub Write(ByVal message As String)
    Sub ClearLog()
    Function GetLastMessage() As String
    ReadOnly Property Count() As Integer
End Interface

Public Class RTAURTNullLogger
    Implements IRTAUrtMessageLog

    Private ReadOnly Property IRTAUrtMessageLog_Count As Integer Implements IRTAUrtMessageLog.Count
        Get
            Return 0
        End Get
    End Property

    Public Sub ClearLog() Implements IRTAUrtMessageLog.ClearLog
    End Sub

    Public Sub Write(message As String) Implements IRTAUrtMessageLog.Write
    End Sub

    Public Function GetLastMessage() As String Implements IRTAUrtMessageLog.GetLastMessage
        Return String.Empty
    End Function
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

    Public ReadOnly Property Count As Integer Implements IRTAUrtMessageLog.Count
        Get
            Return _messageList.Count
        End Get
    End Property


    Public Function GetLastMessage() As String Implements IRTAUrtMessageLog.GetLastMessage
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
