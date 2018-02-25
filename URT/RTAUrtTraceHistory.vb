Imports URT
Imports UrtTlbLib
Imports RTAInterfaces
Imports System.IO

Public Interface IRTAUrtHistoryLog
    Sub Historize()
    Sub ClearHistory()
    Function RegisterItem(ByRef treeMember As IUrtTreeMember, ByVal name As String) As Boolean
End Interface

Public Class RTAUrtNullHistory
    Implements IRTAUrtHistoryLog

    Public Sub ClearHistory() Implements IRTAUrtHistoryLog.ClearHistory
    End Sub

    Public Sub Historize() Implements IRTAUrtHistoryLog.Historize
    End Sub

    Public Function RegisterItem(ByRef treeMember As IUrtTreeMember, name As String) As Boolean Implements IRTAUrtHistoryLog.RegisterItem
        Return True
    End Function

End Class

Public Class RTAUrtTraceHistory
    Implements IRTAUrtHistoryLog

    Private _count As Integer
    Private _scalarList As List(Of IRTAUrtData)
    Private _arrayList As List(Of IRTAUrtData)

    Public ReadOnly Property CountTags As Integer
        Get
            Return _scalarList.Count + _arrayList.Count
        End Get
    End Property

    Public Sub New()
        _count = 0
        _scalarList = New List(Of IRTAUrtData)
        _arrayList = New List(Of IRTAUrtData)
    End Sub

    Public Sub Historize() Implements IRTAUrtHistoryLog.Historize

        Dim sr As StringWriter = New StringWriter

        sr.Write(String.Format("History Step ({1}):", _scalarList.Count + _arrayList.Count, _count.ToString))
        For ii = 0 To _scalarList.Count - 1
            sr.Write(_scalarList(ii).Item(0))
            If ii < _scalarList.Count - 2 Then
                sr.Write(",")
            End If
        Next
        Dim arraysize As Integer
        For ii = 0 To _arrayList.Count - 1
            arraysize = CType(_arrayList(ii), IUrtData).Size
            For jj = 0 To arraysize - 1
                sr.Write(_arrayList(ii).Item(jj))
                If jj < arraysize - 2 Then
                    sr.Write(",")
                End If
            Next
            If ii < _scalarList.Count - 2 Then
                sr.Write(",")
            End If
        Next

        Trace.WriteLine(sr.ToString)

        _count += 1
    End Sub

    Public Function RegisterItem(ByRef treeMember As IUrtTreeMember, name As String) As Boolean Implements IRTAUrtHistoryLog.RegisterItem
        Dim irtaData As IRTAUrtData
        Dim theGuid As Guid
        treeMember.Find(name, theGuid, irtaData)
        If irtaData Is Nothing Then Return False

        If CType(irtaData, IUrtData).Size = 1 Then
            _scalarList.Add(irtaData)
        Else
            _arrayList.Add(irtaData)
        End If

        Return True
    End Function

    Public Sub ClearHistory() Implements IRTAUrtHistoryLog.ClearHistory
        _count = 0
    End Sub

    Public ReadOnly Property CountTimeStamps As Integer
        Get
            Return _count
        End Get
    End Property
End Class
