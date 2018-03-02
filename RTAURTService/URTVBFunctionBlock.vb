Option Explicit On
Option Strict On
Imports URT
Imports UrtTlbLib

Public MustInherit Class URTVBFunctionBlock
    Inherits CUrtFBBase
    Private _theScript As CVBScriptBase
    'Private _logger As IRTAUrtMessageLog

    Private Sub New()
        '_theScript = New URT.VBScript

        'Dim base As CVBScriptBase = _theScript

        'base.CmpPtr = Me
        '_childElements = New List(Of ConScalarBase)
    End Sub

    Public Sub New(ByVal script As CVBScriptBase, ByVal loggr As IRTAUrtMessageLog, ByVal HistLog As IRTAUrtHistoryLog)
        _theScript = script
        _theScript.CmpPtr = Me

        Logger = loggr
        HistoryLog = HistLog

        NumberOfExecuteExecutions = 0
        NumberOfConnectionsExecutions = 0

    End Sub

    'Default Public Property Index(x As Integer) As Integer Implements IUrtTreeMember.Index
    '    Get
    '        Throw New NotImplementedException()
    '    End Get
    '    Set(value As Integer)
    '        Throw New NotImplementedException()
    '    End Set
    'End Property

    Public ReadOnly Property NumberOfExecuteExecutions As Integer
    Public ReadOnly Property NumberOfConnectionsExecutions As Integer

    Protected MustOverride Sub InternalConnect(ByVal WitInit As Boolean)

    Protected MustOverride Sub InternalExecute(ByVal iCause As Integer, ByVal pITMScheduler As IUrtTreeMember)

    Public Sub Connect(ByVal init As Boolean)
        _NumberOfConnectionsExecutions += 1
        InternalConnect(init)
    End Sub

    Public Sub Execute(ByVal iCause As Integer, ByVal pITMScheduler As IUrtTreeMember)
        _NumberOfExecuteExecutions += 1
        InternalExecute(iCause, pITMScheduler)
        HistoryLog.Historize()
    End Sub

    Public Sub ClearMessageLog()
        Logger.ClearLog()
    End Sub

    Public Sub ClearHistoryLog()
        HistoryLog.ClearHistory()
    End Sub

    Public Function GetLogMessageCount() As Integer
        Return Logger.Count
    End Function

    Public Function GetLastLogMessage() As String
        Return Logger.GetLastMessage
    End Function
End Class
