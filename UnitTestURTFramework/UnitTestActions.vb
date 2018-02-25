Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports URT
Imports RTAURTService
Imports UrtTlbLib
Imports RTAInterfaces
Imports URTVBQALSimpleScript

Public Class ScriptVarNames
    Public Shared inFloat1 As String = "inFloat1"
    Public Shared inFloat2 As String = "inFloat2"
    Public Shared outFloat1 As String = "outFloat1"
    Public Shared outFloat2 As String = "outFloat2"
    Public Shared inSize As String = "inSize"
    Public Shared outCounter As String = "outCounter"
    Public Shared inUseUpCounter As String = "inUseUpCounter"
    Public Shared inArrayBool As String = "inArrayBool"
    Public Shared inArrayFloat As String = "inArrayFloat"
End Class

<TestClass()> Public Class UnitTestActions
    Private Const _NUM_SCRIPT_PARAMS As Integer = 8
    <TestMethod()> Public Sub TestActionConnect()
        Dim urtVBFB = New URTVBQALSimpleScript.URTVBQALSimpleScript

        Dim ConnectAction As New RTAURTActionConnect
        Dim ConnectActionParams As New RTAURTActionConnectParameters(True)

        ConnectAction.Execute(urtVBFB, ConnectActionParams)

        Assert.AreEqual(_NUM_SCRIPT_PARAMS, urtVBFB.GetElements.Count)
        Assert.AreEqual(1, urtVBFB.NumberOfConnectionsExecutions)
    End Sub

    <TestMethod()> Public Sub TestActionConnectAnd1Execute()
        Dim urtVBFB = New URTVBQALSimpleScript.URTVBQALSimpleScript

        Dim ConnectAction As New RTAURTActionConnect
        Dim ConnectActionParams As New RTAURTActionConnectParameters(True)

        Dim ExecuteAction As New RTAURTActionExecute
        Dim ExecuteActionParams As New RTAURTActionExecuteParameters()

        ConnectAction.Execute(urtVBFB, ConnectActionParams)

        Assert.AreEqual(_NUM_SCRIPT_PARAMS, urtVBFB.GetElements.Count)
        ExecuteAction.Execute(urtVBFB, ExecuteActionParams)

        Assert.AreEqual(1, urtVBFB.NumberOfExecuteExecutions)
    End Sub

    <TestMethod()> Public Sub TestActionConnectAnd5Executes()
        Dim urtVBFB = New URTVBQALSimpleScript.URTVBQALSimpleScript

        Dim ConnectAction As New RTAURTActionConnect
        Dim ConnectActionParams As New RTAURTActionConnectParameters(True)

        Dim NumberOfExecutes As Integer = 5
        Dim ExecuteAction As New RTAURTActionExecute()
        Dim ExecuteActionParams As New RTAURTActionExecuteParameters(NumberOfExecutes)

        ConnectAction.Execute(urtVBFB, ConnectActionParams)

        Dim outCounter As ConInt = urtVBFB.GetElement(ScriptVarNames.outCounter)
        Dim InitCounterVal As Integer = 3
        outCounter.Val = InitCounterVal

        Assert.AreEqual(_NUM_SCRIPT_PARAMS, urtVBFB.GetElements.Count)
        ExecuteAction.Execute(urtVBFB, ExecuteActionParams)

        Assert.AreEqual(NumberOfExecutes, urtVBFB.NumberOfExecuteExecutions)
        Assert.AreEqual(1, urtVBFB.NumberOfConnectionsExecutions)
        Assert.AreEqual(NumberOfExecutes + InitCounterVal, outCounter.Val)
    End Sub

    <TestMethod()> Public Sub TestActionSetValuesFor1IntValue()
        Dim urtVBFB = New URTVBQALSimpleScript.URTVBQALSimpleScript

        Dim ConnectAction As New RTAURTActionConnect
        Dim ConnectActionParams As New RTAURTActionConnectParameters(True)

        Dim NumberOfExecutes As Integer = 3
        Dim ExecuteAction As New RTAURTActionExecute()
        Dim ExecuteActionParams As New RTAURTActionExecuteParameters(NumberOfExecutes)

        ConnectAction.Execute(urtVBFB, ConnectActionParams)

        Dim outCounter As ConInt = urtVBFB.GetElement(ScriptVarNames.outCounter)
        Dim InitCounterVal As Integer = 7

        Dim SetValuesAction As New RTAURTActionSetValues()
        Dim SetValuesActionParams As New RTAURTActionSetValuesParameters()
        SetValuesActionParams.AddValue("outCounter", InitCounterVal)

        SetValuesAction.Execute(urtVBFB, SetValuesActionParams)

        Assert.AreEqual(_NUM_SCRIPT_PARAMS, urtVBFB.GetElements.Count)
        ExecuteAction.Execute(urtVBFB, ExecuteActionParams)

        Assert.AreEqual(NumberOfExecutes, urtVBFB.NumberOfExecuteExecutions)
        Assert.AreEqual(1, urtVBFB.NumberOfConnectionsExecutions)
        Assert.AreEqual(InitCounterVal + NumberOfExecutes, outCounter.Val)
    End Sub

    <TestMethod()> Public Sub TestActionSetValuesFor2ValuesIntAndBool()
        Dim urtVBFB = New URTVBQALSimpleScript.URTVBQALSimpleScript

        Dim ConnectAction As New RTAURTActionConnect
        Dim ConnectActionParams As New RTAURTActionConnectParameters(True)

        Dim NumberOfExecutes As Integer = 3
        Dim ExecuteAction As New RTAURTActionExecute()
        Dim ExecuteActionParams As New RTAURTActionExecuteParameters(NumberOfExecutes)

        ConnectAction.Execute(urtVBFB, ConnectActionParams)

        Dim outCounter As ConInt = urtVBFB.GetElement(ScriptVarNames.outCounter)
        Dim InitCounterVal As Integer = 7

        Dim SetValuesAction As New RTAURTActionSetValues()
        Dim SetValuesActionParams As New RTAURTActionSetValuesParameters()
        SetValuesActionParams.AddValue(ScriptVarNames.outCounter, InitCounterVal)
        SetValuesActionParams.AddValue(ScriptVarNames.inUseUpCounter, False)

        SetValuesAction.Execute(urtVBFB, SetValuesActionParams)

        Assert.AreEqual(_NUM_SCRIPT_PARAMS, urtVBFB.GetElements.Count)
        ExecuteAction.Execute(urtVBFB, ExecuteActionParams)

        Assert.AreEqual(NumberOfExecutes, urtVBFB.NumberOfExecuteExecutions)
        Assert.AreEqual(1, urtVBFB.NumberOfConnectionsExecutions)
        Assert.AreEqual(InitCounterVal - NumberOfExecutes, outCounter.Val)
    End Sub

    <TestMethod()> Public Sub TestActionSetValuesForNothingNameRaisesMessage()
        Dim logger As RTAUrtTraceLogger = New RTAUrtTraceLogger
        Dim urtVBFB = New URTVBQALSimpleScript.URTVBQALSimpleScript(logger)

        Dim newElementValue As Single = 10.3!

        Dim SetValuesAction As New RTAURTActionSetValues()
        Dim SetValuesActionParams As New RTAURTActionSetValuesParameters()
        SetValuesActionParams.AddValue(Nothing, newElementValue)

        SetValuesAction.Execute(urtVBFB, SetValuesActionParams)

        Assert.AreEqual(1, logger.Count)
        Assert.AreEqual(SetValuesAction.ErrorMessages.CantFindElement, logger.GetLastMessage)

    End Sub

    <TestMethod()> Public Sub TestActionSetValuesForNothingNameRaisesMessageAfterConnect()
        Dim logger As RTAUrtTraceLogger = New RTAUrtTraceLogger
        Dim urtVBFB = New URTVBQALSimpleScript.URTVBQALSimpleScript(logger)

        Dim ConnectAction As New RTAURTActionConnect
        Dim ConnectActionParams As New RTAURTActionConnectParameters(True)
        ConnectAction.Execute(urtVBFB, ConnectActionParams)

        Dim newElementValue As Single = 10.3!

        Dim SetValuesAction As New RTAURTActionSetValues()
        Dim SetValuesActionParams As New RTAURTActionSetValuesParameters()
        SetValuesActionParams.AddValue(Nothing, newElementValue)

        SetValuesAction.Execute(urtVBFB, SetValuesActionParams)

        Assert.AreEqual(1, logger.Count)
        Assert.AreEqual(SetValuesAction.ErrorMessages.CantFindElement, logger.GetLastMessage)

    End Sub

    <TestMethod()> Public Sub TestActionSetValuesForUnknownNameRaisesMessageAfterConnect()
        Dim logger As RTAUrtTraceLogger = New RTAUrtTraceLogger
        Dim urtVBFB = New URTVBQALSimpleScript.URTVBQALSimpleScript(logger)

        Dim ConnectAction As New RTAURTActionConnect
        Dim ConnectActionParams As New RTAURTActionConnectParameters(True)
        ConnectAction.Execute(urtVBFB, ConnectActionParams)

        Dim newElementValue As Single = 10.3!

        Dim SetValuesAction As New RTAURTActionSetValues()
        Dim SetValuesActionParams As New RTAURTActionSetValuesParameters()
        SetValuesActionParams.AddValue("Does not exist", newElementValue)

        SetValuesAction.Execute(urtVBFB, SetValuesActionParams)

        Assert.AreEqual(1, logger.Count)
        Assert.AreEqual(SetValuesAction.ErrorMessages.CantFindElement, logger.GetLastMessage)

    End Sub

    <TestMethod()> Public Sub TestActionSetValuesSingleWrongValueTypeGivesError()
        Dim logger As RTAUrtTraceLogger = New RTAUrtTraceLogger
        Dim urtVBFB = New URTVBQALSimpleScript.URTVBQALSimpleScript(logger)

        Dim ConnectAction As New RTAURTActionConnect
        Dim ConnectActionParams As New RTAURTActionConnectParameters(True)
        ConnectAction.Execute(urtVBFB, ConnectActionParams)

        Dim newElementValue As String = "this cannot be a single"

        Dim SetValuesAction As New RTAURTActionSetValues()
        Dim SetValuesActionParams As New RTAURTActionSetValuesParameters()
        SetValuesActionParams.AddValue(ScriptVarNames.inFloat1, newElementValue)

        SetValuesAction.Execute(urtVBFB, SetValuesActionParams)

        Assert.AreEqual(1, logger.Count)
        Assert.AreEqual(SetValuesAction.ErrorMessages.ValueFailsToConvert, logger.GetLastMessage)

    End Sub

    <TestMethod()> Public Sub TestActionSetValuesSingleArrayWrongValueTypeGivesError()
        Dim logger As RTAUrtTraceLogger = New RTAUrtTraceLogger
        Dim urtVBFB = New URTVBQALSimpleScript.URTVBQALSimpleScript(logger)

        Dim ConnectAction As New RTAURTActionConnect
        Dim ConnectActionParams As New RTAURTActionConnectParameters(True)
        ConnectAction.Execute(urtVBFB, ConnectActionParams)

        Dim newElementValue As String = "this cannot be a single"

        Dim SetValuesAction As New RTAURTActionSetValues()
        Dim SetValuesActionParams As New RTAURTActionSetValuesParameters()
        SetValuesActionParams.AddValue(ScriptVarNames.inArrayFloat, newElementValue, 0)

        SetValuesAction.Execute(urtVBFB, SetValuesActionParams)

        Assert.AreEqual(1, logger.Count)
        Assert.AreEqual(SetValuesAction.ErrorMessages.ValueFailsToConvert, logger.GetLastMessage)

    End Sub

    <TestMethod()> Public Sub TestActionSetValuesFor1IntArrayValue()
        Dim urtVBFB = New URTVBQALSimpleScript.URTVBQALSimpleScript

        Dim ConnectAction As New RTAURTActionConnect
        Dim ConnectActionParams As New RTAURTActionConnectParameters(True)
        ConnectAction.Execute(urtVBFB, Nothing)

        Dim OutFloat2 As ConFloat = urtVBFB.GetElement(ScriptVarNames.outFloat2)

        Dim newElementValue As Single = 10.3!

        Dim SetValuesAction As New RTAURTActionSetValues()
        Dim SetValuesActionParams As New RTAURTActionSetValuesParameters()
        SetValuesActionParams.AddValue(ScriptVarNames.inArrayFloat, newElementValue)

        SetValuesAction.Execute(urtVBFB, SetValuesActionParams)

        Dim ExecuteAction As New RTAURTActionExecute()
        Dim ExecuteActionParams As New RTAURTActionExecuteParameters()
        Assert.AreEqual(_NUM_SCRIPT_PARAMS, urtVBFB.GetElements.Count)
        ExecuteAction.Execute(urtVBFB, ExecuteActionParams)

        Assert.AreEqual(newElementValue + 1.0! * 3.0!, OutFloat2.Val)
    End Sub

    Private Sub Connect(ByVal urtVBFB As URTVBQALSimpleScript.URTVBQALSimpleScript)
        Dim ConnectAction As New RTAURTActionConnect
        Dim ConnectActionParams As New RTAURTActionConnectParameters(True)
        ConnectAction.Execute(urtVBFB, ConnectActionParams)
    End Sub

    Private Sub Execute(ByVal urtVBFB As URTVBQALSimpleScript.URTVBQALSimpleScript, ByVal numberOfExecutions As Integer)
        Dim ExecuteAction As New RTAURTActionExecute()
        Dim ExecuteActionParams As New RTAURTActionExecuteParameters(numberOfExecutions)
        Assert.AreEqual(_NUM_SCRIPT_PARAMS, urtVBFB.GetElements.Count)
        ExecuteAction.Execute(urtVBFB, ExecuteActionParams)
    End Sub

    <TestMethod()> Public Sub TestActionEnableHistoryFor1Scalar()

        Dim logger As RTAUrtTraceLogger = New RTAUrtTraceLogger
        Dim urtVBFB = New URTVBQALSimpleScript.URTVBQALSimpleScript(logger)

        Connect(urtVBFB)

        Dim EnableHistoryAction As New RTAURTActionEnableHistory
        Dim history As RTAUrtTraceHistory = New RTAUrtTraceHistory
        Dim EnableHistoryActionParams As New RTAURTActionEnableHistoryParameters(history)

        EnableHistoryActionParams.EnableHistoryFor(ScriptVarNames.outCounter)
        EnableHistoryAction.Execute(urtVBFB, EnableHistoryActionParams)

        Execute(urtVBFB, 1)

        Assert.AreEqual(1, history.CountTimeStamps())
        Assert.AreEqual(1, history.CountTags())
    End Sub

    <TestMethod()> Public Sub TestActionEnableHistoryFor1BoolArray()

        Dim logger As RTAUrtTraceLogger = New RTAUrtTraceLogger
        Dim urtVBFB = New URTVBQALSimpleScript.URTVBQALSimpleScript(logger)

        Connect(urtVBFB)

        Dim EnableHistoryAction As New RTAURTActionEnableHistory
        Dim history As RTAUrtTraceHistory = New RTAUrtTraceHistory
        Dim EnableHistoryActionParams As New RTAURTActionEnableHistoryParameters(history)

        EnableHistoryActionParams.EnableHistoryFor(ScriptVarNames.inArrayBool)
        EnableHistoryAction.Execute(urtVBFB, EnableHistoryActionParams)

        Execute(urtVBFB, 1)

        Assert.AreEqual(1, history.CountTimeStamps())
        Assert.AreEqual(1, history.CountTags())
    End Sub

    <TestMethod()> Public Sub TestActionEnableHistoryFor2Items2Executes()

        Dim logger As RTAUrtTraceLogger = New RTAUrtTraceLogger
        Dim urtVBFB = New URTVBQALSimpleScript.URTVBQALSimpleScript(logger)

        Connect(urtVBFB)

        Dim EnableHistoryAction As New RTAURTActionEnableHistory
        Dim history As RTAUrtTraceHistory = New RTAUrtTraceHistory
        Dim EnableHistoryActionParams As New RTAURTActionEnableHistoryParameters(history)

        EnableHistoryActionParams.EnableHistoryFor(ScriptVarNames.inArrayBool)
        EnableHistoryActionParams.EnableHistoryFor(ScriptVarNames.outCounter)
        EnableHistoryActionParams.EnableHistoryFor(ScriptVarNames.inArrayFloat)
        EnableHistoryAction.Execute(urtVBFB, EnableHistoryActionParams)

        Execute(urtVBFB, 2)

        Assert.AreEqual(2, history.CountTimeStamps())
        Assert.AreEqual(3, history.CountTags())
    End Sub

    <TestMethod()> Public Sub TestActionMessageCommandWrite1Message()

        Dim logger As RTAUrtTraceLogger = New RTAUrtTraceLogger
        Dim urtVBFB = New URTVBQALSimpleScript.URTVBQALSimpleScript(logger)

        Connect(urtVBFB)

        Dim testMsg As String = "Test Message from CommandMessage"
        Dim MessageAction As New RTAURTActionMessage
        Dim MessageActionParams As New RTAURTActionMessageParameters(testMsg)

        MessageAction.Execute(urtVBFB, MessageActionParams)

        Assert.AreEqual(1, logger.Count)
        Assert.AreEqual(testMsg, logger.GetLastMessage)
    End Sub

    <TestMethod()> Public Sub TestActionClearLogsClearsOnlyMessages()

        Dim logger As RTAUrtTraceLogger = New RTAUrtTraceLogger
        Dim urtVBFB = New URTVBQALSimpleScript.URTVBQALSimpleScript(logger)

        Connect(urtVBFB)

        Dim EnableHistoryAction As New RTAURTActionEnableHistory
        Dim history As RTAUrtTraceHistory = New RTAUrtTraceHistory
        Dim EnableHistoryActionParams As New RTAURTActionEnableHistoryParameters(history)
        EnableHistoryActionParams.EnableHistoryFor(ScriptVarNames.inArrayFloat)
        EnableHistoryAction.Execute(urtVBFB, EnableHistoryActionParams)

        Dim testMsg As String = "Test Message from CommandMessage"
        Dim MessageAction As New RTAURTActionMessage
        Dim MessageActionParams As New RTAURTActionMessageParameters(testMsg)

        MessageAction.Execute(urtVBFB, MessageActionParams)

        Assert.AreEqual(1, logger.Count)
        Assert.AreEqual(testMsg, logger.GetLastMessage)

        Execute(urtVBFB, 2)
        Assert.AreEqual(1, logger.Count)
        Assert.AreEqual(2, history.CountTimeStamps())

        Dim ClearLogAction As New RTAURTActionClearLogs
        Dim ClearMessageLogActionParams As New RTAURTActionClearLogsParameters(RTAURTActionClearLogs.Logs.MessageLog)
        ClearLogAction.Execute(urtVBFB, ClearMessageLogActionParams)
        Assert.AreEqual(0, logger.Count)
        Assert.AreEqual(2, history.CountTimeStamps())
    End Sub

    <TestMethod()> Public Sub TestActionClearLogsClearsOnlyHistory()

        Dim logger As RTAUrtTraceLogger = New RTAUrtTraceLogger
        Dim urtVBFB = New URTVBQALSimpleScript.URTVBQALSimpleScript(logger)

        Connect(urtVBFB)

        Dim EnableHistoryAction As New RTAURTActionEnableHistory
        Dim history As RTAUrtTraceHistory = New RTAUrtTraceHistory
        Dim EnableHistoryActionParams As New RTAURTActionEnableHistoryParameters(history)
        EnableHistoryActionParams.EnableHistoryFor(ScriptVarNames.inArrayFloat)
        EnableHistoryAction.Execute(urtVBFB, EnableHistoryActionParams)

        Dim testMsg As String = "Test Message from CommandMessage"
        Dim MessageAction As New RTAURTActionMessage
        Dim MessageActionParams As New RTAURTActionMessageParameters(testMsg)

        MessageAction.Execute(urtVBFB, MessageActionParams)

        Assert.AreEqual(1, logger.Count)
        Assert.AreEqual(testMsg, logger.GetLastMessage)

        Execute(urtVBFB, 2)
        Assert.AreEqual(1, logger.Count)
        Assert.AreEqual(2, history.CountTimeStamps())

        Dim ClearLogAction As New RTAURTActionClearLogs
        Dim ClearMessageLogActionParams As New RTAURTActionClearLogsParameters(RTAURTActionClearLogs.Logs.HistoryLog)
        ClearLogAction.Execute(urtVBFB, ClearMessageLogActionParams)
        Assert.AreEqual(1, logger.Count)
        Assert.AreEqual(0, history.CountTimeStamps())
    End Sub

    <TestMethod()> Public Sub TestActionClearLogsClearsMessageAndHistory()

        Dim logger As RTAUrtTraceLogger = New RTAUrtTraceLogger
        Dim urtVBFB = New URTVBQALSimpleScript.URTVBQALSimpleScript(logger)

        Connect(urtVBFB)

        Dim EnableHistoryAction As New RTAURTActionEnableHistory
        Dim history As RTAUrtTraceHistory = New RTAUrtTraceHistory
        Dim EnableHistoryActionParams As New RTAURTActionEnableHistoryParameters(history)
        EnableHistoryActionParams.EnableHistoryFor(ScriptVarNames.inArrayFloat)
        EnableHistoryAction.Execute(urtVBFB, EnableHistoryActionParams)

        Dim testMsg As String = "Test Message from CommandMessage"
        Dim MessageAction As New RTAURTActionMessage
        Dim MessageActionParams As New RTAURTActionMessageParameters(testMsg)

        MessageAction.Execute(urtVBFB, MessageActionParams)

        Assert.AreEqual(1, logger.Count)
        Assert.AreEqual(testMsg, logger.GetLastMessage)

        Execute(urtVBFB, 2)
        Assert.AreEqual(1, logger.Count)
        Assert.AreEqual(2, history.CountTimeStamps())

        Dim ClearLogAction As New RTAURTActionClearLogs
        Dim ClearMessageLogActionParams As New RTAURTActionClearLogsParameters(RTAURTActionClearLogs.Logs.MessageAndHistory)
        ClearLogAction.Execute(urtVBFB, ClearMessageLogActionParams)
        Assert.AreEqual(0, logger.Count)
        Assert.AreEqual(0, history.CountTimeStamps())
    End Sub

End Class