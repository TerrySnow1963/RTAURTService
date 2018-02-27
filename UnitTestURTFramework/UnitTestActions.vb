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

<TestClass()> Public Class UnitTestCommands
    Private Const _NUM_SCRIPT_PARAMS As Integer = 8
    <TestMethod()> Public Sub TestCommandConnect()
        Dim urtVBFB = New URTVBQALSimpleScript.URTVBQALSimpleScript

        Dim ConnectCommand As New RTAURTCommandConnect
        Dim ConnectCommandParams As New RTAURTCommandConnectParameters(True)

        ConnectCommand.Execute(urtVBFB, ConnectCommandParams)

        Assert.AreEqual(_NUM_SCRIPT_PARAMS, urtVBFB.GetElements.Count)
        Assert.AreEqual(1, urtVBFB.NumberOfConnectionsExecutions)
    End Sub

    <TestMethod()> Public Sub TestCommandConnectAnd1Execute()
        Dim urtVBFB = New URTVBQALSimpleScript.URTVBQALSimpleScript

        Dim ConnectCommand As New RTAURTCommandConnect
        Dim ConnectCommandParams As New RTAURTCommandConnectParameters(True)

        Dim ExecuteCommand As New RTAURTCommandExecute
        Dim ExecuteCommandParams As New RTAURTCommandExecuteParameters()

        ConnectCommand.Execute(urtVBFB, ConnectCommandParams)

        Assert.AreEqual(_NUM_SCRIPT_PARAMS, urtVBFB.GetElements.Count)
        ExecuteCommand.Execute(urtVBFB, ExecuteCommandParams)

        Assert.AreEqual(1, urtVBFB.NumberOfExecuteExecutions)
    End Sub

    <TestMethod()> Public Sub TestCommandConnectAnd5Executes()
        Dim urtVBFB = New URTVBQALSimpleScript.URTVBQALSimpleScript

        Dim ConnectCommand As New RTAURTCommandConnect
        Dim ConnectCommandParams As New RTAURTCommandConnectParameters(True)

        Dim NumberOfExecutes As Integer = 5
        Dim ExecuteCommand As New RTAURTCommandExecute()
        Dim ExecuteCommandParams As New RTAURTCommandExecuteParameters(NumberOfExecutes)

        ConnectCommand.Execute(urtVBFB, ConnectCommandParams)

        Dim outCounter As ConInt = urtVBFB.GetElement(ScriptVarNames.outCounter)
        Dim InitCounterVal As Integer = 3
        outCounter.Val = InitCounterVal

        Assert.AreEqual(_NUM_SCRIPT_PARAMS, urtVBFB.GetElements.Count)
        ExecuteCommand.Execute(urtVBFB, ExecuteCommandParams)

        Assert.AreEqual(NumberOfExecutes, urtVBFB.NumberOfExecuteExecutions)
        Assert.AreEqual(1, urtVBFB.NumberOfConnectionsExecutions)
        Assert.AreEqual(NumberOfExecutes + InitCounterVal, outCounter.Val)
    End Sub

    <TestMethod()> Public Sub TestCommandSetValuesFor1IntValue()
        Dim urtVBFB = New URTVBQALSimpleScript.URTVBQALSimpleScript

        Dim ConnectCommand As New RTAURTCommandConnect
        Dim ConnectCommandParams As New RTAURTCommandConnectParameters(True)

        Dim NumberOfExecutes As Integer = 3
        Dim ExecuteCommand As New RTAURTCommandExecute()
        Dim ExecuteCommandParams As New RTAURTCommandExecuteParameters(NumberOfExecutes)

        ConnectCommand.Execute(urtVBFB, ConnectCommandParams)

        Dim outCounter As ConInt = urtVBFB.GetElement(ScriptVarNames.outCounter)
        Dim InitCounterVal As Integer = 7

        Dim SetValuesCommand As New RTAURTCommandSetValues()
        Dim SetValuesCommandParams As New RTAURTCommandSetValuesParameters()
        SetValuesCommandParams.AddValue("outCounter", InitCounterVal)

        SetValuesCommand.Execute(urtVBFB, SetValuesCommandParams)

        Assert.AreEqual(_NUM_SCRIPT_PARAMS, urtVBFB.GetElements.Count)
        ExecuteCommand.Execute(urtVBFB, ExecuteCommandParams)

        Assert.AreEqual(NumberOfExecutes, urtVBFB.NumberOfExecuteExecutions)
        Assert.AreEqual(1, urtVBFB.NumberOfConnectionsExecutions)
        Assert.AreEqual(InitCounterVal + NumberOfExecutes, outCounter.Val)
    End Sub

    <TestMethod()> Public Sub TestCommandSetValuesFor2ValuesIntAndBool()
        Dim urtVBFB = New URTVBQALSimpleScript.URTVBQALSimpleScript

        Dim ConnectCommand As New RTAURTCommandConnect
        Dim ConnectCommandParams As New RTAURTCommandConnectParameters(True)

        Dim NumberOfExecutes As Integer = 3
        Dim ExecuteCommand As New RTAURTCommandExecute()
        Dim ExecuteCommandParams As New RTAURTCommandExecuteParameters(NumberOfExecutes)

        ConnectCommand.Execute(urtVBFB, ConnectCommandParams)

        Dim outCounter As ConInt = urtVBFB.GetElement(ScriptVarNames.outCounter)
        Dim InitCounterVal As Integer = 7

        Dim SetValuesCommand As New RTAURTCommandSetValues()
        Dim SetValuesCommandParams As New RTAURTCommandSetValuesParameters()
        SetValuesCommandParams.AddValue(ScriptVarNames.outCounter, InitCounterVal)
        SetValuesCommandParams.AddValue(ScriptVarNames.inUseUpCounter, False)

        SetValuesCommand.Execute(urtVBFB, SetValuesCommandParams)

        Assert.AreEqual(_NUM_SCRIPT_PARAMS, urtVBFB.GetElements.Count)
        ExecuteCommand.Execute(urtVBFB, ExecuteCommandParams)

        Assert.AreEqual(NumberOfExecutes, urtVBFB.NumberOfExecuteExecutions)
        Assert.AreEqual(1, urtVBFB.NumberOfConnectionsExecutions)
        Assert.AreEqual(InitCounterVal - NumberOfExecutes, outCounter.Val)
    End Sub

    <TestMethod()> Public Sub TestCommandSetValuesForNothingNameRaisesMessage()
        Dim logger As RTAUrtTraceLogger = New RTAUrtTraceLogger
        Dim urtVBFB = New URTVBQALSimpleScript.URTVBQALSimpleScript(logger)

        Dim newElementValue As Single = 10.3!

        Dim SetValuesCommand As New RTAURTCommandSetValues()
        Dim SetValuesCommandParams As New RTAURTCommandSetValuesParameters()
        SetValuesCommandParams.AddValue(Nothing, newElementValue)

        SetValuesCommand.Execute(urtVBFB, SetValuesCommandParams)

        Assert.AreEqual(1, logger.Count)
        Assert.AreEqual(SetValuesCommand.ErrorMessages.CantFindElement, logger.GetLastMessage)

    End Sub

    <TestMethod()> Public Sub TestCommandSetValuesForNothingNameRaisesMessageAfterConnect()
        Dim logger As RTAUrtTraceLogger = New RTAUrtTraceLogger
        Dim urtVBFB = New URTVBQALSimpleScript.URTVBQALSimpleScript(logger)

        Dim ConnectCommand As New RTAURTCommandConnect
        Dim ConnectCommandParams As New RTAURTCommandConnectParameters(True)
        ConnectCommand.Execute(urtVBFB, ConnectCommandParams)

        Dim newElementValue As Single = 10.3!

        Dim SetValuesCommand As New RTAURTCommandSetValues()
        Dim SetValuesCommandParams As New RTAURTCommandSetValuesParameters()
        SetValuesCommandParams.AddValue(Nothing, newElementValue)

        SetValuesCommand.Execute(urtVBFB, SetValuesCommandParams)

        Assert.AreEqual(1, logger.Count)
        Assert.AreEqual(SetValuesCommand.ErrorMessages.CantFindElement, logger.GetLastMessage)

    End Sub

    <TestMethod()> Public Sub TestCommandSetValuesForUnknownNameRaisesMessageAfterConnect()
        Dim logger As RTAUrtTraceLogger = New RTAUrtTraceLogger
        Dim urtVBFB = New URTVBQALSimpleScript.URTVBQALSimpleScript(logger)

        Dim ConnectCommand As New RTAURTCommandConnect
        Dim ConnectCommandParams As New RTAURTCommandConnectParameters(True)
        ConnectCommand.Execute(urtVBFB, ConnectCommandParams)

        Dim newElementValue As Single = 10.3!

        Dim SetValuesCommand As New RTAURTCommandSetValues()
        Dim SetValuesCommandParams As New RTAURTCommandSetValuesParameters()
        SetValuesCommandParams.AddValue("Does not exist", newElementValue)

        SetValuesCommand.Execute(urtVBFB, SetValuesCommandParams)

        Assert.AreEqual(1, logger.Count)
        Assert.AreEqual(SetValuesCommand.ErrorMessages.CantFindElement, logger.GetLastMessage)

    End Sub

    <TestMethod()> Public Sub TestCommandSetValuesSingleWrongValueTypeGivesError()
        Dim logger As RTAUrtTraceLogger = New RTAUrtTraceLogger
        Dim urtVBFB = New URTVBQALSimpleScript.URTVBQALSimpleScript(logger)

        Dim ConnectCommand As New RTAURTCommandConnect
        Dim ConnectCommandParams As New RTAURTCommandConnectParameters(True)
        ConnectCommand.Execute(urtVBFB, ConnectCommandParams)

        Dim newElementValue As String = "this cannot be a single"

        Dim SetValuesCommand As New RTAURTCommandSetValues()
        Dim SetValuesCommandParams As New RTAURTCommandSetValuesParameters()
        SetValuesCommandParams.AddValue(ScriptVarNames.inFloat1, newElementValue)

        SetValuesCommand.Execute(urtVBFB, SetValuesCommandParams)

        Assert.AreEqual(1, logger.Count)
        Assert.AreEqual(SetValuesCommand.ErrorMessages.ValueFailsToConvert, logger.GetLastMessage)

    End Sub

    <TestMethod()> Public Sub TestCommandSetValuesSingleArrayWrongValueTypeGivesError()
        Dim logger As RTAUrtTraceLogger = New RTAUrtTraceLogger
        Dim urtVBFB = New URTVBQALSimpleScript.URTVBQALSimpleScript(logger)

        Dim ConnectCommand As New RTAURTCommandConnect
        Dim ConnectCommandParams As New RTAURTCommandConnectParameters(True)
        ConnectCommand.Execute(urtVBFB, ConnectCommandParams)

        Dim newElementValue As String = "this cannot be a single"

        Dim SetValuesCommand As New RTAURTCommandSetValues()
        Dim SetValuesCommandParams As New RTAURTCommandSetValuesParameters()
        SetValuesCommandParams.AddValue(ScriptVarNames.inArrayFloat, newElementValue, 0)

        SetValuesCommand.Execute(urtVBFB, SetValuesCommandParams)

        Assert.AreEqual(1, logger.Count)
        Assert.AreEqual(SetValuesCommand.ErrorMessages.ValueFailsToConvert, logger.GetLastMessage)

    End Sub

    <TestMethod()> Public Sub TestCommandSetValuesFor1IntArrayValue()
        Dim urtVBFB = New URTVBQALSimpleScript.URTVBQALSimpleScript

        Dim ConnectCommand As New RTAURTCommandConnect
        Dim ConnectCommandParams As New RTAURTCommandConnectParameters(True)
        ConnectCommand.Execute(urtVBFB, Nothing)

        Dim OutFloat2 As ConFloat = urtVBFB.GetElement(ScriptVarNames.outFloat2)

        Dim newElementValue As Single = 10.3!

        Dim SetValuesCommand As New RTAURTCommandSetValues()
        Dim SetValuesCommandParams As New RTAURTCommandSetValuesParameters()
        SetValuesCommandParams.AddValue(ScriptVarNames.inArrayFloat, newElementValue)

        SetValuesCommand.Execute(urtVBFB, SetValuesCommandParams)

        Dim ExecuteCommand As New RTAURTCommandExecute()
        Dim ExecuteCommandParams As New RTAURTCommandExecuteParameters()
        Assert.AreEqual(_NUM_SCRIPT_PARAMS, urtVBFB.GetElements.Count)
        ExecuteCommand.Execute(urtVBFB, ExecuteCommandParams)

        Assert.AreEqual(newElementValue + 1.0! * 3.0!, OutFloat2.Val)
    End Sub

    Private Sub Connect(ByVal urtVBFB As URTVBQALSimpleScript.URTVBQALSimpleScript)
        Dim ConnectCommand As New RTAURTCommandConnect
        Dim ConnectCommandParams As New RTAURTCommandConnectParameters(True)
        ConnectCommand.Execute(urtVBFB, ConnectCommandParams)
    End Sub

    Private Sub Execute(ByVal urtVBFB As URTVBQALSimpleScript.URTVBQALSimpleScript, ByVal numberOfExecutions As Integer)
        Dim ExecuteCommand As New RTAURTCommandExecute()
        Dim ExecuteCommandParams As New RTAURTCommandExecuteParameters(numberOfExecutions)
        Assert.AreEqual(_NUM_SCRIPT_PARAMS, urtVBFB.GetElements.Count)
        ExecuteCommand.Execute(urtVBFB, ExecuteCommandParams)
    End Sub

    <TestMethod()> Public Sub TestCommandEnableHistoryFor1Scalar()

        Dim logger As RTAUrtTraceLogger = New RTAUrtTraceLogger
        Dim urtVBFB = New URTVBQALSimpleScript.URTVBQALSimpleScript(logger)

        Connect(urtVBFB)

        Dim EnableHistoryCommand As New RTAURTCommandEnableHistory
        Dim history As RTAUrtTraceHistory = New RTAUrtTraceHistory
        Dim EnableHistoryCommandParams As New RTAURTCommandEnableHistoryParameters(history)

        EnableHistoryCommandParams.EnableHistoryFor(ScriptVarNames.outCounter)
        EnableHistoryCommand.Execute(urtVBFB, EnableHistoryCommandParams)

        Execute(urtVBFB, 1)

        Assert.AreEqual(1, history.CountTimeStamps())
        Assert.AreEqual(1, history.CountTags())
    End Sub

    <TestMethod()> Public Sub TestCommandEnableHistoryFor1BoolArray()

        Dim logger As RTAUrtTraceLogger = New RTAUrtTraceLogger
        Dim urtVBFB = New URTVBQALSimpleScript.URTVBQALSimpleScript(logger)

        Connect(urtVBFB)

        Dim EnableHistoryCommand As New RTAURTCommandEnableHistory
        Dim history As RTAUrtTraceHistory = New RTAUrtTraceHistory
        Dim EnableHistoryCommandParams As New RTAURTCommandEnableHistoryParameters(history)

        EnableHistoryCommandParams.EnableHistoryFor(ScriptVarNames.inArrayBool)
        EnableHistoryCommand.Execute(urtVBFB, EnableHistoryCommandParams)

        Execute(urtVBFB, 1)

        Assert.AreEqual(1, history.CountTimeStamps())
        Assert.AreEqual(1, history.CountTags())
    End Sub

    <TestMethod()> Public Sub TestCommandEnableHistoryFor2Items2Executes()

        Dim logger As RTAUrtTraceLogger = New RTAUrtTraceLogger
        Dim urtVBFB = New URTVBQALSimpleScript.URTVBQALSimpleScript(logger)

        Connect(urtVBFB)

        Dim EnableHistoryCommand As New RTAURTCommandEnableHistory
        Dim history As RTAUrtTraceHistory = New RTAUrtTraceHistory
        Dim EnableHistoryCommandParams As New RTAURTCommandEnableHistoryParameters(history)

        EnableHistoryCommandParams.EnableHistoryFor(ScriptVarNames.inArrayBool)
        EnableHistoryCommandParams.EnableHistoryFor(ScriptVarNames.outCounter)
        EnableHistoryCommandParams.EnableHistoryFor(ScriptVarNames.inArrayFloat)
        EnableHistoryCommand.Execute(urtVBFB, EnableHistoryCommandParams)

        Execute(urtVBFB, 2)

        Assert.AreEqual(2, history.CountTimeStamps())
        Assert.AreEqual(3, history.CountTags())
    End Sub

    <TestMethod()> Public Sub TestCommandMessageCommandWrite1Message()

        Dim logger As RTAUrtTraceLogger = New RTAUrtTraceLogger
        Dim urtVBFB = New URTVBQALSimpleScript.URTVBQALSimpleScript(logger)

        Connect(urtVBFB)

        Dim testMsg As String = "This is a test Message"
        Dim MessageCommandParams = New RTAURTCommandMessageParameters(testMsg)

        MessageCommandParams.GetCommand.Execute(urtVBFB, MessageCommandParams)

        Assert.AreEqual(1, logger.Count)
        Assert.AreEqual(testMsg, logger.GetLastMessage)
    End Sub

    <TestMethod()> Public Sub TestCommandClearLogsClearsOnlyMessages()

        Dim logger As RTAUrtTraceLogger = New RTAUrtTraceLogger
        Dim urtVBFB = New URTVBQALSimpleScript.URTVBQALSimpleScript(logger)

        Connect(urtVBFB)

        Dim EnableHistoryCommand As New RTAURTCommandEnableHistory
        Dim history As RTAUrtTraceHistory = New RTAUrtTraceHistory
        Dim EnableHistoryCommandParams As New RTAURTCommandEnableHistoryParameters(history)
        EnableHistoryCommandParams.EnableHistoryFor(ScriptVarNames.inArrayFloat)
        EnableHistoryCommand.Execute(urtVBFB, EnableHistoryCommandParams)

        Dim testMsg As String = "Test Message from CommandMessage"
        Dim MessageCommand As New RTAURTCommandMessage
        Dim MessageCommandParams As New RTAURTCommandMessageParameters(testMsg)

        MessageCommand.Execute(urtVBFB, MessageCommandParams)

        Assert.AreEqual(1, logger.Count)
        Assert.AreEqual(testMsg, logger.GetLastMessage)

        Execute(urtVBFB, 2)
        Assert.AreEqual(1, logger.Count)
        Assert.AreEqual(2, history.CountTimeStamps())

        Dim ClearLogCommand As New RTAURTCommandClearLogs
        Dim ClearMessageLogCommandParams As New RTAURTCommandClearLogsParameters(RTAURTCommandClearLogs.Logs.MessageLog)
        ClearLogCommand.Execute(urtVBFB, ClearMessageLogCommandParams)
        Assert.AreEqual(0, logger.Count)
        Assert.AreEqual(2, history.CountTimeStamps())
    End Sub

    <TestMethod()> Public Sub TestCommandClearLogsClearsOnlyHistory()

        Dim logger As RTAUrtTraceLogger = New RTAUrtTraceLogger
        Dim urtVBFB = New URTVBQALSimpleScript.URTVBQALSimpleScript(logger)

        Connect(urtVBFB)

        Dim EnableHistoryCommand As New RTAURTCommandEnableHistory
        Dim history As RTAUrtTraceHistory = New RTAUrtTraceHistory
        Dim EnableHistoryCommandParams As New RTAURTCommandEnableHistoryParameters(history)
        EnableHistoryCommandParams.EnableHistoryFor(ScriptVarNames.inArrayFloat)
        EnableHistoryCommand.Execute(urtVBFB, EnableHistoryCommandParams)

        Dim testMsg As String = "Test Message from CommandMessage"
        Dim MessageCommand As New RTAURTCommandMessage
        Dim MessageCommandParams As New RTAURTCommandMessageParameters(testMsg)

        MessageCommand.Execute(urtVBFB, MessageCommandParams)

        Assert.AreEqual(1, logger.Count)
        Assert.AreEqual(testMsg, logger.GetLastMessage)

        Execute(urtVBFB, 2)
        Assert.AreEqual(1, logger.Count)
        Assert.AreEqual(2, history.CountTimeStamps())

        Dim ClearLogCommand As New RTAURTCommandClearLogs
        Dim ClearMessageLogCommandParams As New RTAURTCommandClearLogsParameters(RTAURTCommandClearLogs.Logs.HistoryLog)
        ClearLogCommand.Execute(urtVBFB, ClearMessageLogCommandParams)
        Assert.AreEqual(1, logger.Count)
        Assert.AreEqual(0, history.CountTimeStamps())
    End Sub

    <TestMethod()> Public Sub TestCommandClearLogsClearsMessageAndHistory()

        Dim logger As RTAUrtTraceLogger = New RTAUrtTraceLogger
        Dim urtVBFB = New URTVBQALSimpleScript.URTVBQALSimpleScript(logger)

        Connect(urtVBFB)

        Dim EnableHistoryCommand As New RTAURTCommandEnableHistory
        Dim history As RTAUrtTraceHistory = New RTAUrtTraceHistory
        Dim EnableHistoryCommandParams As New RTAURTCommandEnableHistoryParameters(history)
        EnableHistoryCommandParams.EnableHistoryFor(ScriptVarNames.inArrayFloat)
        EnableHistoryCommand.Execute(urtVBFB, EnableHistoryCommandParams)

        Dim testMsg As String = "Test Message from CommandMessage"
        Dim MessageCommand As New RTAURTCommandMessage
        Dim MessageCommandParams As New RTAURTCommandMessageParameters(testMsg)

        MessageCommand.Execute(urtVBFB, MessageCommandParams)

        Assert.AreEqual(1, logger.Count)
        Assert.AreEqual(testMsg, logger.GetLastMessage)

        Execute(urtVBFB, 2)
        Assert.AreEqual(1, logger.Count)
        Assert.AreEqual(2, history.CountTimeStamps())

        Dim ClearLogCommand As New RTAURTCommandClearLogs
        Dim ClearMessageLogCommandParams As New RTAURTCommandClearLogsParameters(RTAURTCommandClearLogs.Logs.MessageAndHistory)
        ClearLogCommand.Execute(urtVBFB, ClearMessageLogCommandParams)
        Assert.AreEqual(0, logger.Count)
        Assert.AreEqual(0, history.CountTimeStamps())
    End Sub

End Class