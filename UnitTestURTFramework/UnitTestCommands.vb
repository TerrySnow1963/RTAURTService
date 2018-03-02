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

    Private Function TestCommandIsDone(params As IRTAURTCommandParameters, Optional cmdExec As CommandExecutive = Nothing) As CommandExecutive

        If cmdExec Is Nothing Then
            Dim urtVBFB = New URTVBQALSimpleScript.URTVBQALSimpleScript(New RTAUrtTraceLogger)
            cmdExec = New CommandExecutive(urtVBFB)
        End If

        Dim cmdResult As ICommandResult = cmdExec.Invoke(params)
        Assert.AreEqual(RTAURTCommandResultCode.CMD_DONE, cmdResult.GetResultCode)
        Return cmdExec
    End Function

    Private Function TestCommandHasError(params As IRTAURTCommandParameters, expectedErrMsg As String, Optional cmdExec As CommandExecutive = Nothing) As CommandExecutive

        If cmdExec Is Nothing Then
            Dim urtVBFB = New URTVBQALSimpleScript.URTVBQALSimpleScript(New RTAUrtTraceLogger)
            cmdExec = New CommandExecutive(urtVBFB)
        End If

        Dim cmdResult As ICommandResult = cmdExec.Invoke(params)
        Assert.AreEqual(RTAURTCommandResultCode.CMD_ERROR, cmdResult.GetResultCode)
        Assert.AreEqual(1, cmdExec.vbfb.GetLogMessageCount)
        Assert.AreEqual(expectedErrMsg, cmdExec.vbfb.GetLastLogMessage)
        Return cmdExec
    End Function

    Private Sub TestEndOfExecuteVBRun(cmdExec As CommandExecutive, numberOfExecutesExpected As Integer)
        Trace.WriteLine("Started TestEndOfExecuteVBRun")
        Trace.WriteLine("Testing Number of ExecuteVB Calls")
        Assert.AreEqual(numberOfExecutesExpected, cmdExec.vbfb.NumberOfExecuteExecutions)
        Trace.WriteLine("Testing Number of Connection Calls")
        Assert.AreEqual(1, cmdExec.vbfb.NumberOfConnectionsExecutions)
        Trace.WriteLine("Finished TestEndOfExecuteVBRun")
    End Sub

    Private Function StartWithCommandConnect() As CommandExecutive
        Dim params As New RTAURTCommandConnectParameters(True)
        Dim cmdExec As CommandExecutive = TestCommandIsDone(params)

        Trace.WriteLine("Started StartWithCommandConnect")
        Assert.AreEqual(_NUM_SCRIPT_PARAMS, cmdExec.vbfb.GetElements.Count)
        Trace.WriteLine("Finished StartWithCommandConnect")
        Return cmdExec
    End Function

    <TestMethod()> Public Sub TestCommandConnect()
        StartWithCommandConnect()
    End Sub

    <TestMethod()> Public Sub TestCommandConnectAnd1Execute()
        Dim cmdExec As CommandExecutive = StartWithCommandConnect()
        Dim ExecuteCommandParams As New RTAURTCommandExecuteVBParameters()

        TestCommandIsDone(ExecuteCommandParams, cmdExec)

        TestEndOfExecuteVBRun(cmdExec, 1)
    End Sub

    <TestMethod()> Public Sub TestCommandConnectAnd5Executes()
        Dim cmdExec As CommandExecutive = StartWithCommandConnect()

        Dim NumberOfExecutes As Integer = 5
        Dim ExecuteCommandParams As New RTAURTCommandExecuteVBParameters(NumberOfExecutes)

        Dim outCounter As ConInt = cmdExec.vbfb.GetElement(ScriptVarNames.outCounter)
        Dim InitCounterVal As Integer = 3
        outCounter.Val = InitCounterVal

        TestCommandIsDone(ExecuteCommandParams, cmdExec)

        TestEndOfExecuteVBRun(cmdExec, NumberOfExecutes)
        Assert.AreEqual(NumberOfExecutes + InitCounterVal, outCounter.Val)
    End Sub

    <TestMethod()> Public Sub TestCommandSetValuesFor1IntValue()
        Dim cmdExec As CommandExecutive = StartWithCommandConnect()

        Dim NumberOfExecutes As Integer = 3
        Dim ExecuteVBCommandParams As New RTAURTCommandExecuteVBParameters(NumberOfExecutes)

        Dim outCounter As ConInt = cmdExec.vbfb.GetElement(ScriptVarNames.outCounter)
        Dim InitCounterVal As Integer = 7

        Dim SetValuesCommandParams As New RTAURTCommandSetValuesParameters()
        SetValuesCommandParams.AddValue("outCounter", InitCounterVal)

        TestCommandIsDone(SetValuesCommandParams, cmdExec)
        TestCommandIsDone(ExecuteVBCommandParams, cmdExec)

        TestEndOfExecuteVBRun(cmdExec, NumberOfExecutes)
        Assert.AreEqual(InitCounterVal + NumberOfExecutes, outCounter.Val)
    End Sub

    <TestMethod()> Public Sub TestCommandSetValuesFor2ValuesIntAndBool()
        Dim cmdExec As CommandExecutive = StartWithCommandConnect()

        Dim NumberOfExecutes As Integer = 3
        Dim ExecuteVBCommandParams As New RTAURTCommandExecuteVBParameters(NumberOfExecutes)

        Dim outCounter As ConInt = cmdExec.vbfb.GetElement(ScriptVarNames.outCounter)
        Dim InitCounterVal As Integer = 7

        Dim SetValuesCommandParams As New RTAURTCommandSetValuesParameters()
        SetValuesCommandParams.AddValue(ScriptVarNames.outCounter, InitCounterVal)
        SetValuesCommandParams.AddValue(ScriptVarNames.inUseUpCounter, False)

        TestCommandIsDone(SetValuesCommandParams, cmdExec)
        TestCommandIsDone(ExecuteVBCommandParams, cmdExec)
        TestEndOfExecuteVBRun(cmdExec, NumberOfExecutes)

        Assert.AreEqual(InitCounterVal - NumberOfExecutes, outCounter.Val)
    End Sub

    <TestMethod()> Public Sub TestCommandSetValuesWithNoConnectForNothingNameRaisesMessage()
        'Dim cmdExec As CommandExecutive = StartWithCommandConnect()

        Dim newElementValue As Single = 10.3!

        Dim SetValuesCommand As New RTAURTCommandSetValues()
        Dim SetValuesCommandParams As New RTAURTCommandSetValuesParameters()
        SetValuesCommandParams.AddValue(Nothing, newElementValue)

        TestCommandHasError(SetValuesCommandParams, SetValuesCommand.ErrorMessages.CantFindElement)
        'TestEndOfExecuteVBRun(cmdExec, 0)
    End Sub

    <TestMethod()> Public Sub TestCommandSetValuesForNothingNameRaisesMessageAfterConnect()
        Dim cmdExec As CommandExecutive = StartWithCommandConnect()

        Dim newElementValue As Single = 10.3!
        Dim SetValuesCommand As New RTAURTCommandSetValues()
        Dim SetValuesCommandParams As New RTAURTCommandSetValuesParameters()
        SetValuesCommandParams.AddValue(Nothing, newElementValue)

        TestCommandHasError(SetValuesCommandParams, SetValuesCommand.ErrorMessages.CantFindElement)

    End Sub

    <TestMethod()> Public Sub TestCommandSetValuesForUnknownNameRaisesMessageAfterConnect()
        Dim cmdExec As CommandExecutive = StartWithCommandConnect()

        Dim newElementValue As Single = 10.3!
        Dim SetValuesCommand As New RTAURTCommandSetValues()
        Dim SetValuesCommandParams As New RTAURTCommandSetValuesParameters()
        SetValuesCommandParams.AddValue(Nothing, newElementValue)
        SetValuesCommandParams.AddValue("Does not exist", newElementValue)

        TestCommandHasError(SetValuesCommandParams, RTAURTCommandSetValues.ErrorMessages.CantFindElement)

    End Sub

    <TestMethod()> Public Sub TestCommandSetValuesSingleWrongValueTypeGivesError()
        Dim cmdExec As CommandExecutive = StartWithCommandConnect()

        Dim newElementValue As String = "this cannot be a single"
        Dim SetValuesCommand As New RTAURTCommandSetValues()
        Dim SetValuesCommandParams As New RTAURTCommandSetValuesParameters()
        SetValuesCommandParams.AddValue(ScriptVarNames.inFloat1, newElementValue)

        TestCommandHasError(SetValuesCommandParams, RTAURTCommandSetValues.ErrorMessages.ValueFailsToConvert, cmdExec)
    End Sub

    <TestMethod()> Public Sub TestCommandSetValuesSingleArrayWrongValueTypeGivesError()
        Dim cmdExec As CommandExecutive = StartWithCommandConnect()

        Dim newElementValue As String = "this cannot be a single"
        Dim SetValuesCommand As New RTAURTCommandSetValues()
        Dim SetValuesCommandParams As New RTAURTCommandSetValuesParameters()
        SetValuesCommandParams.AddValue(ScriptVarNames.inArrayFloat, newElementValue, 0)

        TestCommandHasError(SetValuesCommandParams, RTAURTCommandSetValues.ErrorMessages.ValueFailsToConvert, cmdExec)
    End Sub

    <TestMethod()> Public Sub TestCommandSetValuesFor1IntArrayValue()
        Dim cmdExec As CommandExecutive = StartWithCommandConnect()

        Dim OutFloat2 As ConFloat = cmdExec.vbfb.GetElement(ScriptVarNames.outFloat2)

        Dim newElementValue As Single = 10.3!
        Dim SetValuesCommandParams As New RTAURTCommandSetValuesParameters()
        SetValuesCommandParams.AddValue(ScriptVarNames.inArrayFloat, newElementValue, 0)

        Dim ExecuteVBCommandParams As New RTAURTCommandExecuteVBParameters()

        TestCommandIsDone(SetValuesCommandParams, cmdExec)
        TestCommandIsDone(ExecuteVBCommandParams, cmdExec)

        Assert.AreEqual(newElementValue + 1.0! * 3.0!, OutFloat2.Val)

    End Sub

    Private Sub Connect(ByVal urtVBFB As URTVBQALSimpleScript.URTVBQALSimpleScript)
        Dim ConnectCommand As New RTAURTCommandConnect
        Dim ConnectCommandParams As New RTAURTCommandConnectParameters(True)
        ConnectCommand.Execute(urtVBFB, ConnectCommandParams)
    End Sub

    Private Sub Execute(ByVal urtVBFB As URTVBQALSimpleScript.URTVBQALSimpleScript, ByVal numberOfExecutions As Integer)
        Dim ExecuteCommand As New RTAURTCommandExecuteVB()
        Dim ExecuteCommandParams As New RTAURTCommandExecuteVBParameters(numberOfExecutions)
        Assert.AreEqual(_NUM_SCRIPT_PARAMS, urtVBFB.GetElements.Count)
        Dim cmdResult As ICommandResult
        cmdResult = ExecuteCommand.Execute(urtVBFB, ExecuteCommandParams)
        Assert.AreEqual(RTAURTCommandResultCode.CMD_DONE, cmdResult.GetResultCode)
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

        Dim cmdResult As ICommandResult
        cmdResult = MessageCommandParams.GetCommand.Execute(urtVBFB, MessageCommandParams)

        Assert.AreEqual(1, logger.Count)
        Assert.AreEqual(testMsg, logger.GetLastMessage)
        Assert.AreEqual(RTAURTCommandResultCode.CMD_DONE, cmdResult.GetResultCode)
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
        Dim cmdResult As ICommandResult
        cmdResult = ClearLogCommand.Execute(urtVBFB, ClearMessageLogCommandParams)
        Assert.AreEqual(0, logger.Count)
        Assert.AreEqual(2, history.CountTimeStamps())
        Assert.AreEqual(RTAURTCommandResultCode.CMD_DONE, cmdResult.GetResultCode)
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
        Dim cmdResult As ICommandResult
        cmdResult = ClearLogCommand.Execute(urtVBFB, ClearMessageLogCommandParams)
        Assert.AreEqual(1, logger.Count)
        Assert.AreEqual(0, history.CountTimeStamps())
        Assert.AreEqual(RTAURTCommandResultCode.CMD_DONE, cmdResult.GetResultCode)
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
        Dim cmdResult As ICommandResult
        cmdResult = ClearLogCommand.Execute(urtVBFB, ClearMessageLogCommandParams)
        Assert.AreEqual(0, logger.Count)
        Assert.AreEqual(0, history.CountTimeStamps())
        Assert.AreEqual(RTAURTCommandResultCode.CMD_DONE, cmdResult.GetResultCode)
    End Sub

End Class