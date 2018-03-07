Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports URT
Imports RTAURTService
Imports URTVBQALSimpleScript


<TestClass()> Public Class UnitTestCommandCallBacks

    Private Sub TestCommandStops(params As IRTAURTCommandParameters, Optional expectedMessages As Integer = 0)
        Dim logger As RTAUrtTraceLogger = New RTAUrtTraceLogger
        Dim vbFB As URTVBQALSimpleScript.URTVBQALSimpleScript = New URTVBQALSimpleScript.URTVBQALSimpleScript(logger)
        Dim callback As ICommandCallback = CommandCallbacks.Stop

        Dim cmdExec = New CommandExecutive(vbFB, callback)

        Dim cmdResult As ICommandResult = cmdExec.Invoke(params)
        Assert.AreEqual(RTAURTCommandResultCode.CMD_STOPPED, cmdResult.GetResultCode)
        Assert.AreEqual(expectedMessages, logger.Count)
    End Sub

    Private Sub TestCommandLimitCountCommandStops(params As IRTAURTCommandParameters,
                                                  commandCallLimit As Integer,
                                                  Optional expectedMessages As Integer = 0)
        Dim logger As RTAUrtTraceLogger = New RTAUrtTraceLogger
        Dim vbFB As URTVBQALSimpleScript.URTVBQALSimpleScript = New URTVBQALSimpleScript.URTVBQALSimpleScript(logger)
        Dim callback As ICommandCallback = CommandCallbacks.LimitCommandCallsTo(commandCallLimit)

        Dim cmdExec = New CommandExecutive(vbFB, callback)

        Dim cmdResult As ICommandResult

        If commandCallLimit < 1 Then commandCallLimit = 1

        Dim counter As Integer = 0

        For counter = 0 To commandCallLimit
            cmdResult = cmdExec.Invoke(params)
            If cmdResult.GetResultCode = RTAURTCommandResultCode.CMD_STOPPED Then Exit For
        Next
        Assert.IsNotNull(cmdResult)
        Assert.AreEqual(RTAURTCommandResultCode.CMD_STOPPED, cmdResult.GetResultCode)
        Assert.AreEqual(counter, commandCallLimit)
        Assert.AreEqual(expectedMessages, logger.Count)
    End Sub


    <TestMethod()> Public Sub TestCommandMessageWithStopCallbackStops()
        Dim testMsg As String = "Test Message"
        Dim params As IRTAURTCommandParameters = New RTAURTCommandMessageParameters(testMsg)

        TestCommandStops(params)
    End Sub

    <TestMethod()> Public Sub TestCommandExecuteWithStopCallbackStops()
        Dim params As IRTAURTCommandParameters = New RTAURTCommandExecuteVBParameters(1)

        TestCommandStops(params)
    End Sub

    <TestMethod()> Public Sub TestCommandConnectWithStopCallbackStops()
        Dim params As IRTAURTCommandParameters = New RTAURTCommandConnectParameters(True)

        TestCommandStops(params)
    End Sub

    <TestMethod()> Public Sub TestCommandSetValuesWithStopCallbackStops()
        Dim params As IRTAURTCommandParameters = New RTAURTCommandSetValuesParameters

        TestCommandStops(params)
    End Sub

    <TestMethod()> Public Sub TestCommandEnableHistorysWithStopCallbackStops()
        Dim params As IRTAURTCommandParameters = New RTAURTCommandEnableHistoryParameters(Nothing)

        TestCommandStops(params)
    End Sub

    <TestMethod()> Public Sub TestCommandSequencesWithStopCallbackStops()
        Dim params As IRTAURTCommandParameters = SequenceFactory.MakeSequence

        TestCommandStops(params)
    End Sub

    <TestMethod()> Public Sub TestCommandClearLogsWithStopCallbackStops()
        Dim params As IRTAURTCommandParameters = New RTAURTCommandClearLogsParameters(RTAURTCommandClearLogs.Logs.HistoryLog)

        TestCommandStops(params)
    End Sub

    <TestMethod()> Public Sub TestCommandMessageWithLimitCommandCountCallbackStops()
        Dim testMsg As String = "Test Message"
        Dim params As IRTAURTCommandParameters = New RTAURTCommandMessageParameters(testMsg)

        TestCommandLimitCountCommandStops(params, 3, 3)
    End Sub

End Class