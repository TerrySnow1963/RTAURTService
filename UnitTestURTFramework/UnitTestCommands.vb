Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports URT
Imports RTAURTService
Imports UrtTlbLib
Imports RTAInterfaces
Imports URTVBQALSimpleScript

Public Class ScriptDataElementNames
    Public Shared inFloat1 As String = "inFloat1"
    Public Shared outFloat1 As String = "outFloat1"
    Public Shared outFloat2 As String = "outFloat2"
    Public Shared inSize As String = "inSize"
    Public Shared outCounter As String = "outCounter"
    Public Shared inUseUpCounter As String = "inUseUpCounter"
    Public Shared inArrayBool As String = "inArrayBool"
    Public Shared inArrayFloat As String = "inArrayFloat"

    Public Shared ReadOnly Property CountScriptDataElements As Integer
        Get
            Return 8
        End Get
    End Property

End Class

Public Class UnitTestCommandsBase

    Protected Function MakeCommandExecutiveWithSimpleScript(ByVal recDepth As Integer) As CommandExecutive
        Assert.IsTrue(recDepth > 0)
        Dim urtVBFB = New URTVBQALSimpleScript.URTVBQALSimpleScript(New RTAUrtTraceLogger)
        Dim cmdExec As CommandExecutive = New CommandExecutive(urtVBFB)
        cmdExec.RecursionDepth = recDepth
        Return cmdExec

    End Function

    Protected Function TestCommandReturnsCode(params As IRTAURTCommandParameters,
                                              expectedCode As RTAURTCommandResultCode,
                                              Optional cmdExec As CommandExecutive = Nothing) As CommandExecutive

        If cmdExec Is Nothing Then
            Dim urtVBFB = New URTVBQALSimpleScript.URTVBQALSimpleScript(New RTAUrtTraceLogger)
            cmdExec = New CommandExecutive(urtVBFB, callBack:=Nothing)
        End If

        Dim cmdResult As ICommandResult = cmdExec.Invoke(params)
        Assert.AreEqual(expectedCode, cmdResult.GetResultCode)
        Return cmdExec
    End Function

    Protected Function TestCommandReturnsDone(params As IRTAURTCommandParameters, Optional cmdExec As CommandExecutive = Nothing) As CommandExecutive

        'If cmdExec Is Nothing Then
        '    Dim urtVBFB = New URTVBQALSimpleScript.URTVBQALSimpleScript(New RTAUrtTraceLogger)
        '    cmdExec = New CommandExecutive(urtVBFB)
        'End If

        'Dim cmdResult As ICommandResult = cmdExec.Invoke(params)
        'Assert.AreEqual(RTAURTCommandResultCode.CMD_DONE, cmdResult.GetResultCode)
        'Return cmdExec
        Return TestCommandReturnsCode(params, RTAURTCommandResultCode.CMD_DONE, cmdExec)
    End Function

    Protected Function TestCommandReturnsStopped(params As IRTAURTCommandParameters, Optional cmdExec As CommandExecutive = Nothing) As CommandExecutive

        'If cmdExec Is Nothing Then
        '    Dim urtVBFB = New URTVBQALSimpleScript.URTVBQALSimpleScript(New RTAUrtTraceLogger)
        '    cmdExec = New CommandExecutive(urtVBFB)
        'End If

        'Dim cmdResult As ICommandResult = cmdExec.Invoke(params)
        'Assert.AreEqual(RTAURTCommandResultCode.CMD_STOPPED, cmdResult.GetResultCode)
        'Return cmdExec

        Return TestCommandReturnsCode(params, RTAURTCommandResultCode.CMD_STOPPED, cmdExec)
    End Function

    Protected Function TestCommandReturnsError(params As IRTAURTCommandParameters, expectedErrMsg As String, Optional cmdExec As CommandExecutive = Nothing) As CommandExecutive

        If cmdExec Is Nothing Then
            Dim urtVBFB = New URTVBQALSimpleScript.URTVBQALSimpleScript(New RTAUrtTraceLogger)
            cmdExec = New CommandExecutive(urtVBFB, callBack:=Nothing)
        End If

        Dim cmdResult As ICommandResult = cmdExec.Invoke(params)
        Assert.AreEqual(RTAURTCommandResultCode.CMD_ERROR, cmdResult.GetResultCode)
        Assert.AreEqual(1, cmdExec.vbfb.GetLogMessageCount)
        Assert.AreEqual(expectedErrMsg, cmdExec.vbfb.GetLastLogMessage)
        Return cmdExec
    End Function

    Protected Sub TestEndOfExecuteVBRun(cmdExec As CommandExecutive, numberOfExecutesExpected As Integer)
        Trace.WriteLine("Started TestEndOfExecuteVBRun")
        Trace.WriteLine("Testing Number of ExecuteVB Calls")
        Assert.AreEqual(numberOfExecutesExpected, cmdExec.vbfb.NumberOfExecuteExecutions)
        Trace.WriteLine("Testing Number of Connection Calls")
        Assert.AreEqual(1, cmdExec.vbfb.NumberOfConnectionsExecutions)
        Trace.WriteLine("Finished TestEndOfExecuteVBRun")
    End Sub

    Protected Function StartWithCommandConnect() As CommandExecutive
        Dim params As New RTAURTCommandConnectParameters(True)
        Dim cmdExec As CommandExecutive = TestCommandReturnsDone(params)

        Trace.WriteLine("Started StartWithCommandConnect")
        Assert.AreEqual(ScriptDataElementNames.CountScriptDataElements, cmdExec.vbfb.GetElements.Count)
        Trace.WriteLine("Finished StartWithCommandConnect")
        Return cmdExec
    End Function

    Protected Sub TestExecuteVBCommandReturnsDone(ByVal cmdExec As CommandExecutive, ByVal numberOfExecutions As Integer)
        Dim params As New RTAURTCommandExecuteVBParameters(numberOfExecutions)
        Dim cmdResult As ICommandResult = cmdExec.Invoke(params)
        Assert.AreEqual(RTAURTCommandResultCode.CMD_DONE, cmdResult.GetResultCode)
    End Sub

    Protected Function TestEnableHistoryCommandReturnsDone(cmdExec As CommandExecutive, vars() As String) As Object
        Dim history As RTAUrtTraceHistory = New RTAUrtTraceHistory
        Dim EnableHistoryCommand As New RTAURTCommandEnableHistory
        Dim EnableHistoryCommandParams As New RTAURTCommandEnableHistoryParameters(history)
        For Each v In vars
            EnableHistoryCommandParams.EnableHistoryFor(v)
        Next
        TestCommandReturnsDone(EnableHistoryCommandParams, cmdExec)
        Return history
    End Function

    Protected Function TestEnableHistoryCommandReturnsError(cmdExec As CommandExecutive, vars() As String, expectedMsg As String) As Object
        Dim history As RTAUrtTraceHistory = New RTAUrtTraceHistory
        Dim EnableHistoryCommand As New RTAURTCommandEnableHistory
        Dim EnableHistoryCommandParams As New RTAURTCommandEnableHistoryParameters(history)
        For Each v In vars
            EnableHistoryCommandParams.EnableHistoryFor(v)
        Next
        TestCommandReturnsError(EnableHistoryCommandParams, expectedMsg, cmdExec)
        Return history
    End Function

    Protected Sub TestTestMessageIsDone(cmdExec As CommandExecutive, msg As String)
        Dim MessageCommandParams = New RTAURTCommandMessageParameters(msg)
        TestCommandReturnsDone(MessageCommandParams, cmdExec)
        Assert.AreEqual(1, cmdExec.vbfb.Logger.Count)
        Assert.AreEqual(msg, cmdExec.vbfb.Logger.GetLastMessage)
    End Sub
End Class

<TestClass()> Public Class UnitTestCommands
    Inherits UnitTestCommandsBase

    <TestMethod()> Public Sub TestCommandConnect()
        StartWithCommandConnect()
    End Sub

    <TestMethod()> Public Sub TestCommandConnectAnd1Execute()
        Dim cmdExec As CommandExecutive = StartWithCommandConnect()
        Dim ExecuteCommandParams As New RTAURTCommandExecuteVBParameters()

        TestCommandReturnsDone(ExecuteCommandParams, cmdExec)

        TestEndOfExecuteVBRun(cmdExec, 1)
    End Sub

    <TestMethod()> Public Sub TestCommandConnectAnd5Executes()
        Dim cmdExec As CommandExecutive = StartWithCommandConnect()

        Dim NumberOfExecutes As Integer = 5
        Dim ExecuteCommandParams As New RTAURTCommandExecuteVBParameters(NumberOfExecutes)

        Dim outCounter As ConInt = cmdExec.vbfb.GetElement(ScriptDataElementNames.outCounter)
        Dim InitCounterVal As Integer = 3
        outCounter.Val = InitCounterVal

        TestCommandReturnsDone(ExecuteCommandParams, cmdExec)

        TestEndOfExecuteVBRun(cmdExec, NumberOfExecutes)
        Assert.AreEqual(NumberOfExecutes + InitCounterVal, outCounter.Val)
    End Sub

    <TestMethod()> Public Sub TestCommandSetValuesFor1IntValue()
        Dim cmdExec As CommandExecutive = StartWithCommandConnect()

        Dim NumberOfExecutes As Integer = 3
        Dim ExecuteVBCommandParams As New RTAURTCommandExecuteVBParameters(NumberOfExecutes)

        Dim outCounter As ConInt = cmdExec.vbfb.GetElement(ScriptDataElementNames.outCounter)
        Dim InitCounterVal As Integer = 7

        Dim SetValuesCommandParams As New RTAURTCommandSetValuesParameters()
        SetValuesCommandParams.AddValue("outCounter", InitCounterVal)

        TestCommandReturnsDone(SetValuesCommandParams, cmdExec)
        TestCommandReturnsDone(ExecuteVBCommandParams, cmdExec)

        TestEndOfExecuteVBRun(cmdExec, NumberOfExecutes)
        Assert.AreEqual(InitCounterVal + NumberOfExecutes, outCounter.Val)
    End Sub

    <TestMethod()> Public Sub TestCommandSetValuesFor2ValuesIntAndBool()
        Dim cmdExec As CommandExecutive = StartWithCommandConnect()

        Dim NumberOfExecutes As Integer = 3
        Dim ExecuteVBCommandParams As New RTAURTCommandExecuteVBParameters(NumberOfExecutes)

        Dim outCounter As ConInt = cmdExec.vbfb.GetElement(ScriptDataElementNames.outCounter)
        Dim InitCounterVal As Integer = 7

        Dim SetValuesCommandParams As New RTAURTCommandSetValuesParameters()
        SetValuesCommandParams.AddValue(ScriptDataElementNames.outCounter, InitCounterVal)
        SetValuesCommandParams.AddValue(ScriptDataElementNames.inUseUpCounter, False)

        TestCommandReturnsDone(SetValuesCommandParams, cmdExec)
        TestCommandReturnsDone(ExecuteVBCommandParams, cmdExec)
        TestEndOfExecuteVBRun(cmdExec, NumberOfExecutes)

        Assert.AreEqual(InitCounterVal - NumberOfExecutes, outCounter.Val)
    End Sub

    <TestMethod()> Public Sub TestCommandSetValuesWithNoConnectForNothingNameRaisesMessage()
        'Dim cmdExec As CommandExecutive = StartWithCommandConnect()

        Dim newElementValue As Single = 10.3!

        Dim SetValuesCommand As New RTAURTCommandSetValues()
        Dim SetValuesCommandParams As New RTAURTCommandSetValuesParameters()
        SetValuesCommandParams.AddValue(Nothing, newElementValue)

        TestCommandReturnsError(SetValuesCommandParams, SetValuesCommand.ErrorMessages.CantFindElement)

    End Sub

    <TestMethod()> Public Sub TestCommandSetValuesForNothingNameRaisesMessageAfterConnect()
        Dim cmdExec As CommandExecutive = StartWithCommandConnect()

        Dim newElementValue As Single = 10.3!
        Dim SetValuesCommand As New RTAURTCommandSetValues()
        Dim SetValuesCommandParams As New RTAURTCommandSetValuesParameters()
        SetValuesCommandParams.AddValue(Nothing, newElementValue)

        TestCommandReturnsError(SetValuesCommandParams, RTAURTCommandSetValues.ErrorMessages.CantFindElement)

    End Sub

    <TestMethod()> Public Sub TestCommandSetValuesForUnknownNameRaisesMessageAfterConnect()
        Dim cmdExec As CommandExecutive = StartWithCommandConnect()

        Dim newElementValue As Single = 10.3!
        Dim SetValuesCommand As New RTAURTCommandSetValues()
        Dim SetValuesCommandParams As New RTAURTCommandSetValuesParameters()
        SetValuesCommandParams.AddValue(Nothing, newElementValue)
        SetValuesCommandParams.AddValue("Does not exist", newElementValue)

        TestCommandReturnsError(SetValuesCommandParams, RTAURTCommandSetValues.ErrorMessages.CantFindElement)

    End Sub

    <TestMethod()> Public Sub TestCommandSetValuesSingleWrongValueTypeGivesError()
        Dim cmdExec As CommandExecutive = StartWithCommandConnect()

        Dim newElementValue As String = "this cannot be a single"
        Dim SetValuesCommand As New RTAURTCommandSetValues()
        Dim SetValuesCommandParams As New RTAURTCommandSetValuesParameters()
        SetValuesCommandParams.AddValue(ScriptDataElementNames.inFloat1, newElementValue)

        TestCommandReturnsError(SetValuesCommandParams, RTAURTCommandSetValues.ErrorMessages.ValueFailsToConvert, cmdExec)
    End Sub

    <TestMethod()> Public Sub TestCommandSetValuesSingleArrayWrongValueTypeGivesError()
        Dim cmdExec As CommandExecutive = StartWithCommandConnect()

        Dim newElementValue As String = "this cannot be a single"
        Dim SetValuesCommand As New RTAURTCommandSetValues()
        Dim SetValuesCommandParams As New RTAURTCommandSetValuesParameters()
        SetValuesCommandParams.AddValue(ScriptDataElementNames.inArrayFloat, newElementValue, 0)

        TestCommandReturnsError(SetValuesCommandParams, RTAURTCommandSetValues.ErrorMessages.ValueFailsToConvert, cmdExec)
    End Sub

    <TestMethod()> Public Sub TestCommandSetValuesFor1IntArrayValue()
        Dim cmdExec As CommandExecutive = StartWithCommandConnect()

        Dim OutFloat2 As ConFloat = cmdExec.vbfb.GetElement(ScriptDataElementNames.outFloat2)

        Dim newElementValue As Single = 10.3!
        Dim SetValuesCommandParams As New RTAURTCommandSetValuesParameters()
        SetValuesCommandParams.AddValue(ScriptDataElementNames.inArrayFloat, newElementValue, 0)

        Dim ExecuteVBCommandParams As New RTAURTCommandExecuteVBParameters()

        TestCommandReturnsDone(SetValuesCommandParams, cmdExec)
        TestCommandReturnsDone(ExecuteVBCommandParams, cmdExec)

        Assert.AreEqual(newElementValue + 1.0! * 3.0!, OutFloat2.Val)

    End Sub

    <TestMethod()>
    <ExpectedException(GetType(ArgumentOutOfRangeException))>
    Public Sub TestCommandEnableHistoryNullNameThrowsException()
        Dim cmdExec As CommandExecutive = StartWithCommandConnect()

        Dim history As RTAUrtTraceHistory = TestEnableHistoryCommandReturnsDone(cmdExec, {Nothing})
    End Sub

    <TestMethod()>
    Public Sub TestCommandEnableHistoryInvalidParameterReturnsError()
        Dim cmdExec As CommandExecutive = StartWithCommandConnect()

        Dim history As RTAUrtTraceHistory =
            TestEnableHistoryCommandReturnsError(cmdExec,
                                                 {"Not a valid VB Element Name"},
                                                 RTAURTCommandEnableHistory.ErrorMessages.CantFindElement)
    End Sub

    <TestMethod()>
    Public Sub TestCommandEnableHistoryFor1ScalarSucceeds()
        Dim cmdExec As CommandExecutive = StartWithCommandConnect()

        Dim history As RTAUrtTraceHistory = TestEnableHistoryCommandReturnsDone(cmdExec, {ScriptDataElementNames.outCounter})
        TestExecuteVBCommandReturnsDone(cmdExec, 1)

        Assert.AreEqual(1, history.CountTimeStamps())
        Assert.AreEqual(1, history.CountTags())
    End Sub

    <TestMethod()>
    Public Sub TestCommandEnableHistoryFor1BoolArray()
        Dim cmdExec As CommandExecutive = StartWithCommandConnect()

        Dim history As RTAUrtTraceHistory = TestEnableHistoryCommandReturnsDone(cmdExec, {ScriptDataElementNames.inArrayBool})

        TestExecuteVBCommandReturnsDone(cmdExec, 1)

        Assert.AreEqual(1, history.CountTimeStamps())
        Assert.AreEqual(1, history.CountTags())
    End Sub

    <TestMethod()>
    Public Sub TestCommandEnableHistoryFor3Items2Executes()
        Dim cmdExec As CommandExecutive = StartWithCommandConnect()
        Dim history As RTAUrtTraceHistory = TestEnableHistoryCommandReturnsDone(cmdExec,
                                                                        {ScriptDataElementNames.inArrayBool,
                                                                        ScriptDataElementNames.outCounter,
                                                                        ScriptDataElementNames.inArrayFloat})
        TestExecuteVBCommandReturnsDone(cmdExec, 2)

        Dim _INARRAYBOOL_SIZE As Integer = 4
        Dim _INARRAYFLOAT_SIZE As Integer = 4
        Assert.AreEqual(2, history.CountTimeStamps())
        Assert.AreEqual(3, history.CountTags())
        Assert.AreEqual(_INARRAYBOOL_SIZE + _INARRAYFLOAT_SIZE + 1 - 1, history.GetLog.Last.Count(Function(c As Char) c = ","c))
    End Sub

    <TestMethod()> Public Sub TestCommandMessageCommandWrite1Message()
        Dim cmdExec As CommandExecutive = StartWithCommandConnect()
        TestTestMessageIsDone(cmdExec, "This is a Test Message from TestCommandMessageCommandWrite1Message")
    End Sub

    <TestMethod()> Public Sub TestCommandClearLogsClearsOnlyMessages()
        Dim cmdExec As CommandExecutive = StartWithCommandConnect()

        TestEnableHistoryCommandReturnsDone(cmdExec, {ScriptDataElementNames.inArrayFloat})
        TestTestMessageIsDone(cmdExec, "This is a Test Message from TestCommandClearLogsClearsOnlyMessages")

        TestExecuteVBCommandReturnsDone(cmdExec, 2)
        Assert.AreEqual(1, cmdExec.vbfb.Logger.Count)
        Assert.AreEqual(2, cmdExec.vbfb.HistoryLog.CountTimeStamps())

        Dim ClearLogCommand As New RTAURTCommandClearLogs
        Dim ClearMessageLogCommandParams As New RTAURTCommandClearLogsParameters(RTAURTCommandClearLogs.Logs.MessageLog)
        TestCommandReturnsDone(ClearMessageLogCommandParams, cmdExec)
        Assert.AreEqual(0, cmdExec.vbfb.Logger.Count)
        Assert.AreEqual(2, cmdExec.vbfb.HistoryLog.CountTimeStamps())
    End Sub

    <TestMethod()> Public Sub TestCommandClearLogsClearsOnlyHistory()
        Dim cmdExec As CommandExecutive = StartWithCommandConnect()
        TestEnableHistoryCommandReturnsDone(cmdExec, {ScriptDataElementNames.inArrayFloat})

        TestTestMessageIsDone(cmdExec, "This is a Test Message from TestCommandClearLogsClearsOnlyHistory")

        TestExecuteVBCommandReturnsDone(cmdExec, 2)
        Assert.AreEqual(1, cmdExec.vbfb.Logger.Count)
        Assert.AreEqual(2, cmdExec.vbfb.HistoryLog.CountTimeStamps())

        Dim ClearLogCommand As New RTAURTCommandClearLogs
        Dim ClearMessageLogCommandParams As New RTAURTCommandClearLogsParameters(RTAURTCommandClearLogs.Logs.HistoryLog)

        TestCommandReturnsDone(ClearMessageLogCommandParams, cmdExec)
        Assert.AreEqual(1, cmdExec.vbfb.Logger.Count)
        Assert.AreEqual(0, cmdExec.vbfb.HistoryLog.CountTimeStamps())

    End Sub

    <TestMethod()> Public Sub TestCommandClearLogsClearsMessageAndHistory()
        Dim cmdExec As CommandExecutive = StartWithCommandConnect()
        TestEnableHistoryCommandReturnsDone(cmdExec, {ScriptDataElementNames.inArrayFloat})

        TestTestMessageIsDone(cmdExec, "This is a Test Message from TestCommandClearLogsClearsMessageAndHistory")

        TestExecuteVBCommandReturnsDone(cmdExec, 2)
        Assert.AreEqual(1, cmdExec.vbfb.Logger.Count)
        Assert.AreEqual(2, cmdExec.vbfb.HistoryLog.CountTimeStamps())

        Dim ClearLogCommand As New RTAURTCommandClearLogs
        Dim ClearMessageLogCommandParams As New RTAURTCommandClearLogsParameters(RTAURTCommandClearLogs.Logs.MessageAndHistory)

        TestCommandReturnsDone(ClearMessageLogCommandParams, cmdExec)
        Assert.AreEqual(0, cmdExec.vbfb.Logger.Count)
        Assert.AreEqual(0, cmdExec.vbfb.HistoryLog.CountTimeStamps())
    End Sub

End Class