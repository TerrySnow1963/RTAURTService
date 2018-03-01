Option Explicit On
Option Strict On

Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports RTAURTService
Imports URT
Imports UrtTlbLib
Imports URTVBQALSimpleScript


<TestClass()> Public Class UnitTestCommandExecutive
    Private _vbfb As URTVBFunctionBlock
    Private _logger As RTAUrtTraceLogger

    Public Sub New()
        _logger = New RTAUrtTraceLogger
        _vbfb = New URTVBQALSimpleScript.URTVBQALSimpleScript(_logger)
    End Sub

    <TestMethod()> Public Sub TestCEInvokesMessageCmd()
        Dim cmdExec = New CommandExecutive(_vbfb)
        Dim params = New RTAURTCommandMessageParameters("A message")
        cmdExec.Invoke(params)

        Assert.AreEqual(1, cmdExec.TotalCommandsExecuted)

    End Sub
    <TestMethod()> Public Sub TestCEInvokes2MessageCmd()
        Dim cmdExec = New CommandExecutive(_vbfb)
        Dim params = New RTAURTCommandMessageParameters("A message")
        cmdExec.Invoke(params)
        cmdExec.Invoke(params)

        Assert.AreEqual(2, cmdExec.TotalCommandsExecuted)

    End Sub

    <TestMethod()> Public Sub TestCEInvokesConnectCmd()
        Dim cmdExec = New CommandExecutive(_vbfb)
        Dim params = New RTAURTCommandConnectParameters(True)
        cmdExec.Invoke(params)

        Assert.AreEqual(1, cmdExec.TotalCommandsExecuted)
    End Sub

    <TestMethod()> Public Sub TestCEInvokesClearLogsCmd()
        Dim cmdExec = New CommandExecutive(_vbfb)
        Dim params = New RTAURTCommandClearLogsParameters(RTAURTCommandClearLogs.Logs.MessageAndHistory)
        cmdExec.Invoke(params)

        Assert.AreEqual(1, cmdExec.TotalCommandsExecuted)
    End Sub

    <TestMethod()> Public Sub TestCEInvokesExecuteVBCmd()
        Dim cmdExec = New CommandExecutive(_vbfb)
        Dim params = New RTAURTCommandExecuteVBParameters()
        cmdExec.Invoke(params)

        Assert.AreEqual(1, cmdExec.TotalCommandsExecuted)
    End Sub

    <TestMethod()> Public Sub TestCEInvokesSetValuesCmd()
        Dim cmdExec = New CommandExecutive(_vbfb)
        Dim conParams = New RTAURTCommandConnectParameters(True)
        cmdExec.Invoke(conParams)

        Dim svParams = New RTAURTCommandSetValuesParameters()
        svParams.AddValue("outCounter", 10)
        cmdExec.Invoke(svParams)

        Assert.AreEqual(2, cmdExec.TotalCommandsExecuted)
    End Sub
End Class