Option Explicit On
Option Strict On

Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports URT
Imports RTAURTService


<TestClass()> Public Class UnitTestSequences
    Inherits UnitTestCommandsBase

    Private _logger As RTAUrtTraceLogger
    Dim _vbfb As URTVBQALSimpleScript.URTVBQALSimpleScript
    Dim _seqParams As RTAURTCommandSequenceParameters
    Dim _cmdExec As CommandExecutive

    Private Sub InitCommandExecutive()
        _logger = New RTAUrtTraceLogger
        _vbfb = New URTVBQALSimpleScript.URTVBQALSimpleScript(_logger)
        _cmdExec = New CommandExecutive(_vbfb)
    End Sub

    Private Sub InitSequenceWithLogger()
        InitCommandExecutive()
        _seqParams = SequenceFactory.MakeSequence
        Assert.IsNotNull(_seqParams)
    End Sub

    Private Function MakeLoadedTestSequence1() As RTAURTCommandSequenceParameters
        InitSequenceWithLogger()

        Dim testMsg As String = "This is a test"
        Dim MessageCommandParams = New RTAURTCommandMessageParameters(testMsg)

        Dim testMsg2 As String = "This is a 2nd test"
        Dim MessageCommandParams2 = New RTAURTCommandMessageParameters(testMsg2)
        _seqParams.AddCommandParameters(MessageCommandParams)
        _seqParams.AddCommandParameters(MessageCommandParams2)

        Return _seqParams
    End Function

    <TestMethod()> Public Sub TestSequenceFactoryMakesSequenceParams()
        InitSequenceWithLogger()
    End Sub

    <TestMethod()> Public Sub TestSequenceFactoryMakesUniqueSequenceParams()
        Dim seqParams1 As RTAURTCommandSequenceParameters = SequenceFactory.MakeSequence
        Dim seqParams2 As RTAURTCommandSequenceParameters = SequenceFactory.MakeSequence

        Assert.AreNotEqual(seqParams1.Name, seqParams2.Name)
    End Sub

    <TestMethod()> Public Sub TestSequenceExecuteWithNoCommandsMakes1Message()
        InitSequenceWithLogger()

        Dim seq As RTAURTCommandSequence = New RTAURTCommandSequence

        seq.Execute(_vbfb, CType(_seqParams, RTAURTCommandParameters))
        Assert.AreEqual(1, _logger.Count)
        Assert.AreEqual(RTAURTCommandErrorMessages.SequenceErrorNoCommands, _logger.GetLastMessage)
    End Sub

    <TestMethod()> Public Sub TestSequenceExecuteWith1MessageCommandsMakes1Message()
        InitSequenceWithLogger()

        Dim seq As RTAURTCommandSequence = New RTAURTCommandSequence

        Dim testMsg As String = "This is a test"
        Dim MessageCommandParams = New RTAURTCommandMessageParameters(testMsg)
        _seqParams.AddCommandParameters(MessageCommandParams)
        seq.Execute(_vbfb, CType(_seqParams, RTAURTCommandParameters))
        Assert.AreEqual(1, _logger.Count)
        Assert.AreEqual(testMsg, _logger.GetLastMessage)
    End Sub

    <TestMethod()> Public Sub TestLinkedSequence()
        Dim seqParams As RTAURTCommandSequenceParameters = MakeLoadedTestSequence1()

        Dim LinkToSeqParams As RTAURTCommandLinkedSequenceParameters = New RTAURTCommandLinkedSequenceParameters(seqParams.Name)

        TestCommandReturnsDone(LinkToSeqParams, _cmdExec)

        Assert.AreEqual(4, _cmdExec.TotalCommandsExecuted)
    End Sub



    <TestMethod()> Public Sub TestLinkedSequenceRecursiveCall()
        Dim seqParams As RTAURTCommandSequenceParameters = MakeLoadedTestSequence1()

        Dim LinkToSeqParams As RTAURTCommandLinkedSequenceParameters = New RTAURTCommandLinkedSequenceParameters(seqParams.Name)
        seqParams.AddCommandParameters(LinkToSeqParams)

        Dim recDepth As Integer = 4

        _cmdExec.RecursionDepth = recDepth
        TestCommandReturnsStopped(LinkToSeqParams, _cmdExec)
        'Dim callback As ICommandCallback = CommandCallbacks.LimitRecursionTo(20)
        'Dim callback As ICommandCallback = CommandCallbacks.LimitCommandCallsTo(20)
        'cmdResult = seqParams.GetCommand.Execute(_urtVBFB, CType(_seqParams, RTAURTCommandParameters), callback)

        Dim numberOfCommandsInSeq As Integer = 3

        Assert.AreEqual(CInt((numberOfCommandsInSeq + 1) * recDepth + 1), _cmdExec.TotalCommandsExecuted)


    End Sub

End Class