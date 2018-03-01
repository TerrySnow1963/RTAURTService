Option Explicit On
Option Strict On

Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports URT
Imports RTAURTService


<TestClass()> Public Class UnitTestSequences
    Private _logger As RTAUrtTraceLogger
    Dim _urtVBFB As URTVBQALSimpleScript.URTVBQALSimpleScript
    Dim _seqParams As RTAURTCommandSequenceParameters

    Private Sub InitSequenceWithLogger()
        _logger = New RTAUrtTraceLogger
        _urtVBFB = New URTVBQALSimpleScript.URTVBQALSimpleScript(_logger)
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

    <TestMethod()> Public Sub TestSequenceExecuteWithNoCommandsMakes1Message()
        InitSequenceWithLogger()

        Dim seq As RTAURTCommandSequence = New RTAURTCommandSequence

        seq.Execute(_urtVBFB, CType(_seqParams, RTAURTCommandParameters))
        Assert.AreEqual(1, _logger.Count)
        Assert.AreEqual(RTAURTCommandErrorMessages.SequenceErrorNoCommands, _logger.GetLastMessage)
    End Sub

    <TestMethod()> Public Sub TestSequenceExecuteWith1MessageCommandsMakes1Message()
        InitSequenceWithLogger()

        Dim seq As RTAURTCommandSequence = New RTAURTCommandSequence

        Dim testMsg As String = "This is a test"
        Dim MessageCommandParams = New RTAURTCommandMessageParameters(testMsg)
        _seqParams.AddCommandParameters(MessageCommandParams)
        seq.Execute(_urtVBFB, CType(_seqParams, RTAURTCommandParameters))
        Assert.AreEqual(1, _logger.Count)
        Assert.AreEqual(testMsg, _logger.GetLastMessage)
    End Sub

    <TestMethod()> Public Sub TestLinkedSequence()
        Dim seqParams As RTAURTCommandSequenceParameters = MakeLoadedTestSequence1()

        Dim LinkToSeqParams As RTAURTCommandLinkedSequenceParameters = New RTAURTCommandLinkedSequenceParameters(seqParams.Name)

        Dim cmdResult As ICommandResult
        cmdResult = seqParams.GetCommand.Execute(_urtVBFB, CType(_seqParams, RTAURTCommandParameters))
        Assert.AreEqual(2, _logger.Count)
    End Sub



    <TestMethod()> Public Sub TestLinkedSequenceRecursiveCall()
        Dim seqParams As RTAURTCommandSequenceParameters = MakeLoadedTestSequence1()

        Dim LinkToSeqParams As RTAURTCommandLinkedSequenceParameters = New RTAURTCommandLinkedSequenceParameters(_seqParams.Name)
        _seqParams.AddCommandParameters(LinkToSeqParams)

        Dim cmdResult As ICommandResult
        'Dim callback As ICommandCallback = CommandCallbacks.LimitRecursionTo(20)
        Dim callback As ICommandCallback = CommandCallbacks.LimitCommandCallsTo(20)
        cmdResult = _seqParams.GetCommand.Execute(_urtVBFB, CType(_seqParams, RTAURTCommandParameters), callback)
        Assert.AreEqual(CInt(20 / 2), _logger.Count)
    End Sub

End Class