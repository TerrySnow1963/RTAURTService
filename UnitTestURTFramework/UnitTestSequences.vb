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

    <TestMethod()> Public Sub TestSequenceBreaksAtSomeRecursionLevel()
        InitSequenceWithLogger()

        Dim seq As RTAURTCommandSequence = New RTAURTCommandSequence

        Dim testMsg As String = "This is a test"
        Dim MessageCommandParams = New RTAURTCommandMessageParameters(testMsg)

        Dim testMsg2 As String = "This is a 2nd test"
        Dim MessageCommandParams2 = New RTAURTCommandMessageParameters(testMsg2)
        _seqParams.AddCommandParameters(MessageCommandParams)
        _seqParams.AddCommandParameters(MessageCommandParams2)
        seq.Execute(_urtVBFB, CType(_seqParams, RTAURTCommandParameters))
        Assert.AreEqual(2, _logger.Count)
        Assert.AreEqual(testMsg2, _logger.GetLastMessage)
    End Sub

    <TestMethod()> Public Sub TestSequenceExecuteWith2MessageCommandsMakes2Message()
        InitSequenceWithLogger()

        Dim seq As RTAURTCommandSequence = New RTAURTCommandSequence

        Assert.AreEqual("Need To", " Test for recursive calls.")
    End Sub

End Class