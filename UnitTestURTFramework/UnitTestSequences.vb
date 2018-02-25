Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports URT
Imports RTAURTService

<TestClass()> Public Class UnitTestSequences
    Private _logger As RTAUrtTraceLogger
    Dim _urtVBFB As URTVBQALSimpleScript.URTVBQALSimpleScript
    Dim _seqParams As RTAURTActionSequenceParameters

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

        Dim seq As RTAURTActionSequence = New RTAURTActionSequence

        seq.Execute(_urtVBFB, _seqParams)
        Assert.AreEqual(1, _logger.Count)
        Assert.AreEqual(RTAURTActionErrorMessages.SequenceErrorNoCommands, _logger.GetLastMessage)
    End Sub

    <TestMethod()> Public Sub TestSequenceExecuteWithOneMessageCommandsMakes1Message()
        InitSequenceWithLogger()

        Dim seq As RTAURTActionSequence = New RTAURTActionSequence

        Dim testMsg As String = "This is a test"
        Dim MessageActionParams = New RTAURTActionMessageParameters(testMsg)
        _seqParams.AddCommandParameters(MessageActionParams)
        seq.Execute(_urtVBFB, _seqParams)
        Assert.AreEqual(1, _logger.Count)
        Assert.AreEqual(testMsg, _logger.GetLastMessage)
    End Sub

End Class