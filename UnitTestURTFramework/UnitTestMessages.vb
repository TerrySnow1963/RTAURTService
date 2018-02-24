Option Explicit On
Option Strict On

Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting

<TestClass()> Public Class UnitTestMessages

    <TestMethod()> Public Sub TestZeroMessagesReturnZero()
        Dim log As RTAURTService.RTAUrtTraceLogger = New RTAURTService.RTAUrtTraceLogger
        Assert.AreEqual(0, log.Count)
    End Sub

    <TestMethod()> Public Sub TestAddNullOrEmptyMessagesReturnZero()
        Dim log As RTAURTService.RTAUrtTraceLogger = New RTAURTService.RTAUrtTraceLogger

        log.Write(Nothing)
        log.Write(String.Empty)
        Assert.AreEqual(0, log.Count)
    End Sub

    <TestMethod()> Public Sub TestAdd2MessagesReturn2AndGetLastMessage()
        Dim log As RTAURTService.RTAUrtTraceLogger = New RTAURTService.RTAUrtTraceLogger
        Dim msgs() As String = {"message 1", "another message"}

        For Each msg In msgs
            log.Write(msg)
        Next

        Assert.AreEqual(2, log.Count)
        Assert.AreEqual(msgs.Last, log.GetLastMessage)
    End Sub

    <TestMethod()> Public Sub TestZeroMessagesGetLastMessageIsEmpty()
        Dim log As RTAURTService.RTAUrtTraceLogger = New RTAURTService.RTAUrtTraceLogger

        Assert.AreEqual(String.Empty, log.GetLastMessage)
    End Sub
End Class