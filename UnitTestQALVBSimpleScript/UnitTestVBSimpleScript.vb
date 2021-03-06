﻿Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports UnitTestQALVBSimpleScript
Imports URT
Imports RTAURTService
Imports UrtTlbLib
Imports RTAInterfaces
Imports URTVBQALSimpleScript

<TestClass()> Public Class UnitTestVBSimpleScript

    <TestMethod()> Public Sub TestAllElementsConnected()
        Dim urtVBFB = New URTVBQALSimpleScript.URTVBQALSimpleScript


        urtVBFB.Connect(True)

        TestAllElementsConnected(urtVBFB)

    End Sub

    Private Sub TestAllElementsConnected(urtVBFB As URTVBQALSimpleScript.URTVBQALSimpleScript)
        Dim elementNames() As String = {"inFloat1", "outFloat1", "inArrayBool", "inSize", "inArrayFloat"}

        Dim element As IRTAUrtData

        For Each name In elementNames
            Trace.Write(String.Format("Testing {0} is not Null - ", name))
            element = CType(urtVBFB.GetElement(name), IRTAUrtData)
            Assert.IsNotNull(element)
            Assert.AreEqual(name, element.Name)
            Trace.WriteLine("OK")
        Next
    End Sub

    <TestMethod()> Public Sub TestExecuteUsingInit()
        Dim logger As RTAUrtTraceLogger = New RTAUrtTraceLogger
        Dim urtVBFB = New URTVBQALSimpleScript.URTVBQALSimpleScript(logger)

        urtVBFB.Connect(True)
        TestAllElementsConnected(urtVBFB)

        Dim inFloat1 As ConFloat = urtVBFB.GetElement("inFloat1")
        Dim outFloat1 As ConFloat = urtVBFB.GetElement("outFloat1")

        Dim inArrayBool As ConArrayBool = urtVBFB.GetElement("inArrayBool")

        inFloat1.Val = 5.0
        urtVBFB.Execute(0, Nothing)

        Trace.WriteLine("Testing InArray(0) AND InArray(1) (both True) gives 5.0 + 5.0")
        Assert.AreEqual(10.0!, outFloat1.Val)

        Trace.WriteLine("Testing InArray(0) AND InArray(1) (one False True) gives 3*5.0 = 15.0")
        inArrayBool.Item(1) = False
        urtVBFB.Execute(0, Nothing)
        Assert.AreEqual(15.0!, outFloat1.Val)
        Assert.AreEqual(0, logger.Count)

    End Sub

    <TestMethod()> Public Sub TestConArrayBoolSize()
        Dim urtVBFB = New URTVBQALSimpleScript.URTVBQALSimpleScript
        urtVBFB.Connect(True)
        TestAllElementsConnected(urtVBFB)

        Dim inArrayBool As UrtTlbLib.IUrtData = urtVBFB.GetElement("inArrayBool")

        Assert.AreEqual(4, inArrayBool.Size(urtBUF.dbWork))

    End Sub

    <TestMethod()> Public Sub TestExecuteConArrayBoolChangeSize()
        Dim urtVBFB = New URTVBQALSimpleScript.URTVBQALSimpleScript
        urtVBFB.Connect(True)
        TestAllElementsConnected(urtVBFB)

        Dim inSize As ConInt = urtVBFB.GetElement("inSize")

        Dim inArrayBool As UrtTlbLib.IUrtData = urtVBFB.GetElement("inArrayBool")

        inSize.Val = 5

        urtVBFB.Execute(0, Nothing)

        Assert.AreEqual(5, inArrayBool.Size(urtBUF.dbWork))

    End Sub

    <TestMethod()> Public Sub TestExecuteChangingArrayBool()
        Dim urtVBFB = New URTVBQALSimpleScript.URTVBQALSimpleScript
        urtVBFB.Connect(True)
        TestAllElementsConnected(urtVBFB)

        Dim inFloat1 As ConFloat = urtVBFB.GetElement("inFloat1")
        Dim outFloat1 As ConFloat = urtVBFB.GetElement("outFloat1")
        Dim inArrayBool As ConArrayBool = CType(urtVBFB.GetElement("inArrayBool"), ConArrayBool)

        inFloat1.Val = 5.0!
        inArrayBool.Item(0) = True
        inArrayBool.Item(1) = True


        Trace.WriteLine("Testing inArray(0) is True, and inArray(1) is True")
        urtVBFB.Execute(0, Nothing)
        Dim expectedResult As Single

        expectedResult = 2.0! * 5.0!
        Assert.AreEqual(expectedResult, outFloat1.Val)

        inArrayBool.Item(0) = False

        expectedResult = 3.0! * 5.0!

        Trace.WriteLine("Testing inArray(0) is False, and inArray(1) is True")
        urtVBFB.Execute(0, Nothing)
        Assert.AreEqual(expectedResult, outFloat1.Val)

    End Sub

End Class