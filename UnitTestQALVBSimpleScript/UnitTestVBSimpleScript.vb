Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports UnitTestQALVBSimpleScript
Imports URT
Imports RTAURTService
Imports UrtTlbLib

<TestClass()> Public Class UnitTestVBSimpleScript

    <TestMethod()> Public Sub TestAllElementsConnected()
        Dim urtVBFB = New URTVBQALSimpleScript.URTVBQALSimpleScript


        urtVBFB.Connect(True)

        TestAllElementsConnected(urtVBFB)

    End Sub

    Private Sub TestAllElementsConnected(urtVBFB As URTVBQALSimpleScript.URTVBQALSimpleScript)
        Dim elementNames() As String = {"inFloat1", "outFloat1", "inArrayBool", "inSize"}

        Dim element As IUrtData

        For Each name In elementNames
            Trace.Write(String.Format("Testing {0} is not Null - ", name))
            element = urtVBFB.GetElement(name)
            Assert.IsNotNull(element)
            Assert.AreEqual(name, element.Name)
            Trace.WriteLine("OK")
        Next
    End Sub

    <TestMethod()> Public Sub TestExecuteUsingInit()
        Dim urtVBFB = New URTVBQALSimpleScript.URTVBQALSimpleScript
        urtVBFB.Connect(True)
        TestAllElementsConnected(urtVBFB)

        Dim inFloat1 As ConFloat = urtVBFB.GetElement("inFloat1")
        Dim outFloat1 As ConFloat = urtVBFB.GetElement("outFloat1")

        inFloat1.Val = 5.0
        urtVBFB.Execute(0, Nothing)

        Assert.AreEqual(10.0!, outFloat1.Val)

    End Sub

    <TestMethod()> Public Sub TestConArrayBoolSize()
        Dim urtVBFB = New URTVBQALSimpleScript.URTVBQALSimpleScript
        urtVBFB.Connect(True)
        TestAllElementsConnected(urtVBFB)

        Dim inArrayBool As IUrtData = urtVBFB.GetElement("inArrayBool")

        Assert.AreEqual(4, inArrayBool.Size)

    End Sub

    <TestMethod()> Public Sub TestExecuteConArrayBoolChangeSize()
        Dim urtVBFB = New URTVBQALSimpleScript.URTVBQALSimpleScript
        urtVBFB.Connect(True)
        TestAllElementsConnected(urtVBFB)

        Dim inSize As ConInt = urtVBFB.GetElement("inSize")

        Dim inArrayBool As IUrtData = urtVBFB.GetElement("inArrayBool")

        inSize.Val = 5

        urtVBFB.Execute(0, Nothing)

        Assert.AreEqual(5, inArrayBool.Size)

    End Sub

    <TestMethod()> Public Sub TestExecuteChangingArrayBool()
        Dim urtVBFB = New URTVBQALSimpleScript.URTVBQALSimpleScript
        urtVBFB.Connect(True)
        TestAllElementsConnected(urtVBFB)

        Dim inFloat1 As ConFloat = urtVBFB.GetElement("inFloat1")
        Dim outFloat1 As ConFloat = urtVBFB.GetElement("outFloat1")
        Dim inArrayBool As ConArrayBool = CType(urtVBFB.GetElement("inArrayBool"), ConArrayBool)

        inFloat1.Val = 5.0!
        inArrayBool(0) = True
        inArrayBool(1) = True


        Trace.WriteLine("Testing inArray(0) is True, and inArray(1) is True")
        urtVBFB.Execute(0, Nothing)

        Dim expectedResult As Single = 3.0! * 5.0!

        Assert.AreEqual(expectedResult, outFloat1.Val)

        inArrayBool(0) = False
        expectedResult = 2.0! * 5.0!

        Trace.WriteLine("Testing inArray(0) is False, and inArray(1) is True")
        urtVBFB.Execute(0, Nothing)
        Assert.AreEqual(expectedResult, outFloat1.Val)

    End Sub

End Class