Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports UrtTlbLib
Imports URTVBFunctionBlock
Imports RTAInterfaces
Imports RTAURTService

<TestClass()> Public Class UnitTestQALFillTankMarried1

    <TestMethod()> Public Sub TestConnectConBool()
        Dim messageLog As RTAUrtTraceLogger = New RTAUrtTraceLogger

        Dim urtVBFB = New URTVBQALFillTankMarried1(messageLog)

        urtVBFB.Connect(True)

        Dim ClearError As ConBool = urtVBFB.GetElement("ClearError")

        Assert.IsNotNull(ClearError)
        Assert.AreEqual("ClearError", CType(ClearError, IRTAUrtData).Name)
        Assert.AreEqual("Clear Error Messages.", CType(ClearError, IRTAUrtData).Description)
        Assert.IsFalse(ClearError.Val)

        Assert.AreEqual(0, messageLog.Count())

    End Sub

    <TestMethod()> Public Sub TestConnectConString()

        Dim urtVBFB = New URTVBFunctionBlock.URTVBQALFillTankMarried1

        urtVBFB.Connect(True)

        Dim FBErrors As ConString = urtVBFB.GetElement("FBErrors")

        Assert.IsNotNull(FBErrors)
        Assert.AreEqual("FBErrors", CType(FBErrors, IRTAUrtData).Name)
        Assert.AreEqual("Error on Last Run.", CType(FBErrors, IRTAUrtData).Description)
        Assert.AreEqual(String.Empty, FBErrors.Val)

    End Sub

    <TestMethod()> Public Sub TestConnectConInt()

        Dim urtVBFB = New URTVBFunctionBlock.URTVBQALFillTankMarried1

        urtVBFB.Connect(True)

        Dim FBExecTime As ConInt = urtVBFB.GetElement("FBExecTime")

        Assert.IsNotNull(FBExecTime)
        Assert.AreEqual("FBExecTime", CType(FBExecTime, IRTAUrtData).Name)
        Assert.AreEqual("Run Time of Last FB Execution in Milliseconds.", CType(FBExecTime, IRTAUrtData).Description)
        Assert.AreEqual(0, FBExecTime.Val)

    End Sub

    <TestMethod()> Public Sub TestConnectConEnum()

        Dim urtVBFB = New URTVBFunctionBlock.URTVBQALFillTankMarried1

        urtVBFB.Connect(True)

        Dim inAuto As ConEnum = urtVBFB.GetElement("inAuto")

        Assert.IsNotNull(inAuto)
        Assert.AreEqual("inAuto", CType(inAuto, IRTAUrtData).Name)
        Assert.AreEqual("inAuto - Desc", CType(inAuto, IRTAUrtData).Description)

        Dim expectedVal As URTVBFunctionBlock.URT.urtNOYES = URTVBFunctionBlock.URT.urtNOYES.NO
        Assert.AreEqual(CInt(expectedVal), inAuto.Val)

    End Sub

    <TestMethod()> Public Sub TestConnectConFloat()

        Dim urtVBFB = New URTVBFunctionBlock.URTVBQALFillTankMarried1

        urtVBFB.Connect(True)

        Dim inFillEast As ConFloat = urtVBFB.GetElement("inFillEast")

        Assert.IsNotNull(inFillEast)
        Assert.AreEqual("inFillEast", CType(inFillEast, IRTAUrtData).Name)
        Assert.AreEqual("inFillEast - Desc", CType(inFillEast, IRTAUrtData).Description)

    End Sub

    <TestMethod()> Public Sub TestExecuteFBExecTimeEQZero()
        Dim urtVBFB = New URTVBFunctionBlock.URTVBQALFillTankMarried1

        urtVBFB.Connect(True)

        Dim FBExecTime As ConInt = urtVBFB.GetElement("FBExecTime")

        Assert.IsNotNull(FBExecTime)
        Assert.AreEqual(0, FBExecTime.Val)

    End Sub

    <TestMethod()> Public Sub TestConnectAllElements()

        Dim urtVBFB = New URTVBFunctionBlock.URTVBQALFillTankMarried1

        urtVBFB.Connect(True)

        Dim elementNames() As String = {"ClearError", "FBExecTime",
            "FBErrors", "Note", "outStatus", "inFillEast", "inFillWest", "DivorcedTrip", "MarriedTrip",
            "inAuto", "inTankMarriedMode", "outTankMarriedStatus", "outIsMarried"}

        Dim element As IRTAUrtData

        For Each name In elementNames
            Trace.WriteLine("Checking " & name)
            element = CType(urtVBFB.GetElement(name), IRTAUrtData)
            Assert.IsNotNull(element)
            Assert.AreEqual(name, element.Name)
        Next

    End Sub

End Class