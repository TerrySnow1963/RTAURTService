Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports UrtTlbLib
Imports URTVBFunctionBlock

<TestClass()> Public Class UnitTestConnectionDataItems

    <TestMethod()> Public Sub TestConnectConBool()

        Dim urtVBFB = New URTVBFunctionBlock.URTVBFunctionBlock

        urtVBFB.Connect(True)

        Dim ClearError As ConBool
        ClearError = urtVBFB.GetElement("ClearError")

        Assert.AreEqual("ClearError", CType(ClearError, IUrtData).Name)
        Assert.AreEqual("Clear Error Messages.", CType(ClearError, IUrtData).Description)
        Assert.IsFalse(ClearError.Val)

    End Sub

    <TestMethod()> Public Sub TestConnectConString()

        Dim urtVBFB = New URTVBFunctionBlock.URTVBFunctionBlock

        urtVBFB.Connect(True)

        Dim FBErrors As ConString
        FBErrors = urtVBFB.GetElement("FBErrors")

        Assert.AreEqual("FBErrors", CType(FBErrors, IUrtData).Name)
        Assert.AreEqual("Error on Last Run.", CType(FBErrors, IUrtData).Description)
        Assert.AreEqual(String.Empty, FBErrors.val)

    End Sub

    <TestMethod()> Public Sub TestConnectConInt()

        Dim urtVBFB = New URTVBFunctionBlock.URTVBFunctionBlock

        urtVBFB.Connect(True)

        Dim FBExecTime As ConInt
        FBExecTime = urtVBFB.GetElement("FBExecTime")

        Assert.AreEqual("FBExecTime", CType(FBExecTime, IUrtData).Name)
        Assert.AreEqual("Run Time of Last FB Execution in Milliseconds.", CType(FBExecTime, IUrtData).Description)
        Assert.AreEqual(0, FBExecTime.Val)

    End Sub

    <TestMethod()> Public Sub TestConnectConEnum()

        Dim urtVBFB = New URTVBFunctionBlock.URTVBFunctionBlock

        urtVBFB.Connect(True)

        Dim inAuto As ConEnum
        inAuto = urtVBFB.GetElement("inAuto")

        Assert.AreEqual("inAuto", CType(inAuto, IUrtData).Name)
        Assert.AreEqual("inAuto - Desc", CType(inAuto, IUrtData).Description)

        Dim expectedVal As URTVBFunctionBlock.URT.urtNOYES = URTVBFunctionBlock.URT.urtNOYES.NO
        Assert.AreEqual(CInt(expectedVal), inAuto.Val)

    End Sub

    <TestMethod()> Public Sub TestConnectAllElements()

        Dim urtVBFB = New URTVBFunctionBlock.URTVBFunctionBlock

        urtVBFB.Connect(True)

        Dim elementNames() As String = {"ClearError", "FBExecTime",
            "FBErrors", "Note", "outStatus", "inFillEast", "inFillWest", "DivorcedTrip", "MarriedTrip",
            "inAuto", "inTankMarriedMode", "outTankMarriedStatus", "outIsMarried"}

        Dim element As IUrtData

        For Each name In elementNames
            element = urtVBFB.GetElement(name)
            Assert.IsNotNull(element)
            Assert.AreEqual(name, element.Name)
        Next

    End Sub

End Class