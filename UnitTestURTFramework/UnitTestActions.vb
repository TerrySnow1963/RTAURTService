Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports URT
Imports RTAURTService
Imports UrtTlbLib
Imports RTAInterfaces
Imports URTVBQALSimpleScript

<TestClass()> Public Class UnitTestActions

    <TestMethod()> Public Sub TestActionConnect()
        Dim urtVBFB = New URTVBQALSimpleScript.URTVBQALSimpleScript

        Dim ConnectAction As New RTAURTActionConnect
        Dim ConnectActionParams As New RTAURTActionConnectParameters(True)

        ConnectAction.Execute(urtVBFB, ConnectActionParams)

        Assert.AreEqual(7, urtVBFB.GetElements.Count)
        Assert.AreEqual(1, urtVBFB.NumberOfConnectionsExecutions)
    End Sub

    <TestMethod()> Public Sub TestActionConnectAnd1Execute()
        Dim urtVBFB = New URTVBQALSimpleScript.URTVBQALSimpleScript

        Dim ConnectAction As New RTAURTActionConnect
        Dim ConnectActionParams As New RTAURTActionConnectParameters(True)

        Dim ExecuteAction As New RTAURTActionExecute
        Dim ExecuteActionParams As New RTAURTActionExecuteParameters()

        ConnectAction.Execute(urtVBFB, Nothing)

        Assert.AreEqual(7, urtVBFB.GetElements.Count)
        ExecuteAction.Execute(urtVBFB, ExecuteActionParams)

        Assert.AreEqual(1, urtVBFB.NumberOfExecuteExecutions)
    End Sub

    <TestMethod()> Public Sub TestActionConnectAnd5Executes()
        Dim urtVBFB = New URTVBQALSimpleScript.URTVBQALSimpleScript

        Dim ConnectAction As New RTAURTActionConnect
        Dim ConnectActionParams As New RTAURTActionConnectParameters(True)

        Dim NumberOfExecutes As Integer = 5
        Dim ExecuteAction As New RTAURTActionExecute()
        Dim ExecuteActionParams As New RTAURTActionExecuteParameters(NumberOfExecutes)

        ConnectAction.Execute(urtVBFB, Nothing)

        Dim outCounter As ConInt = urtVBFB.GetElement("outCounter")
        Dim InitCounterVal As Integer = 3
        outCounter.Val = InitCounterVal

        Assert.AreEqual(7, urtVBFB.GetElements.Count)
        ExecuteAction.Execute(urtVBFB, ExecuteActionParams)

        Assert.AreEqual(NumberOfExecutes, urtVBFB.NumberOfExecuteExecutions)
        Assert.AreEqual(1, urtVBFB.NumberOfConnectionsExecutions)
        Assert.AreEqual(NumberOfExecutes + InitCounterVal, outCounter.Val)
    End Sub

    <TestMethod()> Public Sub TestActionSetValuesFor1IntValue()
        Dim urtVBFB = New URTVBQALSimpleScript.URTVBQALSimpleScript

        Dim ConnectAction As New RTAURTActionConnect
        Dim ConnectActionParams As New RTAURTActionConnectParameters(True)

        Dim NumberOfExecutes As Integer = 3
        Dim ExecuteAction As New RTAURTActionExecute()
        Dim ExecuteActionParams As New RTAURTActionExecuteParameters(NumberOfExecutes)

        ConnectAction.Execute(urtVBFB, Nothing)

        Dim outCounter As ConInt = urtVBFB.GetElement("outCounter")
        Dim InitCounterVal As Integer = 7

        Dim SetValuesAction As New RTAURTActionSetValues()
        Dim SetValuesActionParams As New RTAURTActionSetValuesParameters()
        SetValuesActionParams.AddValue("outCounter", InitCounterVal)

        SetValuesAction.Execute(urtVBFB, SetValuesActionParams)

        Assert.AreEqual(7, urtVBFB.GetElements.Count)
        ExecuteAction.Execute(urtVBFB, ExecuteActionParams)

        Assert.AreEqual(NumberOfExecutes, urtVBFB.NumberOfExecuteExecutions)
        Assert.AreEqual(1, urtVBFB.NumberOfConnectionsExecutions)
        Assert.AreEqual(InitCounterVal + NumberOfExecutes, outCounter.Val)
    End Sub

    <TestMethod()> Public Sub TestActionSetValuesFor2ValuesIntAndBool()
        Dim urtVBFB = New URTVBQALSimpleScript.URTVBQALSimpleScript

        Dim ConnectAction As New RTAURTActionConnect
        Dim ConnectActionParams As New RTAURTActionConnectParameters(True)

        Dim NumberOfExecutes As Integer = 3
        Dim ExecuteAction As New RTAURTActionExecute()
        Dim ExecuteActionParams As New RTAURTActionExecuteParameters(NumberOfExecutes)

        ConnectAction.Execute(urtVBFB, Nothing)

        Dim outCounter As ConInt = urtVBFB.GetElement("outCounter")
        Dim InitCounterVal As Integer = 7

        Dim SetValuesAction As New RTAURTActionSetValues()
        Dim SetValuesActionParams As New RTAURTActionSetValuesParameters()
        SetValuesActionParams.AddValue("outCounter", InitCounterVal)
        SetValuesActionParams.AddValue("inUseUpCounter", False)

        SetValuesAction.Execute(urtVBFB, SetValuesActionParams)

        Assert.AreEqual(7, urtVBFB.GetElements.Count)
        ExecuteAction.Execute(urtVBFB, ExecuteActionParams)

        Assert.AreEqual(NumberOfExecutes, urtVBFB.NumberOfExecuteExecutions)
        Assert.AreEqual(1, urtVBFB.NumberOfConnectionsExecutions)
        Assert.AreEqual(InitCounterVal - NumberOfExecutes, outCounter.Val)
    End Sub

    <TestMethod()> Public Sub TestActionSetValuesFor1IntArrayValue()
        Dim urtVBFB = New URTVBQALSimpleScript.URTVBQALSimpleScript

        Dim ConnectAction As New RTAURTActionConnect
        Dim ConnectActionParams As New RTAURTActionConnectParameters(True)

        Dim ExecuteAction As New RTAURTActionExecute()
        Dim ExecuteActionParams As New RTAURTActionExecuteParameters()

        ConnectAction.Execute(urtVBFB, Nothing)

        Dim OutFloat2 As ConInt = urtVBFB.GetElement("OutFloat2")

        Dim newElementValue As Single = 10.3!

        Dim SetValuesAction As New RTAURTActionSetValues()
        Dim SetValuesActionParams As New RTAURTActionSetValuesParameters()
        SetValuesActionParams.AddValue("inArrayFloat", newElementValue)

        SetValuesAction.Execute(urtVBFB, SetValuesActionParams)

        Assert.AreEqual(7, urtVBFB.GetElements.Count)
        ExecuteAction.Execute(urtVBFB, ExecuteActionParams)

        Assert.AreEqual(newElementValue + 1.0 * 3, OutFloat2.Val)
    End Sub
End Class