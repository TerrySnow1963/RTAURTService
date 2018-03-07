Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports System.Reflection

<TestClass()> Public Class UnitTestScriptProxy

    <TestMethod()> Public Sub TestConnectAssembyToSimpleScript()
        Dim scriptAssembly As Assembly = Assembly.LoadFile("c:\Users\Terry\Documents\Visual Studio 2015\Projects\RTAURTService\URTVBQALSimpleScript\bin\Debug\URTVBQALSimpleScript.dll")


        Assert.IsNotNull(scriptAssembly)

        Dim ScriptFactoryOb As Object

        ScriptFactoryOb = scriptAssembly.CreateInstance("URTVBQALSimpleScript.ScriptProxyFactory")


        Assert.IsNotNull(ScriptFactoryOb)
    End Sub

End Class