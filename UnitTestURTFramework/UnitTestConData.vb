Option Explicit On
Option Strict On

Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports UrtTlbLib
Imports URTVBQALSimpleScript.URT
Imports RTAInterfaces


<TestClass()> Public Class UnitTestConTypes

    <TestMethod()> Public Sub TestConBoolConstructorAndValueSet()

        Dim myBool As ConBool = New ConBool

        myBool.Val = True
        Assert.IsTrue(myBool.Val)
        myBool.Val = False
        Assert.IsFalse(myBool.Val)

    End Sub

    Private Structure QALUtilConnectParams(Of T)
        Public test1 As T
        Public test2 As T
        Public Sub New(t1 As T, t2 As T)
            test1 = t1
            test2 = t2
        End Sub
    End Structure

    Private Sub TestQALUtilConnect(Of T1 As {con_scalar(Of T3)}, T2, T3)(name As String, params As QALUtilConnectParams(Of T3))
        Dim vbfb As VBScript = New VBScript
        Dim myCon As T1
        Dim myCon1 As IRTAUrtData
        Dim bInit As Boolean = True

        Dim theName As String = name
        Dim theDesc As String = theName & " - Description"

        myCon = QALUtil.URTScale3(Of T1)(theName, theDesc, GetType(T2).GUID, bInit, 0, vbfb.CmpPtr, , 2228224)


        Dim test1 As T3 = params.test1
        Dim test2 As T3 = params.test2

        Dim expected1 As T3 = test1
        Dim expected2 As T3 = test2

        myCon.Val = expected1
        Assert.AreEqual(expected1, myCon.Val)
        myCon.Val = expected2
        Assert.AreEqual(expected2, myCon.Val)

        myCon1 = CType(vbfb.GetElement(theName), IRTAUrtData)
        Assert.AreEqual(expected2, myCon1.Item(0))
    End Sub

    <TestMethod()> Public Sub TestConBoolQALUtilConnect()
        Dim params As QALUtilConnectParams(Of Boolean) = New QALUtilConnectParams(Of Boolean)(True, False)
        TestQALUtilConnect(Of ConBool, ConBoolClass, Boolean)("Bool1", params)
    End Sub

    <TestMethod()> Public Sub TestConIntConstructorAndValueSet()

        Dim myInt As ConInt = New ConInt

        myInt.Val = 10
        Assert.AreEqual(10, myInt.Val)
        myInt.Val = -7
        Assert.AreEqual(-7, myInt.Val)

    End Sub

    <TestMethod()> Public Sub TestConIntQALUtilConnect()
        Dim params As QALUtilConnectParams(Of Integer) = New QALUtilConnectParams(Of Integer)(10, -7)
        TestQALUtilConnect(Of ConInt, ConIntClass, Integer)("Int1", params)
    End Sub

    <TestMethod()> Public Sub TestConStringConstructorAndValueSet()

        Dim myString As ConString = New ConString
        Dim expectedString1 = "Some String"
        Dim expectedString2 = "Some Other String"
        myString.Val = expectedString1
        Assert.AreEqual(expectedString1, myString.Val)
        myString.Val = expectedString2
        Assert.AreEqual(expectedString2, myString.Val)

    End Sub

    <TestMethod()> Public Sub TestConStringQALUtilConnect()
        Dim params As QALUtilConnectParams(Of String) = New QALUtilConnectParams(Of String)("First String", "Second String")
        TestQALUtilConnect(Of ConString, ConStringClass, String)("Bool1", params)
    End Sub

    <TestMethod()> Public Sub TestConFloatConstructorAndValueSet()

        Dim myString As ConFloat = New ConFloat
        Dim expected1 As Single = 12.3
        Dim expected2 As Single = -23.1
        myString.Val = expected1
        Assert.AreEqual(expected1, myString.Val)
        myString.Val = expected2
        Assert.AreEqual(expected2, myString.Val)

    End Sub

    <TestMethod()> Public Sub TestConFloatQALUtilConnect()

        'Todo, using 2 references to TesTscript causes a problem
        Dim params As QALUtilConnectParams(Of Single) = New QALUtilConnectParams(Of Single)(12.1, -3.4))
        TestQALUtilConnect(Of ConFloat, ConFloatClass, Single)("Float1", params)
    End Sub

    <TestMethod()> Public Sub TestConDoubleConstructorAndValueSet()

        Dim myString As ConDouble = New ConDouble
        Dim expected1 As Double = 12.3
        Dim expected2 As Double = -23.1
        myString.Val = expected1
        Assert.AreEqual(expected1, myString.Val)
        myString.Val = expected2
        Assert.AreEqual(expected2, myString.Val)

    End Sub

    <TestMethod()> Public Sub TestConDoubleQALUtilConnect()
        Dim vbfb As VBScript = New VBScript
        Dim params As QALUtilConnectParams(Of Double) = New QALUtilConnectParams(Of Double)
        params.test1 = 12.1
        params.test2 = -3.4
        TestQALUtilConnect(Of ConDouble, ConDoubleClass, Double)("Double1", params)
    End Sub

End Class