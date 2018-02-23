'Option Explicit On
'Option Strict On

Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports UrtTlbLib
Imports URTVBQALSimpleScript
Imports URTVBQALSimpleScript.URT

<TestClass()> Public Class UnitTestConBool

    <TestMethod()> Public Sub TestConBoolConstructorAndValueSet()

        Dim myBool As ConBool = New ConBool

        myBool.Val = True
        Assert.IsTrue(myBool.Val)
        myBool.Val = False
        Assert.IsFalse(myBool.Val)

    End Sub

    <TestMethod()> Public Sub TestConBoolConnectWithInitization()
        Dim vbfb As URTVBQALSimpleScript.URT.VBScript = New VBScript

        Dim myBool As ConBool
        Dim bInit As Boolean = True

        Dim desc As String = "myBool - Desc"

        myBool = QALUtil.URTScale3(Of ConBool)("myBool", desc, GetType(ConBoolClass).GUID, bInit, True, vbfb.CmpPtr(), , 2228224)

        myBool.Val = True
        Assert.IsTrue(myBool.Val)
        myBool.Val = False
        Assert.IsFalse(myBool.Val)

        Assert.AreEqual(desc, myBool.Description)

    End Sub

    '    <TestMethod()> Public Sub UnitTestConBoolCastToConBoolClass()
    '        'ToDo need to work out how to relate a ConBool to a ConBoolClass
    '        Dim myBool As ConBool = New ConBool

    '        myBool.Val = True

    '        Dim myConBoolClass As ConBoolClass = myBool

    '        Assert.IsTrue(myBool.Val)
    '    End Sub

End Class

<TestClass()> Public Class UnitTestConInt

    <TestMethod()> Public Sub TestConIntConstructorAndValueSet()

        Dim myInt As ConInt = New ConInt

        myInt.Val = 10
        Assert.AreEqual(10, myInt.Val)
        myInt.Val = -7
        Assert.AreEqual(-7, myInt.Val)

    End Sub

End Class

<TestClass()> Public Class UnitTestConString

    <TestMethod()> Public Sub TestConStringConstructorAndValueSet()

        Dim myString As ConString = New ConString
        Dim expectedString1 = "Some String"
        Dim expectedString2 = "Some Other String"
        myString.Val = expectedString1
        Assert.AreEqual(expectedString1, myString.Val)
        myString.Val = expectedString2
        Assert.AreEqual(expectedString2, myString.Val)

    End Sub

End Class