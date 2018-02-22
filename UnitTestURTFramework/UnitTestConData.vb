'Option Explicit On
'Option Strict On

Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports UrtTlbLib
Imports URTVBQALSimpleScript

<TestClass()> Public Class UnitTestConBool

    <TestMethod()> Public Sub TestConBoolConstructorAndValueSet()

        Dim myBool As ConBool = New ConBool

        myBool.Val = True
        Assert.IsTrue(myBool.Val)
        myBool.Val = False
        Assert.IsFalse(myBool.Val)

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
        myString.val = expectedString1
        Assert.AreEqual(expectedString1, myString.val)
        myString.val = expectedString2
        Assert.AreEqual(expectedString2, myString.val)

    End Sub

    <TestMethod()> Public Sub TestConStringQALUtilConnect()
        Dim vbfb As URTVBQALSimpleScript.URT.VBScript = New URTVBQALSimpleScript.URT.VBScript
        Dim myString As ConString = New ConString


        Dim expectedString1 = "Some String"
        Dim expectedString2 = "Some Other String"
        myString.Val = expectedString1
        Assert.AreEqual(expectedString1, myString.Val)
        myString.Val = expectedString2
        Assert.AreEqual(expectedString2, myString.Val)

    End Sub


End Class