Option Explicit On
Option Strict On

Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports UrtTlbLib
Imports URTVBQALSimpleScript
Imports URTVBQALSimpleScript.URT
Imports RTAInterfaces
Imports RTAURTService
Imports URT


<TestClass()> Public Class UnitTestConTypes

    Private Structure QALUtilConnectParams(Of T)
        Public test1 As T
        Public test2 As T
        Public Sub New(t1 As T, t2 As T)
            test1 = t1
            test2 = t2
        End Sub
    End Structure
    Private Sub TestConDataConstructorAndValueSet(Of T1 As {New, con_scalar(Of T3, T2)}, T2, T3)(params As QALUtilConnectParams(Of T3))

        Dim myScalar As T1 = New T1

        Dim test1 As T3 = params.test1
        Dim test2 As T3 = params.test2

        Dim expected1 As T3 = test1
        Dim expected2 As T3 = test2

        myScalar.Val = expected1
        Assert.AreEqual(expected1, myScalar.Val)
        myScalar.Val = expected2
        Assert.AreEqual(expected2, myScalar.Val)

    End Sub
    Private Sub TestQALUtilConnect(Of T1 As {con_scalar(Of T3, T2)}, T2, T3)(name As String, params As QALUtilConnectParams(Of T3))
        Dim logger As RTAUrtTraceLogger = New RTAUrtTraceLogger
        Dim vbfb As URTVBQALSimpleScript.URTVBQALSimpleScript = New URTVBQALSimpleScript.URTVBQALSimpleScript(logger)
        Dim myCon As T1
        Dim myCon1 As IRTAUrtData
        Dim bInit As Boolean = True

        Dim theName As String = name
        Dim theDesc As String = theName & " - Description"

        myCon = QALUtil.URTScale3(Of T1)(theName, theDesc, GetType(T2).GUID, bInit, params.test2, vbfb.CmpPtr, , 2228224)


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
        Assert.AreEqual(0, logger.Count)
    End Sub

#Region "Scalar tests"
    <TestMethod()> Public Sub TestConBoolConstructorAndValueSet()
        Dim params As QALUtilConnectParams(Of Boolean) = New QALUtilConnectParams(Of Boolean)(True, False)
        TestConDataConstructorAndValueSet(Of ConBool, ConBoolClass, Boolean)(params)
    End Sub

    <TestMethod()> Public Sub TestConBoolQALUtilConnect()
        Dim params As QALUtilConnectParams(Of Boolean) = New QALUtilConnectParams(Of Boolean)(True, False)
        TestQALUtilConnect(Of ConBool, ConBoolClass, Boolean)("Bool1", params)
    End Sub

    <TestMethod()> Public Sub TestConIntConstructorAndValueSet()
        Dim params As QALUtilConnectParams(Of Integer) = New QALUtilConnectParams(Of Integer)(10, -7)
        TestConDataConstructorAndValueSet(Of ConInt, ConIntClass, Integer)(params)
    End Sub

    <TestMethod()> Public Sub TestConIntQALUtilConnect()
        Dim params As QALUtilConnectParams(Of Integer) = New QALUtilConnectParams(Of Integer)(10, -7)
        TestQALUtilConnect(Of ConInt, ConIntClass, Integer)("Int1", params)
    End Sub

    <TestMethod()> Public Sub TestConStringConstructorAndValueSet()
        Dim params As QALUtilConnectParams(Of String) = New QALUtilConnectParams(Of String)("First String", "Second String")
        TestConDataConstructorAndValueSet(Of ConString, ConStringClass, String)(params)
    End Sub

    <TestMethod()> Public Sub TestConStringQALUtilConnect()
        Dim params As QALUtilConnectParams(Of String) = New QALUtilConnectParams(Of String)("First String", "Second String")
        TestQALUtilConnect(Of ConString, ConStringClass, String)("Bool1", params)
    End Sub

    <TestMethod()> Public Sub TestConFloatConstructorAndValueSet()
        Dim params As QALUtilConnectParams(Of Single) = New QALUtilConnectParams(Of Single)(12.1, -3.4)
        TestConDataConstructorAndValueSet(Of ConFloat, ConFloatClass, Single)(params)
    End Sub

    <TestMethod()> Public Sub TestConFloatQALUtilConnect()
        Dim params As QALUtilConnectParams(Of Single) = New QALUtilConnectParams(Of Single)(12.1, -3.4)
        TestQALUtilConnect(Of ConFloat, ConFloatClass, Single)("Float1", params)
    End Sub

    <TestMethod()> Public Sub TestConDoubleConstructorAndValueSet()
        Dim params As QALUtilConnectParams(Of Double) = New QALUtilConnectParams(Of Double)(12.1, -3.4)
        TestConDataConstructorAndValueSet(Of ConDouble, ConDoubleClass, Double)(params)
    End Sub

    <TestMethod()> Public Sub TestConDoubleQALUtilConnect()
        Dim params As QALUtilConnectParams(Of Double) = New QALUtilConnectParams(Of Double)(12.1, -3.4)
        TestQALUtilConnect(Of ConDouble, ConDoubleClass, Double)("Double1", params)
    End Sub

    <TestMethod()> Public Sub TestConTimeQALUtilConnect()
        Dim params As QALUtilConnectParams(Of DateTime) = New QALUtilConnectParams(Of DateTime)(Now(), Now.AddHours(-3.4))
        TestQALUtilConnect(Of ConTime, ConTimeClass, DateTime)("Time1", params)
    End Sub

    <TestMethod()> Public Sub TestConTimeConstructorAndValueSet()
        Dim params As QALUtilConnectParams(Of DateTime) = New QALUtilConnectParams(Of DateTime)(Now(), Now.AddHours(-3.4))
        TestConDataConstructorAndValueSet(Of ConTime, ConTimeClass, DateTime)(params)
    End Sub

    <TestMethod()> Public Sub TestConShortQALUtilConnect()
        Dim params As QALUtilConnectParams(Of Short) = New QALUtilConnectParams(Of Short)(2, 6)
        TestQALUtilConnect(Of ConShort, ConShortClass, Short)("Double1", params)
    End Sub
    <TestMethod()> Public Sub TestConShortConstructorAndValueSet()
        Dim params As QALUtilConnectParams(Of Short) = New QALUtilConnectParams(Of Short)(2, 6)
        TestQALUtilConnect(Of ConShort, ConShortClass, Short)("Short1", params)
    End Sub
#End Region

#Region "ConArrayHelper Methods"
    Public Sub TestConArrayTypeConstructorResize(Of T1 As {New, con_array(Of T3, T2)}, T2, T3)(size As Integer)
        Dim inArray As T1 = New T1
        Dim newSize As Integer = size

        If newSize <> CType(inArray, T1).Size(urtBUF.dbWork) Then inArray.Resize(newSize, urtBUF.dbWork)
        Assert.AreEqual(newSize, inArray.Size(urtBUF.dbWork))

    End Sub
    Public Sub TestConArrayTypeConstructorSetandGetVal(Of T1 As {New, con_array(Of T3, T2)}, T2, T3)(initVals() As T3)
        Dim inArray As T1 = New T1

        'Dim initVals() As Boolean = {True, False, True, False}

        Dim newSize As Integer = initVals.Length
        If newSize <> CType(inArray, T1).Size(urtBUF.dbWork) Then inArray.Resize(newSize, urtBUF.dbWork)

        Dim I_inArray() As T3 = Nothing
        Dim I_outArray() As T3 = Nothing
        I_inArray = QALUtil.GetArray3(Of T3)(inArray)

        Assert.AreEqual(newSize, I_inArray.Length)

        For ii = 0 To newSize - 1
            I_inArray(ii) = initVals(ii)
        Next

        CType(inArray, IURTArray).PutArray(I_inArray, "")

        I_outArray = QALUtil.GetArray3(Of T3)(inArray)
        For ii = 0 To newSize - 1
            Assert.AreEqual(initVals(ii), I_outArray(ii))
        Next

    End Sub
#End Region

#Region "ConArrayBool"
    <TestMethod()> Public Sub TestConArrayBoolConstructorZeroSize()
        Dim inArrayBool As ConArrayBool = New ConArrayBool
        Assert.AreEqual(0, inArrayBool.Size(urtBUF.dbWork))
    End Sub

    <TestMethod()> Public Sub TestConArrayBoolConstructorResize()
        TestConArrayTypeConstructorResize(Of ConArrayBool, ConArrayBoolClass, Boolean)(4)

    End Sub

    <TestMethod()> Public Sub TestConArrayBoolConstructorSetandGetVal()
        TestConArrayTypeConstructorSetandGetVal(Of ConArrayBool, ConArrayBoolClass, Boolean)({True, False, True, False})

    End Sub
#End Region

#Region "ConArrayFloat"
    <TestMethod()> Public Sub TestConArrayFloatConstructorZeroSize()
        Dim inArray As ConArrayFloat = New ConArrayFloat
        Assert.AreEqual(0, inArray.Size(urtBUF.dbWork))
    End Sub

    <TestMethod()> Public Sub TestConArrayFloatConstructorResize()
        TestConArrayTypeConstructorResize(Of ConArrayFloat, ConArrayFloatClass, Single)(4)
    End Sub

    <TestMethod()> Public Sub TestConArrayFloatConstructorSetandGetVal()
        TestConArrayTypeConstructorSetandGetVal(Of ConArrayDouble, ConArrayDoubleClass, Double)({1.3, 4.5, 6.7, 8.4})
    End Sub
#End Region

#Region "ConArrayDouble"
    <TestMethod()> Public Sub TestConArrayDoubleConstructorZeroSize()
        Dim inArray As ConArrayDouble = New ConArrayDouble
        Assert.AreEqual(0, inArray.Size(urtBUF.dbWork))
    End Sub

    <TestMethod()> Public Sub TestConArrayDoubleConstructorResize()
        TestConArrayTypeConstructorResize(Of ConArrayDouble, ConArrayDoubleClass, Double)(4)
    End Sub

    <TestMethod()> Public Sub TestConArrayDoubleConstructorSetandGetVal()
        TestConArrayTypeConstructorSetandGetVal(Of ConArrayDouble, ConArrayDoubleClass, Double)({1.3, 4.5, 6.7, 8.4})
    End Sub
#End Region

#Region "ConArrayInt"
    <TestMethod()> Public Sub TestConArrayIntConstructorZeroSize()
        Dim inArray As ConArrayInt = New ConArrayInt
        Assert.AreEqual(0, inArray.Size(urtBUF.dbWork))
    End Sub

    <TestMethod()> Public Sub TestConArrayIntConstructorResize()
        TestConArrayTypeConstructorResize(Of ConArrayInt, ConArrayIntClass, Integer)(4)
    End Sub


    <TestMethod()> Public Sub TestConArrayIntConstructorSetandGetVal()
        TestConArrayTypeConstructorSetandGetVal(Of ConArrayInt, ConArrayIntClass, Integer)({1, 5, 6, -8})
    End Sub
#End Region

#Region "ConArrayString"
    <TestMethod()> Public Sub TestConArrayStringConstructorZeroSize()
        Dim inArray As ConArrayString = New ConArrayString
        Assert.AreEqual(0, inArray.Size(urtBUF.dbWork))
    End Sub

    <TestMethod()> Public Sub TestConArrayStringConstructorResize()
        TestConArrayTypeConstructorResize(Of ConArrayString, ConArrayStringClass, String)(4)
    End Sub

    <TestMethod()> Public Sub TestConArrayStringConstructorSetandGetVal()
        TestConArrayTypeConstructorSetandGetVal(Of ConArrayString, ConArrayStringClass, String)({"1", "5", "6", "-8"})
    End Sub
#End Region

#Region "ConArrayEnum"
    <TestMethod()> Public Sub TestConArrayEnumConstructorZeroSize()
        Dim inArray As ConArrayEnum = New ConArrayEnum
        Assert.AreEqual(0, inArray.Size(urtBUF.dbWork))
    End Sub

    <TestMethod()> Public Sub TestConArrayEnumConstructorResize()
        TestConArrayTypeConstructorResize(Of ConArrayEnum, ConArrayEnumClass, Integer)(4)
    End Sub

    <TestMethod()> Public Sub TestConArrayEnumConstructorSetandGetVal()
        TestConArrayTypeConstructorSetandGetVal(Of ConArrayEnum, ConArrayEnumClass, Integer)({1, 5, 6, 8})
    End Sub
#End Region

#Region "ConArrayTime"
    <TestMethod()> Public Sub TestConArrayTimeConstructorZeroSize()
        Dim inArray As ConArrayTime = New ConArrayTime
        Assert.AreEqual(0, inArray.Size(urtBUF.dbWork))
    End Sub

    <TestMethod()> Public Sub TestConArrayTimeConstructorResize()
        TestConArrayTypeConstructorResize(Of ConArrayTime, ConArrayTimeClass, DateTime)(4)
    End Sub

    <TestMethod()> Public Sub TestConArrayTimeConstructorSetandGetVal()
        TestConArrayTypeConstructorSetandGetVal(Of ConArrayTime, ConArrayTimeClass, DateTime)({Now(), Now.AddDays(-1), Now.AddHours(-12), Now.AddMinutes(-24)})
    End Sub
#End Region

#Region "ConArrayShort"
    <TestMethod()> Public Sub TestConArrayShortConstructorZeroSize()
        Dim inArray As ConArrayShort = New ConArrayShort
        Assert.AreEqual(0, inArray.Size(urtBUF.dbWork))
    End Sub

    <TestMethod()> Public Sub TestConArrayShortConstructorResize()
        TestConArrayTypeConstructorResize(Of ConArrayShort, ConArrayShortClass, Short)(4)
    End Sub


    <TestMethod()> Public Sub TestConArrayShortConstructorSetandGetVal()
        TestConArrayTypeConstructorSetandGetVal(Of ConArrayShort, ConArrayShortClass, Short)({14, 45, 66, 118})
    End Sub
#End Region
End Class