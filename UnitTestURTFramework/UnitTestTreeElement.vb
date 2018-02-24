﻿Option Explicit On
Option Strict On

Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports URTVBQALSimpleScript
Imports URTVBQALSimpleScript.URT
Imports RTAInterfaces
Imports RTAURTService
Imports UrtTlbLib

<TestClass()> Public Class UnitTestTreeElement

    <TestMethod()> Public Sub TestCmpPtrGetElementsGivesAllElements()
        Dim logger As RTAUrtTraceLogger = New RTAUrtTraceLogger
        Dim urtVBFB = New URTVBQALSimpleScript.URTVBQALSimpleScript(logger)

        urtVBFB.Connect(True)

        Dim elementNames() As String = {"inFloat1", "outFloat1", "inArrayBool", "inSize", "inArrayFloat", "outCounter", "inUseUpCounter"}

        For Each item In CType(urtVBFB, IRTAUrtTreeMember).GetElements
            Trace.WriteLine("Found " & item.Name)
            Assert.IsTrue(elementNames.Contains(item.Name))
        Next

    End Sub

End Class