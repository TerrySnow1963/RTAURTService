' Title:  			QAL Fill Tank Married Script
' Filename:			QALFillTankMarried1.vb
' Author:			Scott Grimsley (Honeywell) 
' Description:		Fill Tank Married Script.
' Inputs:
' Outputs:      
' REVISION HISTORY 
' ================
' REV:   DATE:       BY:                			DESCRIPTION: 
' 1      2015-08-17  Scott Grimsley                 As Built.
' 2      2015-08-24  Scott Grimsley                 Updated spelling.
' 3      2015-09-18  Terry Snow                     Changed to using SDE for Tank Mode and Status
' 4      2015-10-13  Terry Snow                     Added check for HID valve mode for married / divorced
' 5      2017-06-15  Rakesh Mehta                   Added Note element
' 5.1    2018-01-09  Rakesh Mehta                   Changed to Diagnostic timer
'

Option Explicit On
Option Strict On
Imports System
Imports URT
Imports UrtTlbLib
Imports System.Collections.Generic
Imports Microsoft.VisualBasic
Imports System.Diagnostics

Namespace URT
    Public Enum urtNOYES
        NO = 0
        YES = 1

    End Enum

    Public Enum qalTankMarriedMode
        DIVORCED = 0
        MARRIED = 1
        AUTO = 3
    End Enum

    Public Enum qalTankMarriedStatus
        DIVORCED = 0
        MARRIED = 1
    End Enum


    Public Class VBScript
        Inherits CVBScriptBase
        Private FBErrors As ConString ' Error on Last Run
        Private FBExecTime As ConInt ' Error on Last Run
        Private ClearError As ConBool

        Private inAuto, outIsMarried, inTankMarriedMode, outTankMarriedStatus As ConEnum
        Private inFillEast, inFillWest, DivorcedTrip, MarriedTrip As ConFloat
        Private inDeltaMode As IUrtArray
        Private outStatus As ConString
        ' called when script is recompiled.  bInit is true only the first time the script successfully compiles.
        ' Note that this is not exactly like InitNew() for compiled FBs since bInit is true once each time the platform loads
        ' whereas in compiled FBs, InitNew is called exaclty once when the FB is created and never when the platform loads.
        Sub ConnectDataItems(ByVal bInit As Boolean)
            Try
                '                bInit = QALUtil.SetupVersion3("5.1/2018-01-09", CmpPtr)
                ClearError = QALUtil.URTScale3(Of ConBool)("ClearError", "Clear Error Messages.", GetType(ConBoolClass).GUID, bInit, False, CmpPtr)
                FBErrors = QALUtil.URTScale3(Of ConString)("FBErrors", "Error on Last Run.", GetType(ConStringClass).GUID, bInit, "", CmpPtr)
                FBExecTime = QALUtil.URTScale3(Of ConInt)("FBExecTime", "Run Time of Last FB Execution in Milliseconds.", GetType(ConIntClass).GUID, bInit, 0, CmpPtr)

                QALUtil.URTScale3(Of ConString)("Note", "ToDo - Write Implementation Note", GetType(ConStringClass).GUID, bInit, "", CmpPtr)

                'inAuto = QALUtil.URTScale3(Of ConEnum)("inAuto", "inAuto - Desc", GetType(ConEnumClass).GUID, bInit, False, CmpPtr, "urtNOYES")

                'inTankMarriedMode = QALUtil.URTScale3(Of ConEnum)("inTankMarriedMode", "AUTO: Set by delta level, MARRIED: forced to be married, DIVORCED: Forced to be divorced", _
                '                            GetType(ConEnumClass).GUID, bInit, qalTankMarriedMode.AUTO, CmpPtr, "qalTankMarriedMode")

                'outTankMarriedStatus = QALUtil.URTScale3(Of ConEnum)("outTankMarriedStatus", "Married or Divorced", _
                '                            GetType(ConEnumClass).GUID, bInit, qalTankMarriedStatus.MARRIED, CmpPtr, "qalTankMarriedStatus")

                inFillEast = QALUtil.URTScale3(Of ConFloat)("inFillEast", "inFillEast - Desc", GetType(ConFloatClass).GUID, bInit, 0.0, CmpPtr, , 2228224)
                'inFillWest = QALUtil.URTScale3(Of ConFloat)("inFillWest", "inFillWest - Desc", GetType(ConFloatClass).GUID, bInit, 0.0, CmpPtr, , 2228224)
                'DivorcedTrip = QALUtil.URTScale3(Of ConFloat)("DivorcedTrip", "DivorcedTrip - Desc", GetType(ConFloatClass).GUID, bInit, 5.0, CmpPtr)
                'MarriedTrip = QALUtil.URTScale3(Of ConFloat)("MarriedTrip", "MarriedTrip - Desc", GetType(ConFloatClass).GUID, bInit, 3.0, CmpPtr)
                'inDeltaMode = QALUtil.URTArray3("inDeltaMode", "Array of the HID Feed valve Cascade Modes", GetType(ConArrayBoolClass).GUID, 4, bInit, True, CmpPtr, , 2228224)

                'outIsMarried = QALUtil.URTScale3(Of ConEnum)("outIsMarried", "outIsMarried - Desc", GetType(ConEnumClass).GUID, bInit, False, CmpPtr, "urtNOYES")
                outStatus = QALUtil.URTScale3(Of ConString)("outStatus", "outStatus - Desc", GetType(ConStringClass).GUID, bInit, "Divorced", CmpPtr)

                'If bInit Then
                '    CType(FBErrors, IUrtData).PutSecurityOptions(32, 1, "") ' ReadOnly
                '    CType(FBExecTime, IUrtData).PutSecurityOptions(32, 1, "") ' ReadOnly
                '    CType(inTankMarriedMode, IUrtData).PutSecurityOptions(1, 1, "") ' ReadOnly
                '    CType(outTankMarriedStatus, IUrtData).PutSecurityOptions(32, 1, "") ' ReadOnly
                '    CType(DivorcedTrip, IUrtData).PutSecurityOptions(4, 1, "") ' ReadOnly
                '    CType(MarriedTrip, IUrtData).PutSecurityOptions(4, 1, "") ' ReadOnly

                'End If

            Catch ex As Exception
                Call GenerateError("ConnectDataItems: " & ex.Message)

            End Try

        End Sub

        ' URT scheduler calls PreExecute() then Execute() then PostExecute() each time the schedule executes
        Public Overrides Sub Execute(ByVal iCause As Integer, ByVal pITMScheduler As IUrtTreeMember)

            Dim fbTimer As New StopWatch
            Dim dDelta As Single
            Dim bFlowControlInCasc As Boolean

            Dim I_inDeltaMode() As Boolean = Nothing

            FBErrors.Val = ""
            fbTimer.Start()

            ' Section A:
            ' Copy arrays to local variables.
            Try
                If 4 <> CType(inDeltaMode, IUrtData).Size(urtBUF.dbWORK) Then inDeltaMode.Resize(4, urtBUF.dbWORK)

                I_inDeltaMode = QALUtil.GetArray3(Of Boolean)(inDeltaMode)

            Catch ex As Exception
                GenerateError("A: " & ex.Message)
                Exit Sub

            End Try

            ' Section B:
            ' Loop through each each value:
            Try
                bFlowControlInCasc = (I_inDeltaMode(0) And I_inDeltaMode(1)) Or (I_inDeltaMode(2) And I_inDeltaMode(3))

                If inTankMarriedMode.Val = qalTankMarriedMode.AUTO Then
                    dDelta = Math.Abs(inFillEast.Val - inFillWest.Val)
                    If outIsMarried.Val = urtNOYES.YES Then
                        If dDelta > DivorcedTrip.Val Or bFlowControlInCasc = False Then
                            outIsMarried.Val = urtNOYES.No
                            outStatus.Val = "Divorced"
                            outTankMarriedStatus.Val = qalTankMarriedStatus.DIVORCED
                        End If
                    Else
                        If dDelta < MarriedTrip.Val And bFlowControlInCasc Then
                            outIsMarried.Val = urtNOYES.YES
                            outStatus.Val = "Married"
                            outTankMarriedStatus.Val = qalTankMarriedStatus.MARRIED

                        End If

                    End If
                    'regardless of delta level, 
                Else
                    If inTankMarriedMode.Val = qalTankMarriedMode.MARRIED Then
                        outStatus.Val = "Married"
                        outIsMarried.Val = urtNOYES.YES
                        outTankMarriedStatus.Val = qalTankMarriedSTATUS.MARRIED
                    Else
                        outStatus.Val = "Divorced"
                        outIsMarried.Val = urtNOYES.NO
                        outTankMarriedStatus.Val = qalTankMarriedSTATUS.DIVORCED
                    End If

                End If

            Catch ex As Exception
                GenerateError("B: " & ex.Message)

            End Try

            ' Section Z:
            ' Copy local arrays back to the URT data items.
            Try
                fbTimer.Stop()
                FBExecTime.Val = CInt(fbTimer.elapsedMilliseconds)

            Catch ex As Exception
                GenerateError("Z: " & ex.Message)

            End Try

        End Sub

        ' Title:  			GenerateError
        ' Author:			Scott Grimsley (Honeywell) 
        ' Description:		Sends a message to the operator window only when flag is not set.
        ' REVISION HISTORY 
        ' ================
        ' REV:   DATE:       BY:                			DESCRIPTION: 
        ' 1      2015-04-16  Scott Grimsley                 As Built.
        Private Sub GenerateError(ByVal sError As String)
            Try
                FBErrors.Val += sError
                If ClearError.Val = False Then
                    ClearError.Val = True
                    RaiseMessage(FBErrors.Val)

                End If

            Catch ex As Exception

            End Try

        End Sub

        ' Title:  			RaiseMessage
        ' Author:			Scott Grimsley (Honeywell) 
        ' Description:		Raises a message to the PSOS window.
        ' Inputs:
        '   Message - Message displayed in the message window.
        '   Group - Combines with Priority for the Severity in the message window.
        '   Priority - Combines with Group for the Severity in the message window.
        '   AckRequired - Does the message require an acknowledgement to clear.
        ' Example:
        '  RaiseMessage("Section C: " & ex.Message, urtMSGGROUP.msgERROR, urtMSGPRIORITY.msgHI, False)
        ' Group:
        '   0 - msgSEVERE
        '   1 - msgERROR
        '   2 - msgWARN
        '   3 - msgINFO
        '   4 - msgTRACE
        '   5 - msgVERBOSE
        ' Priority:
        '   0 - msgLO
        '   1 - msgMED
        '   2 - msgHI
        ' REVISION HISTORY 
        ' ================
        ' REV:   DATE:       BY:                			DESCRIPTION: 
        ' 1      2015-04-16  Scott Grimsley                 As Built.
        Private Sub RaiseMessage( _
                ByVal Message As String, _
                Optional ByVal Group As urtMSGGROUP = urtMSGGROUP.msgERROR, _
                Optional ByVal Priority As urtMSGPRIORITY = urtMSGPRIORITY.msgHI, _
                Optional ByVal AckRequired As Boolean = False)
            Dim conMessage As New ConMessageClass
            Dim cookie As Integer
            Try
                conMessage.text = Message
                conMessage.Group = Group
                conMessage.Priority = Priority
                conMessage.AckRequired = AckRequired
                conMessage.type = 1 ' 1 = simple
                DirectCast(CmpPtr(), IUrtMemberSupport).Raise(conMessage, (cookie))

            Catch ex As Exception

            End Try

        End Sub

        Public Overrides Sub OnPreDestruct()
            ' Do not delete.

        End Sub

    End Class


    Public Module QALUtil
        ' Title:  			SetupVersion
        ' Author:			Scott Grimsley (Honeywell) 
        ' Description:		Creates a data item which can be used to determine if the platform is being added for the first time.
        '                   Used to return the status of the bInit flag in ConnectDataItems.
        ' REVISION HISTORY 
        ' ================
        ' REV:   DATE:       BY:                			DESCRIPTION: 
        ' 1      2015-04-16  Scott Grimsley                 As Built.
        Public Function SetupVersion3( _
                            ByVal VersionString As String, _
                            ByVal myCmpPtr As IUrtTreeMember) As Boolean

            Dim bReturn As Boolean = False
            Dim myVersion As ConString        ' Function Block version
            Try
                myVersion = URTScale3(Of ConString)("Version", "Function Block version", GetType(ConStringClass).GUID, False, "", myCmpPtr)
                bReturn = String.IsNullOrEmpty(myVersion.Val)
                myVersion.Val = VersionString
                CType(myVersion, IUrtData).PutSecurityOptions(32, 1, "") ' ReadOnly

            Catch ex As Exception
                Throw ex

            End Try
            Return bReturn

        End Function

        ' Title:  			GetArray
        ' Author:			Scott Grimsley (Honeywell) 
        ' Description:		Returns an iUrtArray of values with the defined type.\
        '           ConArrayBool	Boolean
        '           ConArrayDouble	Double
        '           ConArrayEnum	Integer
        '           ConArrayFloat	Single
        '           ConArrayInt	Integer
        '           ConArrayString	String
        '           ConArrayTime	DateTime
        ' REVISION HISTORY 
        ' ================
        ' REV:   DATE:       BY:                			DESCRIPTION: 
        ' 1      2015-04-16  Scott Grimsley                 As Built.
        Public Function GetArray3(Of T)(ByVal inArray As iUrtArray) As T()
            Dim oTemp As Object = Nothing
            Try
                inArray.GetArray(oTemp, urtBUF.dbWORK)

            Catch ex As Exception
                Throw ex

            End Try
            Return CType(oTemp, T())

        End Function

        ' Title:  			URTArray
        ' Author:			Scott Grimsley (Honeywell) 
        ' Description:		
        ' Creates an array data item if not created.
        ' Establishes a connection to an array data item if already created.
        ' Specified in the ConnectDataItems method.
        ' Inputs:
        '   sName - URT Variable Name as viewed in URT explorer.
        '   sDesc - URT Variable description as viewed in URT explorer.
        '   myType - Guid of URT data type eg GetType(ConFloatClass).GUID.
        '   iSize - Array size.
        '   bInit - True only when component is added to a platform. Leave as bInit.
        '   InitVal - Initial value for all array elements when added to the platform.
        '   ITM - Tree member to add item to. Allows for adding items to other tree structures.
        '   myOption - URT variable option settings.  For more information refer to URTSettings documentation.
        '   EnumType (Enum only) - String type of URT Enumeration. Enumeration types can be found in C:\ProgramData\Honeywell\URT\Enumerations.
        ' Example:
        '  varName = URTNew.URTArray("VarName", "Description", GetType(ConFloatClass).GUID, 12, bInit, 1.0)
        ' REVISION HISTORY 
        ' ================
        ' REV:   DATE:       BY:                			DESCRIPTION: 
        ' 1      2015-04-16  Scott Grimsley                 As Built.
        ' 2      2015-05-29  Scott Grimsley                 Moved settings out of bInit.
        Public Function URTArray3( _
                   ByVal sName As String, _
                   ByVal sDesc As String, _
                   ByVal myType As System.Guid, _
                   ByVal iSize As Integer, _
                   ByVal bInit As Boolean, _
                   ByVal InitVal As Object, _
                   ByVal myCmpPtr As IUrtTreeMember, _
                    Optional ByVal EnumType As String = "", _
                   Optional ByVal myOption As Integer = conOPTIONS.doDEFAULTS) _
                       As IUrtArray
            Dim myArray As IUrtArray = Nothing
            Dim ii As Integer
            Try
                myArray = CType(CUrtFBBase.Setup(sName, sDesc, myCmpPtr, myType, iSize), iUrtArray)
                If EnumType <> "" Then CType(myArray, IUrtEnum).EnumType = EnumType
                If myOption <> conOPTIONS.doDEFAULTS Then URTSettings3(CType(myArray, IUrtData), myOption)
                If bInit Then ' Called when function block first added to URT platform.
                    Dim i_Temp(iSize - 1) As Object
                    For ii = 0 To iSize - 1
                        i_Temp(ii) = InitVal

                    Next
                    myArray.PutArray(i_Temp, "")

                End If

            Catch ex As Exception
                Throw ex

            End Try
            Return myArray

        End Function

        ' Title:  			URTArray
        ' Author:			Scott Grimsley (Honeywell) 
        ' Description:		
        ' Creates an array data item if not created.
        ' Establishes a connection to an array data item if already created.
        ' The connection will be made to the tree member specified in the myTree data item.
        ' This item will generally be a ConArrayComp.
        ' Inputs (where different from above):
        '   iIndex - Index of the ConArrayComp to add this item too. Existing items in that index will be overwritten.
        '   myTree - Tree member to add this item to.
        ' Example:
        '  varName = URTNew.URTArray2("VarName", "Description", GetType(ConFloatClass).GUID, 0, CType(GM, UrtTlbLib.IUrtTreeMember), 12, bInit, 1.0)
        ' REVISION HISTORY 
        ' ================
        ' REV:   DATE:       BY:                			DESCRIPTION: 
        ' 1      2015-04-16  Scott Grimsley                 As Built.
        Public Function URTArray3( _
                                    ByVal sName As String, _
                                    ByVal sDesc As String, _
                                    ByVal myType As System.Guid, _
                                    ByVal iIndex As Integer, _
                                    ByVal myTree As IUrtTreeMember, _
                                    ByVal iSize As Integer, _
                                    ByVal bInit As Boolean, _
                                    ByVal InitVal As Object, _
                                    Optional ByVal EnumType As String = "", _
                                    Optional ByVal myOption As Integer = conOPTIONS.doDEFAULTS) _
                                       As iUrtArray
            Dim myArray As IUrtArray = Nothing
            Dim o As Object = Nothing
            Dim ii As Integer
            Dim GuidITM As Guid = GetType(UrtTlbLib.IUrtTreeMember).GUID
            Try
                If CType(myTree, iUrtData).Size < iIndex + 1 Then CType(myTree, IUrtArray).Resize(iIndex + 1, urtBUF.dbWORK)
                CUrtFBBase.Lib.urtGetItem(sName, myType, myTree, iIndex, Nothing, GuidITM, o, sDesc, 0, 0, "", True, False)
                myArray = CType(o, iUrtArray)
                If CType(myArray, iUrtData).Size <> iSize Then myArray.Resize(iSize, urtBUF.dbWORK)
                If EnumType <> "" Then CType(myArray, IUrtEnum).EnumType = EnumType
                If myOption <> conOPTIONS.doDEFAULTS Then URTSettings3(CType(myArray, IUrtData), myOption)
                If bInit Then ' Called when function block first added to URT platform.
                    Dim i_Temp(iSize - 1) As Object
                    For ii = 0 To iSize - 1
                        i_Temp(ii) = InitVal

                    Next
                    myArray.PutArray(i_Temp, "")

                End If

            Catch ex As Exception
                Throw ex

            End Try
            Return myArray

        End Function

        ' Title:  			URTScale
        ' Author:			Scott Grimsley (Honeywell) 
        ' Description:		
        ' Creates a scalar data item if not created.
        ' Establishes a connection to a scalar data item if already created.
        ' For more information refer to URTArray documentation.
        ' Example:
        '  varName = URTNew.URTScale(Of ConInt)("VarName", "Description",GetType(ConIntClass).GUID, bInit,25)
        ' REVISION HISTORY 
        ' ================
        ' REV:   DATE:       BY:                			DESCRIPTION: 
        ' 1      2015-04-16  Scott Grimsley                 As Built.
        ' 2      2015-05-29  Scott Grimsley                 Moved settings out of bInit.
        Public Function URTScale3(Of T)( _
                                    ByVal sName As String, _
                                    ByVal sDesc As String, _
                                    ByVal myType As System.Guid, _
                                    ByVal bInit As Boolean, _
                                    ByVal InitVal As Object, _
                                    ByVal myCmpPtr As IUrtTreeMember, _
                                    Optional ByVal EnumType As String = "", _
                                    Optional ByVal myOption As Integer = conOPTIONS.doDEFAULTS) _
                                        As T
            Dim myScale As IUrtData = Nothing
            Try
                myScale = CType(CUrtFBBase.Setup(sName, sDesc, myCmpPtr, myType, -1), IUrtData)
                If EnumType <> "" Then CType(myScale, IUrtEnum).EnumType = EnumType
                If myOption <> conOPTIONS.doDEFAULTS Then URTSettings3(myScale, myOption)
                If bInit Then ' Called when function block first added to URT platform.
                    myScale.PutVariantValue(InitVal, "")

                End If

            Catch ex As Exception
                Throw ex

            End Try
            Return CType(myScale, T)

        End Function

        ' Title:  			URTScale
        ' Author:			Scott Grimsley (Honeywell) 
        ' Description:		
        ' Creates an scalar data item if not created.
        ' Establishes a connection to an array data item if already created.
        ' The connection will be made to the tree member specified in the myTree data item.
        ' This item will generally be a ConArrayComp.
        ' Inputs (where different from above):
        '   iIndex - Index of the ConArrayComp to add this item too. Existing items in that index will be overwritten.
        '   myTree - Tree member to add this item to.
        ' Example:
        '  varName = URTNew.URTScale(Of ConInt)("VarName", "Description", GetType(ConIntClass).GUID, 0, CType(GM, UrtTlbLib.IUrtTreeMember), bInit,25)
        ' REVISION HISTORY 
        ' ================
        ' REV:   DATE:       BY:                			DESCRIPTION: 
        ' 1      2015-04-16  Scott Grimsley                 As Built.
        Public Function URTScale3(Of T)( _
                                    ByVal sName As String, _
                                    ByVal sDesc As String, _
                                    ByVal myType As System.Guid, _
                                    ByVal iIndex As Integer, _
                                    ByVal myTree As IUrtTreeMember, _
                                    ByVal bInit As Boolean, _
                                    ByVal InitVal As Object, _
                                    Optional ByVal EnumType As String = "", _
                                    Optional ByVal myOption As Integer = conOPTIONS.doDEFAULTS) _
                                       As T
            Dim o As Object = Nothing
            Dim GuidITM As Guid = GetType(UrtTlbLib.IUrtTreeMember).GUID
            Try
                If CType(myTree, iUrtData).Size < iIndex + 1 Then CType(myTree, IUrtArray).Resize(iIndex + 1, urtBUF.dbWORK)
                CUrtFBBase.Lib.urtGetItem(sName, myType, myTree, iIndex, Nothing, GuidITM, o, sDesc, 0, 0, "", True, False)
                If EnumType <> "" Then CType(o, IUrtEnum).EnumType = EnumType
                If myOption <> conOPTIONS.doDEFAULTS Then URTSettings3(CType(o, iUrtData), myOption)
                If bInit Then ' Called when function block first added to URT platform.
                    CType(o, iUrtData).PutVariantValue(InitVal, "")

                End If

            Catch ex As Exception
                Throw ex

            End Try
            Return CType(o, T)

        End Function

        ' Title:  			URTSettings
        ' Author:			Scott Grimsley (Honeywell) 
        ' Description:		
        ' Enum Member       Internal No.    Bit Item Option
        ' doHISTRETRIEVE    1               1   Retrieve History, Retrieve this item
        ' doHISTCOLLECT     2               2   Collect History, Collect this item
        ' doCHGDETECT       4               3   Output, If Changed
        ' doINBUF           8               4   Input, Input Buffer
        ' doOUTBUF          16              5   Output, Output Buffer
        ' doHISTMANUALINPUT 32              6   Collect History, Manual Input
        ' doCPDONTSAVE      64              7   Check Point, Don't Save
        ' doCPDONTLOAD      128             8   Check Point, Don't Load
        ' doCPDONTSAVEVALS  256             9   Check Point, Don't Save Values
        ' doCPDONTLOADVALS  512             10  Check Point, Don't Load Values
        ' doCONRESIZE       1024            11  Input, Resize According to Target
        ' doLGVHOLD         2048            12  --- (reserved?)
        ' doUSESCHLGVHOLD   4096            13  Input, Use LGVHold of Schedule
        ' doDISABLEINPUT    8192            14  Input, Disabled
        ' doDISABLEOUTPUT   16384           15  Output, Disabled
        ' doMANSUB          32768           16  Value, Manual Substitute(Only when Allow Manual Substitute is checked)
        ' doALLOWMANSUB     65536           17  Value, Allow Manual Substitute
        ' doALLOWNAN        131072          18  Value, Allow Bad (NaN
        ' doOPCTIME         262144          19  Server timestamp via OPC
        ' doALLOWOUTIFBAD   524288          20  Output, Allow Writew/ Bad Quality
        ' doINIFCHANGED     1048576         21  Input, If Changed
        ' doQUALITY         2097152         22  Maintain Quality/Timestamp
        ' doAREA            4194304         23  Define Area
        ' doNANIFBAD        8388608         24  Value, Set NaN on Bad Quality
        ' dsSUSPENDWRITE    16777216        -   --- (not used?)
        ' dsISINPUT         33554432        -   --- (not used?)
        ' dsISOUTPUT        67108864        -   --- (not used?)
        ' dsISWORKING       134217728       -   --- (not used?)
        ' doDEFAULTS        0               1   No option at all
        ' doALLOPTIONS      16777215        -   Set all options
        ' Use this function to set / un-set options for URT variables.
        ' Inputs:
        '   variable - URT variable.
        '   myOption - conOPTION option which is to be set / un-set.  Options can be summed togeather to apply multiple options.
        '   bSet - True for setting the option, False to un-set.
        ' Example:
        '  URTSettings(myBool, conOPTIONS.doOUTBUF + conOPTIONS.doCHGDETECT, TRUE)
        ' REVISION HISTORY 
        ' ================
        ' REV:   DATE:       BY:                			DESCRIPTION: 
        ' 1      2015-04-16  Scott Grimsley                 As Built.
        Public Sub URTSettings3(ByVal myData As IUrtData, _
                              ByVal myOption As Integer, _
                              Optional ByVal bSet As Boolean = True)
            Try
                If bSet Then
                    myData.PutOptions(myOption, myOption, "")

                Else
                    myData.PutOptions(myOption, 0, "")

                End If

            Catch ex As Exception
                Throw ex

            End Try

        End Sub

    End Module


End Namespace


