' Title:  			QAL Simple Script
' Filename:			QALSimpleScript.vb
' Author:			Terry Snow (RTA) 
' Description:		Used to test the URT Test BEd
' Inputs:
' Outputs:      
' REVISION HISTORY 
' ================
' REV:   DATE:       BY:                			DESCRIPTION: 
' 1      2018-02-17  Terry Snow                     As Built.
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

    Public Class VBScript
        Inherits CVBScriptBase

        Private inFloat1, inFloat2PointsToinFloat1, outFloat1, outFloat2 As ConFloat
        Private inSize, outCounter As ConInt
        Private inUseUpCounter As ConBool
        Private inArrayBool, inArrayFloat As IURTArray

        ' called when script is recompiled.  bInit is true only the first time the script successfully compiles.
        ' Note that this is not exactly like InitNew() for compiled FBs since bInit is true once each time the platform loads
        ' whereas in compiled FBs, InitNew is called exaclty once when the FB is created and never when the platform loads.
        Sub ConnectDataItems(ByVal bInit As Boolean)

            Dim nSize As Integer = 4
            Try
                '                bInit = QALUtil.SetupVersion3("5.1/2018-01-09", CmpPtr)

                'Note have deliberataly made inFloat2 the same element as inFloat1
                inFloat1 = QALUtil.URTScale3(Of ConFloat)("inFloat1", "inFloat1 - Desc", GetType(ConFloatClass).GUID, bInit, 0.0, CmpPtr, , 2228224)
                inFloat2PointsToinFloat1 = QALUtil.URTScale3(Of ConFloat)("inFloat1", "inFloat2 - Desc", GetType(ConFloatClass).GUID, bInit, 0.0, CmpPtr, , 2228224)
                inArrayBool = QALUtil.URTArray3("inArrayBool", "inArrayBool - Descs", GetType(ConArrayBoolClass).GUID, nSize, bInit, True, CmpPtr, , 2228224)
                inSize = QALUtil.URTScale3(Of ConInt)("inSize", "inSize - Desc", GetType(ConIntClass).GUID, bInit, nSize, CmpPtr, , 2228224)
                outFloat1 = QALUtil.URTScale3(Of ConFloat)("outFloat1", "outFloat1 - Desc", GetType(ConFloatClass).GUID, bInit, 5.0, CmpPtr)
                outFloat2 = QALUtil.URTScale3(Of ConFloat)("outFloat2", "outFloat1 - Desc", GetType(ConFloatClass).GUID, bInit, 0.0, CmpPtr)
                inArrayFloat = QALUtil.URTArray3("inArrayFloat", "Array of Input Values", GetType(ConArrayFloatClass).GUID, nSize, bInit, 1.0, CmpPtr, , 2228224)
                inUseUpCounter = QALUtil.URTScale3(Of ConBool)("inUseUpCounter", "If True: Count Up, Else Count Down", GetType(ConBoolClass).GUID, bInit, True, CmpPtr, , 2228224)
                outCounter = QALUtil.URTScale3(Of ConInt)("outCounter", "outCounter - Desc", GetType(ConIntClass).GUID, bInit, nSize, CmpPtr, , 2228224)

            Catch ex As Exception
                Call GenerateError("ConnectDataItems: " & ex.Message)

            End Try

        End Sub

        ' URT scheduler calls PreExecute() then Execute() then PostExecute() each time the schedule executes
        Public Overrides Sub Execute(ByVal iCause As Integer, ByVal pITMScheduler As IUrtTreeMember)

            Dim I_inArrayBool() As Boolean = Nothing
            Dim I_inArrayFloat() As Single = Nothing
            Dim total As Single = 0

            Try

                If inSize.Val < 1 Then inSize.Val = 1

                If inSize.Val <> CType(inArrayBool, IUrtData).Size(urtBUF.dbWork) Then inArrayBool.Resize(inSize.Val, urtBUF.dbWork)
                If inSize.Val <> CType(inArrayFloat, IUrtData).Size(urtBUF.dbWork) Then inArrayFloat.Resize(inSize.Val, urtBUF.dbWork)

                I_inArrayBool = QALUtil.GetArray3(Of Boolean)(inArrayBool)
                I_inArrayFloat = QALUtil.GetArray3(Of Single)(inArrayFloat)

            Catch ex As Exception
                GenerateError("A: " & ex.Message)
                Exit Sub

            End Try

            ' Section B:
            ' Loop through each each value:
            Try
                'as inFloat1 is the same as inFloat2, then this is really the same as 2*infloat1

                If I_inArrayBool(0) And I_inArrayBool(1) Then
                    outFloat1.Val = inFloat1.Val + inFloat2PointsToinFloat1.Val
                Else
                    outFloat1.Val = inFloat1.Val * 3.0!
                End If

                For ii = 0 To inSize.Val - 1
                    total += I_inArrayFloat(ii)
                Next

            Catch ex As Exception
                GenerateError("B: " & ex.Message)
            End Try

            ' Section Z:
            ' Copy local arrays back to the URT data items.
            Try
                If inUseUpCounter.Val Then
                    outCounter.Val += 1
                Else
                    outCounter.Val -= 1
                End If

                outFloat2.Val = total

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
                '                FBErrors.Val += sError
                '                If ClearError.Val = False Then
                '         ClearError.Val = True
                RaiseMessage(sError)

                'End If

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
        Private Sub RaiseMessage(
                ByVal Message As String,
                Optional ByVal Group As urtMSGGROUP = urtMSGGROUP.msgERROR,
                Optional ByVal Priority As urtMSGPRIORITY = urtMSGPRIORITY.msgHI,
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
        Public Function SetupVersion3(
                            ByVal VersionString As String,
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
        Public Function URTArray3(
                   ByVal sName As String,
                   ByVal sDesc As String,
                   ByVal myType As System.Guid,
                   ByVal iSize As Integer,
                   ByVal bInit As Boolean,
                   ByVal InitVal As Object,
                   ByVal myCmpPtr As IUrtTreeMember,
                    Optional ByVal EnumType As String = "",
                   Optional ByVal myOption As Integer = conOPTIONS.doDEFAULTS) _
                       As IUrtArray
            Dim myArray As IUrtArray = Nothing
            Dim ii As Integer
            Try
                myArray = CType(CUrtFBBase.Setup(sName, sDesc, myCmpPtr, myType, iSize), IURTArray)
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
        Public Function URTArray3(
                                    ByVal sName As String,
                                    ByVal sDesc As String,
                                    ByVal myType As System.Guid,
                                    ByVal iIndex As Integer,
                                    ByVal myTree As IUrtTreeMember,
                                    ByVal iSize As Integer,
                                    ByVal bInit As Boolean,
                                    ByVal InitVal As Object,
                                    Optional ByVal EnumType As String = "",
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
        Public Function URTScale3(Of T)(
                                    ByVal sName As String,
                                    ByVal sDesc As String,
                                    ByVal myType As System.Guid,
                                    ByVal bInit As Boolean,
                                    ByVal InitVal As Object,
                                    ByVal myCmpPtr As IUrtTreeMember,
                                    Optional ByVal EnumType As String = "",
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
        Public Function URTScale3(Of T)(
                                    ByVal sName As String,
                                    ByVal sDesc As String,
                                    ByVal myType As System.Guid,
                                    ByVal iIndex As Integer,
                                    ByVal myTree As IUrtTreeMember,
                                    ByVal bInit As Boolean,
                                    ByVal InitVal As Object,
                                    Optional ByVal EnumType As String = "",
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
        Public Sub URTSettings3(ByVal myData As IUrtData,
                              ByVal myOption As Integer,
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



