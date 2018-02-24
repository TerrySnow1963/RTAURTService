Imports System
Imports RTAURTService
Imports UrtTlbLib

Public Interface RTAURTAction
    Sub Execute(ByVal vbfb As URTVBFunctionBlock, ByRef commands As RTAURTActionParameters)

End Interface

Public Class RaiseMessage
    Public Shared Sub Raise(ByVal vbfb As URTVBFunctionBlock, ByVal msgString As String)
        Dim msg As New ConMessageClass
        msg.AckRequired = False
        msg.Priority = urtMSGPRIORITY.msgHI
        msg.Group = urtMSGGROUP.msgERROR
        msg.text = "Wrong type of paramter used for Action Parameter"
        vbfb.Raise(msg, 0)
    End Sub
End Class

Public Class RTAURTActionConnect
    Implements RTAURTAction

    Public Sub Execute(vbfb As URTVBFunctionBlock, ByRef params As RTAURTActionParameters) Implements RTAURTAction.Execute

        Dim init As Boolean = True

        Dim connectParams As RTAURTActionConnectParameters
        If Not params Is Nothing Then
            connectParams = TryCast(params, RTAURTActionConnectParameters)
            If Not connectParams Is Nothing Then
                init = connectParams.Init
            Else
                RaiseMessage.Raise(vbfb, "Unexpected Type for action Connect parameter")
            End If
        End If
        vbfb.Connect(init)
    End Sub
End Class


Public Class RTAURTActionExecute
    Implements RTAURTAction

    Public Sub Execute(vbfb As URTVBFunctionBlock, ByRef params As RTAURTActionParameters) Implements RTAURTAction.Execute

        Dim count As Integer = 1

        Dim executeParams As RTAURTActionExecuteParameters
        If Not params Is Nothing Then
            executeParams = TryCast(params, RTAURTActionExecuteParameters)
            If Not executeParams Is Nothing Then
                count = executeParams.Count
                If count < 1 Then
                    RaiseMessage.Raise(vbfb, "Count of executions < 1, forced to 1")
                End If
            Else

            End If
        End If

        For ii = 0 To count - 1
            vbfb.Execute(0, Nothing)
        Next
    End Sub
End Class

Public Class RTAURTActionSetValues
    Implements RTAURTAction

    Public Sub Execute(vbfb As URTVBFunctionBlock, ByRef params As RTAURTActionParameters) Implements RTAURTAction.Execute

        Dim setValueParams As RTAURTActionSetValuesParameters
        If Not params Is Nothing Then
            setValueParams = TryCast(params, RTAURTActionSetValuesParameters)
            If Not setValueParams Is Nothing Then
                For Each svItem In setValueParams.NameValueList
                    If svItem.Index <> -1 Then
                        Dim idata As IUrtData = vbfb.GetElement(svItem.Name)
                        If Not idata Is Nothing Then
                            idata.PutVariantValue(svItem.Value, "")
                        End If
                    Else
                        Dim iarray As IURTArray
                        iarray = TryCast(vbfb.GetElement(svItem.Name), IURTArray)
                        If Not iarray Is Nothing Then
                            Dim I_workingArray() As Object
                            iarray.GetArray(I_workingArray, urtBUF.dbWork)
                            If svItem.Index < 0 Or svItem.Index > I_workingArray.Count - 1 Then
                                RaiseMessage.Raise(vbfb, "SetValues for array , index out of range")
                            Else
                                I_workingArray(svItem.Index) = svItem.Value
                                iarray.PutArray(I_workingArray, "", urtBUF.dbWork)
                            End If
                        End If
                    End If

                Next

            Else

            End If
        End If
    End Sub
End Class


Public MustInherit Class RTAURTActionParameters
End Class

Public Class RTAURTActionConnectParameters
    Inherits RTAURTActionParameters
    Public Property Init As Boolean
    Private Sub New()

    End Sub
    Public Sub New(bInit As Boolean)
        Init = bInit
    End Sub
End Class

Public Class RTAURTActionExecuteParameters
    Inherits RTAURTActionParameters
    Public Property Count As Integer
    Private Sub New()
    End Sub
    Public Sub New(Optional ByVal cnt As Integer = 1)
        If cnt < 1 Then
            cnt = 1
        End If
        Count = cnt
    End Sub
End Class

Public Class SetValueData
    ReadOnly Property Name As String
    ReadOnly Property Value As Object
    ReadOnly Property Index As Integer

    Public Sub New(ByVal nam As String, ByVal val As Object, Optional ByVal idx As Integer = -1)
        Name = nam
        Value = val
        Index = idx

    End Sub
End Class

Public Class RTAURTActionSetValuesParameters
    Inherits RTAURTActionParameters

    Public _list As List(Of SetValueData)
    Public Sub New()
        _list = New List(Of SetValueData)
    End Sub

    Public Sub AddValue(ByVal Name As String, ByVal val As Object, Optional ByVal idx As Integer = -1)
        _list.Add(New SetValueData(Name, val, idx))
    End Sub

    Public ReadOnly Property NameValueList As List(Of SetValueData)
        Get
            Return _list
        End Get
    End Property


End Class

