Imports UrtTlbLib

Public Structure con_data1(Of T)
    Private Class CWrapper(Of T)
        Public _myVal As T
    End Class
    Private _myOtherVal As T

    Private _wrapper As CWrapper(Of T)
    Public Property Val() As T
        Get
            If _wrapper Is Nothing Then
                Return _myOtherVal
            Else
                Return _wrapper._myVal
            End If
        End Get
        Set(value As T)
            _myOtherVal = value
            If Not _wrapper Is Nothing Then
                _wrapper._myVal = value
            End If
        End Set
    End Property

    Public Function SetUp(ByVal Name As String,
                          ByVal pContextITM As IUrtTreeMember,
                          Optional ByVal iElement As Long = -1,
                          Optional ByVal Description As String = "",
                          Optional ByVal iSetOPtions As Long = 0,
                          Optional ByVal Connection As String = "") As Boolean

        Dim ob As Object = Nothing

        Dim guid As Guid
        pContextITM.Find(Name, guid, ob)

        If ob Is Nothing Then
            pContextITM.AddElement(Name, GetType(ConBoolClass).GUID, ob)
        Else
            If guid <> GetType(ConBoolClass).GUID Then
                Throw New Exception("Wrong TypeOf for existing item")
            End If
            _wrapper = New CWrapper(Of T)
            _wrapper._myVal = ob
        End If
        Return True

    End Function

End Structure

Public Class ConTypes
    Private Sub New()
    End Sub
    Public Shared BOOLTYPE As Guid = New Guid("0f8fad5b-d9cb-469f-a165-70867728950e")
    Public Shared STRINGTYPE As Guid = New Guid("0f8fad5b-d9cb-469f-a165-70867728950f")
End Class

Public Enum urtMSGGROUP
    msgERROR = 0
End Enum

Public Enum urtMSGPRIORITY
    msgHI = 0
End Enum

Public Enum conOPTIONS
    doDEFAULTS = 0
End Enum

Public Enum urtBUF
    dbWork = 0
    dbInput = 1
    dbOutput = 2
End Enum

'ToDo need to change as actual implementation has ConFloat as an interface

Public Class ConFloat
    Inherits con_scalar(Of Single)

    Public Val As Single
    'Protected Overrides Sub PutVariantValueInternal(o As Object, str As String)
    '    Val = CSng(o)
    'End Sub

    Protected Overrides Function ConvertVariant(o As Object) As Single
        Return CSng(o)
    End Function
End Class

Public Class ConFloatClass

End Class

Public Class ConDouble
    Inherits con_scalar(Of Double)

    Public Val As Double
    'Protected Overrides Sub PutVariantValueInternal(o As Object, str As String)
    '    Val = CDbl(o)
    'End Sub

    Protected Overrides Function ConvertVariant(o As Object) As Double
        Return CDbl(o)
    End Function
End Class

Public Class ConDoubleClass

End Class


Public Class ConMessageClass
    Public text As String
    Public Group As urtMSGGROUP
    Public Priority As urtMSGPRIORITY
    Public AckRequired As Boolean

    Public type As Integer

End Class

Public Class CUrtFBBase
    Implements IUrtTreeMember
    Implements IUrtMemberSupport

    Private Shared _childElements As List(Of IUrtData) = New List(Of IUrtData)

    Public Sub New()
        '_childElements = New List(Of IUrtData)
    End Sub


    Public Shared Function Setup(
                                ByVal name As String,
                                ByVal description As String,
                                ByVal cmpPtr As IUrtTreeMember,
                                ByVal myType As System.Guid,
                                ByVal iSize As Integer) As Object

        Dim base As Object = Nothing

        'Return cmpPtr.GetOrCreateChildElement(name, description, myType, iSize)
        Return GetOrCreateChildElement(name, description, myType, iSize)

        Return base
    End Function

    Public Shared [Lib] As SomeClass

    Public Shared Function GetOrCreateChildElement(ByVal name As String,
                            ByVal description As String,
                            ByVal myType As System.Guid,
                            ByVal iSize As Integer) As IUrtData
        For Each e In _childElements
            If CType(e, IUrtData).Name = name Then
                Return e
            End If
        Next
        Return AddChild(name, description, myType, iSize)
    End Function

    Private Shared Function AddChild(name As String, description As String, myType As Guid, iSize As Integer) As IUrtData

        Dim base As IUrtData

        Select Case myType
            Case GetType(ConBoolClass).GUID
                base = New ConBool
            Case GetType(ConStringClass).GUID
                base = New ConString
            Case GetType(ConIntClass).GUID
                base = New ConInt
            Case GetType(ConFloatClass).GUID
                base = New ConFloat
            Case GetType(ConEnumClass).GUID
                base = New ConEnum
            Case GetType(ConArrayBoolClass).GUID
                base = New ConArrayBool()
                CType(base, IURTArray).Resize(iSize)
            Case Else
                Throw New Exception(String.Format("Error when trying to connect <{0}> : Unhandled GUID"))
        End Select

        base.Name = name
        base.Description = description

        _childElements.Add(base)
        Return base
    End Function

    Public Sub Raise(conMsg As ConMessageClass, cookie As Integer) Implements IUrtMemberSupport.Raise
        Trace.WriteLine(conMsg.text)
    End Sub

    Public Sub Find(name As String, ByRef iid As Guid, ByRef ppIReq As Object) Implements IUrtTreeMember.Find
        For Each e In _childElements
            If CType(e, IUrtData).Name = name Then
                ppIReq = e
                'ToDo work out how to get the right guid
                iid = Guid.Empty
            End If
        Next
        ppIReq = Nothing
    End Sub

    Public Function GetElement(v As String) As IUrtData
        For Each e In _childElements
            If CType(e, IUrtData).Name = v Then Return e
        Next
        Return Nothing
    End Function

    Public Sub AddElement(name As String, iid As Guid, ByRef ppIReq As Object) Implements IUrtTreeMember.AddElement
        Throw New NotImplementedException()
    End Sub
End Class

Public Interface IUrtEnum
    Property EnumType As String
End Interface

Public Class SomeClass
    Public Sub urtGetItem(ByVal name As String, ByVal myType As System.Guid, ByVal myTree As IUrtTreeMember, ByVal index As Integer,
                          ByVal someString As String, ByVal guidITM As System.Guid, ByVal o As Object, ByVal desc As String,
                          ByVal val1 As Integer, ByVal val2 As Integer, ByVal enumStr As String,
                          ByVal flag1 As Boolean, ByVal flag2 As Boolean)

    End Sub

End Class


