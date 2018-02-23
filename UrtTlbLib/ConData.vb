Imports UrtTlbLib
Imports RTAInterfaces

Public MustInherit Class con_data
    Implements IUrtData
    Implements IRTAUrtData

    Private _name As String
    Private _description As String
    Private _options As Integer
    Public Property Description As String Implements IRTAUrtData.Description
        Get
            Return _description
        End Get
        Set(value As String)
            _description = value
        End Set
    End Property

    'Default Public Property Index(x As Integer) As Integer Implements IUrtTreeMember.Index
    '    Get
    '        Throw New NotImplementedException()
    '    End Get
    '    Set(value As Integer)
    '        Throw New NotImplementedException()
    '    End Set
    'End Property

    Public Property Name As String Implements IRTAUrtData.Name
        Get
            Return _name
        End Get
        Set(value As String)
            _name = value
        End Set
    End Property

    Public MustOverride Property Item(index As Integer) As Object Implements IRTAUrtData.Item
    Public MustOverride ReadOnly Property Guid As Guid Implements IRTAUrtData.Guid

    Public Sub PutOptions(WhichOptions As Integer, SetOptions As Integer, ByRef str As String) Implements IUrtData.PutOptions
        'ToDo align with Honeywell, only the Which Option
        _options = SetOptions
    End Sub

    Public Sub PutSecurityOptions(val1 As Integer, val2 As Integer, val3 As String) Implements IUrtData.PutSecurityOptions
        Throw New NotImplementedException()
    End Sub

    Public Sub PutVariantValue(o As Object, str As String) Implements IUrtData.PutVariantValue
        PutVariantValueInternal(o, str)
    End Sub

    Public MustOverride ReadOnly Property Size() As Integer Implements IUrtData.Size

    Public MustOverride ReadOnly Property Size(whichBuf As urtBUF) As Integer Implements IUrtData.Size

    Protected MustOverride Sub PutVariantValueInternal(o As Object, str As String)


End Class

Public MustInherit Class con_scalar(Of T1, T2)
    Inherits con_data

    Protected _val As T1
    Protected Overrides Sub PutVariantValueInternal(o As Object, str As String)
        _val = ConvertVariant(o)
    End Sub

    Protected MustOverride Function ConvertVariant(o As Object) As T1

    Public Overrides ReadOnly Property Size() As Integer
        Get
            Return 1
        End Get
    End Property

    Public Overrides ReadOnly Property Size(whichBuf As urtBUF) As Integer
        Get
            Return 1
        End Get
    End Property


    Public Overrides Property Item(index As Integer) As Object
        Get
            Return _val
        End Get
        Set(value As Object)
            _val = ConvertVariant(value)
        End Set
    End Property

    Public MustOverride Property Val As T1

    Public Overrides ReadOnly Property Guid As Guid
        Get
            Return GetType(T2).GUID
        End Get
    End Property

End Class

Public Class ConString
    Inherits con_scalar(Of String, ConStringClass)

    Public Sub New()
        _val = String.Empty
    End Sub

    Public Overrides Property Val As String
        Get
            Return _val
        End Get
        Set(value As String)
            _val = value
        End Set
    End Property

    Protected Overrides Function ConvertVariant(o As Object) As String
        Return o.ToString
    End Function
End Class

Public Class ConStringClass
End Class

Public Class ConInt
    Inherits con_scalar(Of Integer, ConIntClass)

    Public Overrides Property Val As Integer
        Get
            Return _val
        End Get
        Set(value As Integer)
            _val = value
        End Set
    End Property

    Protected Overrides Function ConvertVariant(o As Object) As Integer
        Return CInt(o)
    End Function
End Class

Public Class ConIntClass
End Class


Public Class ConBool
    Inherits con_scalar(Of Boolean, ConBoolClass)
    Public Overrides Property Val As Boolean
        Get
            Return _val
        End Get
        Set(value As Boolean)
            _val = value
        End Set
    End Property

    Protected Overrides Function ConvertVariant(o As Object) As Boolean
        Return CBool(o)
    End Function
End Class

Public Class ConBoolClass
End Class

Public Structure tSDENUM
    Public EnumType As String
    Public EnumVal As Integer

    Public Sub New(ByVal i As Integer)
        EnumVal = i
        EnumType = "ByCast"
    End Sub
    Public Shared Widening Operator CType(ByVal d As tSDENUM) As Integer
        Return d.EnumVal
    End Operator
    Public Shared Narrowing Operator CType(ByVal b As Integer) As tSDENUM
        Return New tSDENUM(b)
    End Operator

End Structure

Public Class ConEnum
    Inherits con_scalar(Of Integer, ConEnumClass)
    Implements IUrtEnum

    Private _data As tSDENUM

    Public Overrides Property Val As Integer
        Get
            Return _val
        End Get
        Set(value As Integer)
            _data.EnumVal = value
            _val = value
        End Set
    End Property

    Public Property EnumType As String Implements IUrtEnum.EnumType
        Get
            Return _data.EnumType
        End Get
        Set(value As String)
            _data.EnumType = value
        End Set
    End Property

    Protected Overrides Function ConvertVariant(o As Object) As Integer
        Return CInt(o)
    End Function
End Class

Public Class ConEnumClass
End Class

Public Class ConFloat
    Inherits con_scalar(Of Single, ConFloatClass)

    Public Overrides Property Val As Single
        Get
            Return _val
        End Get
        Set(value As Single)
            _val = value
        End Set
    End Property

    Protected Overrides Function ConvertVariant(o As Object) As Single
        Return CSng(o)
    End Function
End Class

Public Class ConFloatClass

End Class

Public Class ConDouble
    Inherits con_scalar(Of Double, ConDoubleClass)

    Public Overrides Property Val As Double
        Get
            Return _val
        End Get
        Set(value As Double)
            _val = value
        End Set
    End Property

    Protected Overrides Function ConvertVariant(o As Object) As Double
        Return CDbl(o)
    End Function
End Class

Public Class ConDoubleClass
End Class

Public Class ConTime
    Inherits con_scalar(Of DateTime, ConTimeClass)

    Public Overrides Property Val As DateTime
        Get
            Return _val
        End Get
        Set(value As DateTime)
            _val = value
        End Set
    End Property

    Protected Overrides Function ConvertVariant(o As Object) As DateTime
        Return CDate(o)
    End Function
End Class

Public Class ConTimeClass
End Class

Public Class ConShort
    Inherits con_scalar(Of Short, ConShortClass)

    Public Overrides Property Val As Short
        Get
            Return _val
        End Get
        Set(value As Short)
            _val = value
        End Set
    End Property

    Protected Overrides Function ConvertVariant(o As Object) As Short
        Return CShort(o)
    End Function
End Class

Public Class ConShortClass
End Class


