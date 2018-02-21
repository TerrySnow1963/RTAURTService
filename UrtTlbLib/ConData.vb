Imports UrtTlbLib

Public MustInherit Class con_data
    Implements IUrtData

    Private _name As String
    Private _description As String
    Private _options As Integer
    Public Property Description As String Implements IUrtData.Description
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

    Public Property Name As String Implements IUrtData.Name
        Get
            Return _name
        End Get
        Set(value As String)
            _name = value
        End Set
    End Property

    Public MustOverride Property Item(index As Integer) As Object Implements IUrtData.Item

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

    Public MustOverride Function Size() As Integer Implements IUrtData.Size

    Public MustOverride Function Size(whichBuf As urtBUF) As Integer Implements IUrtData.Size

    Protected MustOverride Sub PutVariantValueInternal(o As Object, str As String)


End Class

Public MustInherit Class con_scalar(Of T)
    Inherits con_data

    Protected _val As T
    Protected Overrides Sub PutVariantValueInternal(o As Object, str As String)
        _val = ConvertVariant(o)
    End Sub

    Protected MustOverride Function ConvertVariant(o As Object) As T

    Public Overrides Function Size() As Integer
        Return 1
    End Function

    Public Overrides Function Size(whichBuf As urtBUF) As Integer
        Return 1
    End Function

    Public Overrides Property Item(index As Integer) As Object
        Get
            Return _val
        End Get
        Set(value As Object)
            _val = ConvertVariant(value)
        End Set
    End Property

End Class

Public Class ConString
    Inherits con_scalar(Of String)

    Public Sub New()
        _val = String.Empty
    End Sub

    Public Overridable Property Val As String
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
    Inherits con_scalar(Of Integer)

    Public Overridable Property Val As Integer
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
    Inherits con_scalar(Of Boolean)

    Public Overridable Property Val As Boolean
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
End Structure

Public Class ConEnum
    Inherits con_scalar(Of tSDENUM)
    Implements IUrtEnum

    'Private _data As tSDENUM

    Public Property Val As Integer
        Get
            Return _val.EnumVal
        End Get
        Set(value As Integer)
            _val.EnumVal = value
        End Set
    End Property

    Private _EnumType As String

    Public Property EnumType As String Implements IUrtEnum.EnumType
        Get
            Return _val.EnumType
        End Get
        Set(value As String)
            _val.EnumType = value
        End Set
    End Property

    Protected Overrides Function ConvertVariant(o As Object) As tSDENUM
        Dim newVal As tSDENUM = New tSDENUM
        newVal.EnumVal = CInt(o)
        Return newVal
    End Function
End Class

Public Class ConEnumClass
End Class


