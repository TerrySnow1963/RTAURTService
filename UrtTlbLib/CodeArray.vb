Imports RTAInterfaces

Public MustInherit Class con_array(Of T1, T2)
    Inherits con_data
    Implements IUrtData
    Implements IURTArray

    Protected _data(-1) As T1

    Protected MustOverride Function ConvertVariant(o As Object) As T1

    Public Sub New()

    End Sub

    'Public Sub New(ByVal nSize As Integer)
    '    _data = New Array(Of T)(nSize)
    'End Sub

#Region "IURTData"

    Overrides ReadOnly Property Size() As Integer Implements IUrtData.Size
        Get
            Return _data.Length
        End Get
    End Property

    Public Overrides ReadOnly Property Size(whichBuf As urtBUF) As Integer Implements IUrtData.Size
        Get
            Return _data.Length
        End Get
    End Property

    Public Overrides Property Item(index As Integer) As Object
        Get
            Return _data(index)
        End Get
        Set(value As Object)
            _data(index) = ConvertVariant(value)
        End Set
    End Property
#End Region

#Region "IURTArray"

    Public Sub Resize(nSize As Integer, Optional ByVal eBuf As urtBUF = 1) Implements IURTArray.Resize
        Array.Resize(_data, nSize)
    End Sub

    'Public Sub PutArray(o() As Object, ByRef errStr As String, Optional ByVal eBuf As urtBUF = 1) Implements IURTArray.PutArray
    '    For ii = 0 To o.Length - 1
    '        _data(ii) = ConvertVariant(o(ii))
    '    Next
    'End Sub

    Public Sub PutArray(o As Object, ByRef errStr As String, Optional ByVal eBuf As urtBUF = 1) Implements IURTArray.PutArray
        For ii = 0 To o.Length - 1
            _data(ii) = ConvertVariant(o(ii))
        Next
    End Sub

    Public Sub GetArray(ByRef o As Object, Optional ByVal eBuf As urtBUF = 1) Implements IURTArray.GetArray
        Dim result() As T1
        Array.Resize(result, _data.Length)
        For ii = 0 To _data.Length - 1
            result(ii) = _data(ii)
        Next
        o = result
    End Sub


#End Region

    Public Overrides ReadOnly Property Guid As Guid
        Get
            Return GetType(T2).GUID
        End Get
    End Property


End Class

Public Class ConArrayBool
    Inherits con_array(Of Boolean, ConArrayBoolClass)

    Protected Overrides Sub PutVariantValueInternal(o As Object, str As String)
        Throw New NotImplementedException()
    End Sub

    Protected Overrides Function ConvertVariant(o As Object) As Boolean
        Return CBool(o)
    End Function
End Class

Public Class ConArrayBoolClass
End Class

Public Class ConArrayFloat
    Inherits con_array(Of Single, ConArrayFloatClass)

    Protected Overrides Sub PutVariantValueInternal(o As Object, str As String)
        Throw New NotImplementedException()
    End Sub

    Protected Overrides Function ConvertVariant(o As Object) As Single
        Return CSng(o)
    End Function
End Class

Public Class ConArrayFloatClass
End Class

Public Class ConArrayDouble
    Inherits con_array(Of Double, ConArrayDoubleClass)

    Protected Overrides Sub PutVariantValueInternal(o As Object, str As String)
        Throw New NotImplementedException()
    End Sub

    Protected Overrides Function ConvertVariant(o As Object) As Double
        Return CDbl(o)
    End Function
End Class

Public Class ConArrayDoubleClass
End Class

Public Class ConArrayInt
    Inherits con_array(Of Integer, ConArrayIntClass)

    Protected Overrides Sub PutVariantValueInternal(o As Object, str As String)
        Throw New NotImplementedException()
    End Sub

    Protected Overrides Function ConvertVariant(o As Object) As Integer
        Return CInt(o)
    End Function
End Class

Public Class ConArrayIntClass
End Class

Public Class ConArrayString
    Inherits con_array(Of String, ConArrayStringClass)

    Protected Overrides Sub PutVariantValueInternal(o As Object, str As String)
        Throw New NotImplementedException()
    End Sub

    Protected Overrides Function ConvertVariant(o As Object) As String
        Return o.ToString
    End Function
End Class

Public Class ConArrayStringClass
End Class

'ToDo , in urt, if the ConEnum has not been connected to the platform, then the urt throws an exception if a function is called on the IurtEnum Interface
Public Class ConArrayEnum
    Inherits con_array(Of Integer, ConArrayEnumClass)
    Implements IUrtEnum

    Private _theEnum As tSDENUM

    Protected Overrides Sub PutVariantValueInternal(o As Object, str As String)
        Throw New NotImplementedException()
    End Sub
    Public Property EnumType As String Implements IUrtEnum.EnumType
        Get
            Return _theEnum.EnumType
        End Get
        Set(value As String)
            _theEnum.EnumType = value
        End Set
    End Property

    Protected Overrides Function ConvertVariant(o As Object) As Integer
        Return CInt(o)
    End Function

End Class

Public Class ConArrayEnumClass
End Class

Public Class ConArrayTime
    Inherits con_array(Of DateTime, ConArrayTimeClass)

    Protected Overrides Sub PutVariantValueInternal(o As Object, str As String)
        Throw New NotImplementedException()
    End Sub

    Protected Overrides Function ConvertVariant(o As Object) As DateTime
        Return CDate(o)
    End Function
End Class

Public Class ConArrayTimeClass
End Class

Public Class ConArrayShort
    Inherits con_array(Of Short, ConArrayShortClass)

    Protected Overrides Sub PutVariantValueInternal(o As Object, str As String)
        Throw New NotImplementedException()
    End Sub

    Protected Overrides Function ConvertVariant(o As Object) As Short
        Return CShort(o)
    End Function
End Class

Public Class ConArrayShortClass
End Class
