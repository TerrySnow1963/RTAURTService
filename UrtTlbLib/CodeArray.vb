Imports RTAInterfaces

Public MustInherit Class con_array(Of T1, T2)
    Inherits con_data
    Implements IUrtData
    Implements IURTArray

    Protected _data As T1()

    Protected MustOverride Function ConvertVariant(o As Object) As T1

    Public Sub New()

    End Sub

    'Public Sub New(ByVal nSize As Integer)
    '    _data = New Array(Of T)(nSize)
    'End Sub

#Region "IURTData"

    Public Overrides Function Size() As Integer Implements IUrtData.Size
        Return _data.Length
    End Function

    Public Overrides Function Size(whichBuf As urtBUF) As Integer Implements IUrtData.Size
        Return _data.Length
    End Function

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

    Public Sub PutArray(o() As Object, ByRef errStr As String, Optional ByVal eBuf As urtBUF = 1) Implements IURTArray.PutArray
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
