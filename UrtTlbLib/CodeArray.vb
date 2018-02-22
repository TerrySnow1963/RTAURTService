Imports RTAInterfaces

Public MustInherit Class con_array(Of T)
    Inherits con_data
    Implements IUrtData
    Implements IURTArray

    Protected _data As T()

    Protected MustOverride Function ConvertVariant(o As Object) As T

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
        Dim result() As T
        Array.Resize(result, _data.Length)
        For ii = 0 To _data.Length - 1
            result(ii) = _data(ii)
        Next
        o = result
    End Sub


#End Region


End Class

Public Class ConArrayBool
    Inherits con_array(Of Boolean)

    Protected Overrides Sub PutVariantValueInternal(o As Object, str As String)
        Throw New NotImplementedException()
    End Sub

    Protected Overrides Function ConvertVariant(o As Object) As Boolean
        Return CBool(o)
    End Function
End Class

Public Class ConArrayBoolClass
End Class
