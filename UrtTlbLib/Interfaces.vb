Public Interface IURTArray
    Sub Resize(ByVal nSize As Integer, Optional ByVal eBuf As urtBUF = 1)
    Sub GetArray(ByRef o As Object, Optional ByVal eBuf As urtBUF = 1)
    Sub PutArray(ByVal o As Object, ByRef errStr As String, Optional ByVal eBuf As urtBUF = 1)

End Interface

Public Interface IUrtTreeMember
    Sub Find(ByVal name As String, ByRef iid As Guid, ByRef ppIReq As Object)
    Sub AddElement(name As String, ByVal iid As Guid, ByRef ppIReq As Object)
End Interface

Public Interface IUrtData
    Sub PutSecurityOptions(ByVal val1 As Integer, ByVal val2 As Integer, ByVal val3 As String)
    ReadOnly Property Size() As Integer
    ReadOnly Property Size(whichBuf As urtBUF) As Integer
    Sub PutOptions(ByVal opt1 As Integer, ByVal opt2 As Integer, ByRef str As String)
    Sub PutVariantValue(ByVal o As Object, ByVal str As String)
    'todo correct signatures
    'wstring PutVariantValue(Variant vValue, Long iExtStart, long iIntStart, long nElements, urtBUF eBuffer = dbWORK);
End Interface

Public Interface IUrtMemberSupport
    Sub Raise(ByVal conMsg As ConMessageClass, ByVal cookie As Integer)
End Interface

