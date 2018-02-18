Imports UrtTlbLib

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

Public MustInherit Class CVBScriptBase
    Protected _CmpPtr As IUrtTreeMember

    Public Property CmpPtr() As IUrtTreeMember
        Get
            Return _CmpPtr
        End Get
        Set(value As IUrtTreeMember)
            _CmpPtr = value
        End Set
    End Property


    Public MustOverride Sub Execute(ByVal iCause As Integer, ByVal pITMScheduler As IUrtTreeMember)
    Public Overridable Sub OnPreDestruct()
    End Sub

    Public Sub New()

    End Sub
End Class


Public MustInherit Class con_data(Of T)
    Implements IUrtData
    Implements IUrtTreeMember

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

    Public Function GetOrCreateChildElement(ByVal name As String,
                                ByVal description As String,
                                ByVal myType As System.Guid,
                                ByVal iSize As Integer) As IUrtTreeMember Implements IUrtTreeMember.GetOrCreateChildElement
        Throw New NotImplementedException()
    End Function

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

    Public Function Size() As Integer Implements IUrtData.Size
        Throw New NotImplementedException()
    End Function

    Public Function Size(whichBuf As urtBUF) As Integer Implements IUrtData.Size
        Throw New NotImplementedException()
    End Function

    Protected MustOverride Sub PutVariantValueInternal(o As Object, str As String)
End Class

Public MustInherit Class con_array(Of T)
    Implements IUrtData
    Implements IUrtTreeMember

    Protected _data() As T

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

    'Default Public Property Index(ii As Integer) As Integer Implements IUrtTreeMember.Index
    '    Get
    '        Return _data(ii)
    '    End Get
    '    Set(value As Integer)
    '        Throw New NotImplementedException()
    '    End Set
    'End Property

    Default Public Property Item(ii As Integer) As T
        Get
            Return _data(ii)
        End Get
        Set(value As T)
            _data(ii) = value
        End Set
    End Property
    Public Property Name As String Implements IUrtData.Name
        Get
            Return _name
        End Get
        Set(value As String)
            _name = value
        End Set
    End Property

    Public Function GetOrCreateChildElement(ByVal name As String,
                                ByVal description As String,
                                ByVal myType As System.Guid,
                                ByVal iSize As Integer) As IUrtTreeMember Implements IUrtTreeMember.GetOrCreateChildElement
        Throw New NotImplementedException()
    End Function

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

    Public Function Size() As Integer Implements IUrtData.Size
        Return _data.Length
    End Function

    Protected MustOverride Sub PutVariantValueInternal(o As Object, str As String)

    Public Function Size(whichBuf As urtBUF) As Integer Implements IUrtData.Size
        Return _data.Length
    End Function
End Class

Public Class ConString
    Inherits con_data(Of String)
    Public val As String

    Protected Overrides Sub PutVariantValueInternal(o As Object, str As String)
        val = o.ToString
    End Sub
End Class

Public Class ConStringClass

End Class

Public Class ConInt
    Inherits con_data(Of Integer)
    Public Val As Integer
    Protected Overrides Sub PutVariantValueInternal(o As Object, str As String)
        Val = CInt(o)
    End Sub
End Class

Public Class ConIntClass

End Class

Public Class ConBool
    Inherits con_data(Of Boolean)
    Public Val As Boolean

    Protected Overrides Sub PutVariantValueInternal(o As Object, str As String)
        Val = CBool(o)
    End Sub
End Class

Public Class ConBoolClass
End Class

Public Class ConArrayBool
    Inherits con_array(Of Boolean)
    Implements IURTArray

    Private Sub New()

    End Sub
    Public Sub New(ByVal size As Integer)
        Array.Resize(_data, size)
    End Sub

    Protected Overrides Sub PutVariantValueInternal(o As Object, str As String)
        Throw New NotImplementedException()
    End Sub

    Public Sub Resize(nSize As Integer, whichBuf As urtBUF) Implements IURTArray.Resize
        If nSize > 0 Then
            Array.Resize(_data, nSize)
        End If
    End Sub

    Public Sub PutArray(o() As Object, ByRef errStr As String) Implements IURTArray.PutArray
        For ii = 0 To o.Length - 1
            _data(ii) = CBool(o(ii))
        Next
    End Sub

    Public Sub GetArray(ByRef o As Object, whichBuf As urtBUF) Implements IURTArray.GetArray
        Dim result() As Boolean
        Array.Resize(result, _data.Length)
        For ii = 0 To _data.Length - 1
            result(ii) = _data(ii)
        Next
        o = result
    End Sub
End Class

Public Class ConArrayBoolClass
End Class

Public Structure tSDENUM
    Public EnumType As String
    Public EnumVal As Integer
End Structure

Public Class ConEnum
    Inherits con_data(Of tSDENUM)
    Implements IUrtEnum

    Private _data As tSDENUM

    Public Property Val As Integer
        Get
            Return _data.EnumVal
        End Get
        Set(value As Integer)
            _data.EnumVal = value
        End Set
    End Property

    Private _EnumType As String

    Public Property EnumType As String Implements IUrtEnum.EnumType
        Get
            Return _data.EnumType
        End Get
        Set(value As String)
            _data.EnumType = value
        End Set
    End Property

    Protected Overrides Sub PutVariantValueInternal(o As Object, str As String)
        _data.EnumVal = CInt(o)
    End Sub
End Class

Public Class ConEnumClass
End Class

'ToDo need to change as actual implementation has ConFloat as an interface

Public Class ConFloat
    Inherits con_data(Of Single)

    Public Val As Single
    Protected Overrides Sub PutVariantValueInternal(o As Object, str As String)
        Val = CSng(o)
    End Sub
End Class

Public Class ConFloatClass

End Class

Public Class ConDouble
    Inherits con_data(Of Double)

    Public Val As Double
    Protected Overrides Sub PutVariantValueInternal(o As Object, str As String)
        Val = CDbl(o)
    End Sub
End Class

Public Class ConDoubleClass

End Class

Public Interface IURTArray
    Sub Resize(ByVal nSize As Integer, ByVal whichBuf As urtBUF)
    Sub GetArray(ByRef o As Object, ByVal whichBuf As urtBUF)
    Sub PutArray(ByVal o() As Object, ByRef errStr As String)

End Interface

Public Interface IUrtTreeMember
    'Default Property Index(ByVal x As Integer) As Integer
    Function GetOrCreateChildElement(ByVal name As String,
                                ByVal description As String,
                                ByVal myType As System.Guid,
                                ByVal iSize As Integer) As IUrtTreeMember

End Interface

Public Interface IUrtData
    Sub PutSecurityOptions(ByVal val1 As Integer, ByVal val2 As Integer, ByVal val3 As String)
    Function Size() As Integer
    Function Size(whichBuf As urtBUF) As Integer
    Sub PutOptions(ByVal opt1 As Integer, ByVal opt2 As Integer, ByRef str As String)
    Sub PutVariantValue(ByVal o As Object, ByVal str As String)
    'todo correct signatures
    'wstring PutVariantValue(Variant vValue, Long iExtStart, long iIntStart, long nElements, urtBUF eBuffer = dbWORK);
    Property Name As String
    Property Description As String
End Interface

Public Interface IUrtMemberSupport
    Sub Raise(ByVal conMsg As ConMessageClass, ByVal cookie As Integer)

End Interface

Public Class ConMessageClass
    Public text As String
    Public Group As urtMSGGROUP
    Public Priority As urtMSGPRIORITY
    Public AckRequired As Boolean

    Public type As Integer

End Class

Public Class CUrtFBBase
    Public Shared Function Setup(
                                ByVal name As String,
                                ByVal description As String,
                                ByVal cmpPtr As IUrtTreeMember,
                                ByVal myType As System.Guid,
                                ByVal iSize As Integer) As IUrtTreeMember

        Dim base As IUrtTreeMember = Nothing

        Return cmpPtr.GetOrCreateChildElement(name, description, myType, iSize)

        Return base
    End Function

    Public Shared [Lib] As SomeClass
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


