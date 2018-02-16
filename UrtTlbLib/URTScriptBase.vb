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

Public Class URTBuf
    Public Shared dbWORK As Byte()
End Class

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


Public MustInherit Class ConScalarBase
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

    Default Public Property Index(x As Integer) As Integer Implements IUrtTreeMember.Index
        Get
            Throw New NotImplementedException()
        End Get
        Set(value As Integer)
            Throw New NotImplementedException()
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

    Public Sub AddChild(child As IUrtTreeMember) Implements IUrtTreeMember.AddChild
        Throw New NotImplementedException()
    End Sub

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

    Public Function Size(buf() As Byte) As Integer Implements IUrtData.Size
        Throw New NotImplementedException()
    End Function

    Protected MustOverride Sub PutVariantValueInternal(o As Object, str As String)
End Class


Public Class ConString
    Inherits ConScalarBase
    Public val As String

    Protected Overrides Sub PutVariantValueInternal(o As Object, str As String)
        val = o.ToString
    End Sub
End Class

Public Class ConStringClass

End Class

Public Class ConInt
    Inherits ConScalarBase
    Public Val As Integer
    Protected Overrides Sub PutVariantValueInternal(o As Object, str As String)
        Val = CInt(o)
    End Sub
End Class

Public Class ConIntClass

End Class

Public Class ConBool
    Inherits ConScalarBase
    Public Val As Boolean

    Protected Overrides Sub PutVariantValueInternal(o As Object, str As String)
        Val = CBool(o)
    End Sub
End Class

Public Class ConBoolClass
End Class

Public Class ConArrayBoolClass
End Class

Public Class ConEnum
    Inherits ConScalarBase
    Implements IUrtEnum

    Public Val As Integer

    Private _EnumType As String

    Public Property EnumType As String Implements IUrtEnum.EnumType
        Get
            Return _EnumType
        End Get
        Set(value As String)
            _EnumType = value
        End Set
    End Property

    Protected Overrides Sub PutVariantValueInternal(o As Object, str As String)
        Val = CInt(o)
    End Sub
End Class

Public Class ConEnumClass
End Class

Public Class ConFloat
    Inherits ConScalarBase

    Public Val As Single
    Protected Overrides Sub PutVariantValueInternal(o As Object, str As String)
        Val = CSng(o)
    End Sub
End Class

Public Class ConFloatClass

End Class

Public Class ConDouble
    Inherits ConScalarBase

    Public Val As Double
    Protected Overrides Sub PutVariantValueInternal(o As Object, str As String)
        Val = CDbl(o)
    End Sub
End Class

Public Class ConDoubleClass

End Class

Public Interface IURTArray
    Sub Resize(ByVal nSize As Integer, ByRef buffer As Byte())
    Sub GetArray(ByVal o As Object, ByRef buffer As Byte())
    Sub PutArray(ByVal o() As Object, ByRef str As String)

End Interface

Public Interface IUrtTreeMember
    Default Property Index(ByVal x As Integer) As Integer
    Sub AddChild(ByVal child As IUrtTreeMember)

End Interface

Public Interface IUrtData
    Sub PutSecurityOptions(ByVal val1 As Integer, ByVal val2 As Integer, ByVal val3 As String)
    Function Size() As Integer
    Function Size(buf As Byte()) As Integer
    Sub PutOptions(ByVal opt1 As Integer, ByVal opt2 As Integer, ByRef str As String)
    Sub PutVariantValue(ByVal o As Object, ByVal str As String)
    Property Name As String
    Property Description As String
End Interface

Public Interface IUrtMemberSupport
    Sub Raise(ByVal conMsg As ConMessageClass, ByVal cookie As Integer)

End Interface

Public Class URTBuffer
    Public dbWork As Integer
End Class

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
                                ByVal iSize As Integer) As ConScalarBase

        Dim base As ConScalarBase = Nothing
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
            Case Else
                Throw New Exception(String.Format("Error when trying to connect <{0}> : Unhandled GUID"))
        End Select

        base.Name = name
        base.Description = description
        cmpPtr.AddChild(base)

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


