Option Strict On

Imports UrtTlbLib

Public MustInherit Class URTVBFunctionBlock
    Implements IUrtTreeMember
    Implements IUrtMemberSupport

    Private _childElements As List(Of IUrtData)
    Private _theScript As CVBScriptBase

    Private Sub New()
        '_theScript = New URT.VBScript

        'Dim base As CVBScriptBase = _theScript

        'base.CmpPtr = Me
        '_childElements = New List(Of ConScalarBase)
    End Sub

    Public Sub New(ByVal script As CVBScriptBase)
        _theScript = script

        Dim base As CVBScriptBase = _theScript

        base.CmpPtr = Me
        _childElements = New List(Of IUrtData)
    End Sub

    'Default Public Property Index(x As Integer) As Integer Implements IUrtTreeMember.Index
    '    Get
    '        Throw New NotImplementedException()
    '    End Get
    '    Set(value As Integer)
    '        Throw New NotImplementedException()
    '    End Set
    'End Property

    'Public Function GetElement(v As String) As IUrtData
    '    For Each e In _childElements
    '        If CType(e, IUrtData).Name = v Then Return e
    '    Next
    '    Return Nothing
    'End Function

    Public MustOverride Sub Connect(ByVal WitInit As Boolean)

    Public MustOverride Sub Execute(ByVal iCause As Integer, ByVal pITMScheduler As IUrtTreeMember)

    Public Function GetOrCreateChildElement(ByVal name As String,
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

    Private Function AddChild(name As String, description As String, myType As Guid, iSize As Integer) As IUrtData

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
                base = New ConArrayBoolClass(iSize)
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
            End If
        Next
        ppIReq = Nothing
    End Sub
End Class
