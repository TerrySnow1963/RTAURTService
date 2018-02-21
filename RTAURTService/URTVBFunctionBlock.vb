Option Strict On
Imports URT
Imports UrtTlbLib

Public MustInherit Class URTVBFunctionBlock
    Inherits CUrtFBBase
    Private _theScript As CVBScriptBase

    Private Sub New()
        '_theScript = New URT.VBScript

        'Dim base As CVBScriptBase = _theScript

        'base.CmpPtr = Me
        '_childElements = New List(Of ConScalarBase)
    End Sub

    Public Sub New(ByVal script As CVBScriptBase)
        _theScript = script
        _theScript.CmpPtr = Me

    End Sub

    'Default Public Property Index(x As Integer) As Integer Implements IUrtTreeMember.Index
    '    Get
    '        Throw New NotImplementedException()
    '    End Get
    '    Set(value As Integer)
    '        Throw New NotImplementedException()
    '    End Set
    'End Property


    Public MustOverride Sub Connect(ByVal WitInit As Boolean)

    Public MustOverride Sub Execute(ByVal iCause As Integer, ByVal pITMScheduler As IUrtTreeMember)


End Class
