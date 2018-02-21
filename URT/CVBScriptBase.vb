Imports UrtTlbLib

Public MustInherit Class CVBScriptBase
    Inherits CUrtFBBase
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


