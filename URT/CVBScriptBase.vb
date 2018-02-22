Imports UrtTlbLib

Public MustInherit Class CVBScriptBase
    Inherits CUrtFBBase

    Public MustOverride Sub Execute(ByVal iCause As Integer, ByVal pITMScheduler As IUrtTreeMember)
    Public Overridable Sub OnPreDestruct()
    End Sub

    Public Sub New()
    End Sub
End Class


