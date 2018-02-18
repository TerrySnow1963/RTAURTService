Imports RTAURTService
Imports UrtTlbLib

Public Class URTVBQALFillTankMarried1
    Inherits RTAURTService.URTVBFunctionBlock

    Private Shared _theScript As URT.VBScript = New URT.VBScript

    Public Sub New()
        MyBase.New(_theScript)
    End Sub

    Public Overrides Sub Connect(ByVal WithInit As Boolean)
        _theScript.ConnectDataItems(WithInit)
    End Sub

    Public Overrides Sub Execute(ByVal iCause As Integer, ByVal pITMScheduler As IUrtTreeMember)
        _theScript.Execute(iCause, pITMScheduler)
    End Sub

End Class
