Imports RTAInterfaces

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

Public Class ConMessageClass
    Public text As String
    Public Group As urtMSGGROUP
    Public Priority As urtMSGPRIORITY
    Public AckRequired As Boolean

    Public type As Integer

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


