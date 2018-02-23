Public Interface IRTAUrtData
    Property Name As String
    Property Description As String
    Property Item(ByVal index As Integer) As Object
    ReadOnly Property Guid As Guid
End Interface

Public Interface IRTAUrtTreeMember
    Function GetOrCreateChildElement(ByVal name As String,
                          ByVal description As String,
                          ByVal myType As System.Guid,
                          ByVal iSize As Integer) As Object
End Interface
