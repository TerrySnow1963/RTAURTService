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

    Function GetElements() As IEnumerable(Of IRTAUrtData)
End Interface

Public Interface IRTAUrtMessageLog
    Sub Write(ByVal message As String)
End Interface