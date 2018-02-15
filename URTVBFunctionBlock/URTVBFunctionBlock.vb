Imports UrtTlbLib

Public Class URTVBFunctionBlock
    Implements IUrtTreeMember
    Implements IUrtMemberSupport

    Private _childElements As List(Of ConScalarBase)
    Private _theScript As URT.VBScript

    Public Sub New()
        _theScript = New URT.VBScript

        Dim base As CVBScriptBase = _theScript

        base.CmpPtr = Me
        _childElements = New List(Of ConScalarBase)
    End Sub

    Default Public Property Index(x As Integer) As Integer Implements IUrtTreeMember.Index
        Get
            Throw New NotImplementedException()
        End Get
        Set(value As Integer)
            Throw New NotImplementedException()
        End Set
    End Property

    Public Function GetElement(v As String) As IUrtData
        For Each e In _childElements
            If CType(e, IUrtData).Name = v Then Return e
        Next
        Return Nothing
    End Function

    Public Sub Connect(ByVal WithInit As Boolean)
        _theScript.ConnectDataItems(WithInit)
    End Sub

    Public Sub AddChild(child As IUrtTreeMember) Implements IUrtTreeMember.AddChild
        For Each e In _childElements
            If CType(e, IUrtData).Name = CType(child, IUrtData).Name Then
                Throw New Exception("Element already exists")
            End If
        Next
        _childElements.Add(child)

    End Sub

    Public Sub Raise(conMsg As ConMessageClass, cookie As Integer) Implements IUrtMemberSupport.Raise
        Trace.WriteLine(conMsg.text)
    End Sub
End Class
