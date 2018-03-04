Option Explicit On
Option Strict On

Imports UrtTlbLib
Imports RTAInterfaces

Public MustInherit Class CUrtFBBase
    Implements IUrtTreeMember
    Implements IRTAUrtTreeMember
    Implements IUrtMemberSupport

    'ToDo Dont know why I used shared for the list, think mistake
    'Private Shared _childElements As List(Of IUrtData) = New List(Of IUrtData)
    Private _childElements As List(Of IUrtData) = New List(Of IUrtData)

    Protected _CmpPtr As IUrtTreeMember

    Private _logger As IRTAUrtMessageLog
    Public Property Logger As IRTAUrtMessageLog
        Get
            Return _logger
        End Get
        Set(value As IRTAUrtMessageLog)
            If Not value Is Nothing Then
                _logger = value
            End If
        End Set
    End Property

    Private _history As IRTAUrtHistoryLog
    Public Property HistoryLog As IRTAUrtHistoryLog
        Get
            Return _history
        End Get
        Set(value As IRTAUrtHistoryLog)
            'ToDo, think there is a problem when this property is written to multiple times
            If Not value Is Nothing Then
                _history = value
            End If
        End Set
    End Property

    Public Sub New()
        _CmpPtr = Me
        _childElements = New List(Of IUrtData)
        _logger = New RTAURTNullLogger
        _history = New RTAUrtNullHistory
    End Sub


    Public Property CmpPtr() As IUrtTreeMember
        Get
            Return _CmpPtr
        End Get
        Set(value As IUrtTreeMember)
            _CmpPtr = value
        End Set
    End Property

    Public Shared Function Setup(
                                ByVal name As String,
                                ByVal description As String,
                                ByVal cmpPtr As IUrtTreeMember,
                                ByVal myType As System.Guid,
                                ByVal iSize As Integer) As Object

        Return CType(cmpPtr, IRTAUrtTreeMember).GetOrCreateChildElement(name, description, myType, iSize)
    End Function

    Public Shared [Lib] As SomeClass

    Public Function GetOrCreateChildElement(ByVal name As String,
                            ByVal description As String,
                            ByVal myType As System.Guid,
                            ByVal iSize As Integer) As Object Implements IRTAUrtTreeMember.GetOrCreateChildElement
        For Each e In _childElements
            If CType(e, IRTAUrtData).Name = name Then
                Return e
            End If
        Next
        Return AddChild(name, description, myType, iSize)
    End Function

    Private Function AddChild(name As String, description As String, myType As Guid, iSize As Integer) As Object

        Dim base As Object

        Select Case myType
            Case GetType(ConBoolClass).GUID
                base = New ConBool
            Case GetType(ConDoubleClass).GUID
                base = New ConDouble
            Case GetType(ConEnumClass).GUID
                base = New ConEnum
            Case GetType(ConFloatClass).GUID
                base = New ConFloat
            Case GetType(ConIntClass).GUID
                base = New ConInt
            Case GetType(ConShortClass).GUID
                base = New ConShort
            Case GetType(ConStringClass).GUID
                base = New ConString
            Case GetType(ConTimeClass).GUID
                base = New ConTime
            Case GetType(ConArrayBoolClass).GUID
                base = New ConArrayBool()
                CType(base, IURTArray).Resize(iSize)
            Case GetType(ConArrayDoubleClass).GUID
                base = New ConArrayDouble()
                CType(base, IURTArray).Resize(iSize)
            Case GetType(ConArrayFloatClass).GUID
                base = New ConArrayFloat()
                CType(base, IURTArray).Resize(iSize)
            Case GetType(ConArrayIntClass).GUID
                base = New ConArrayInt()
                CType(base, IURTArray).Resize(iSize)
            Case GetType(ConArrayStringClass).GUID
                base = New ConArrayString()
                CType(base, IURTArray).Resize(iSize)
            Case GetType(ConArrayTimeClass).GUID
                base = New ConArrayTime()
                CType(base, IURTArray).Resize(iSize)
            Case GetType(ConArrayShortClass).GUID
                base = New ConArrayShort()
                CType(base, IURTArray).Resize(iSize)
            Case Else
                Throw New Exception(String.Format("Error when trying to connect <{0}> : Unhandled GUID"))
        End Select

        CType(base, IRTAUrtData).Name = name
        CType(base, IRTAUrtData).Description = description

        _childElements.Add(CType(base, IUrtData))
        Return base
    End Function

    Public Sub Raise(conMsg As ConMessageClass, cookie As Integer) Implements IUrtMemberSupport.Raise
        _logger.Write(conMsg.text)
    End Sub

    Public Sub Find(name As String, ByRef iid As Guid, ByRef ppIReq As Object) Implements IUrtTreeMember.Find
        ppIReq = Nothing
        For Each e In _childElements
            If CType(e, IRTAUrtData).Name = name Then
                ppIReq = e
                iid = CType(e, IRTAUrtData).Guid
                Exit For
            End If
        Next
    End Sub

    Public Function GetElement(v As String) As IUrtData
        If String.IsNullOrEmpty(v) Then Return Nothing
        For Each e In _childElements
            If CType(e, IRTAUrtData).Name = v Then Return e
        Next
        Return Nothing
    End Function

    Public Sub AddElement(name As String, iid As Guid, ByRef ppIReq As Object) Implements IUrtTreeMember.AddElement
        Throw New NotImplementedException()
    End Sub

    Public Function GetElements() As IEnumerable(Of IRTAUrtData) Implements IRTAUrtTreeMember.GetElements
        Dim list As List(Of IRTAUrtData) = New List(Of IRTAUrtData)

        For Each item In _childElements
            list.Add(CType(item, IRTAUrtData))
        Next

        Return list.AsEnumerable
    End Function
End Class


Public MustInherit Class CVBScriptBase
    Inherits CUrtFBBase

    Public MustOverride Sub Execute(ByVal iCause As Integer, ByVal pITMScheduler As IUrtTreeMember)
    Public Overridable Sub OnPreDestruct()
    End Sub

    Public Sub New()
    End Sub
End Class


