Public Class PAEnt
  Public Loadorder As Integer  'Load order
  Dim sid As String

  Public mContainer As Object
  Dim sLastUpdate As Date

  Public Property ID() As String
    Get
      ID = sid
    End Get
    Set(ByVal value As String)
      sid = value
    End Set
  End Property

  Public Property Lastupdate() As Date
    Get
      Lastupdate = sLastUpdate
    End Get
    Set(ByVal value As Date)
      sLastUpdate = value
    End Set
  End Property

  Public Overridable Function ChildEntityString(ByVal ent As String) As String
    ChildEntityString = ""
  End Function

  Public Overridable Sub ChildEntityString(ByVal ent As String, ByVal strEnt As String)
  End Sub

  Public Overridable Function ChildEntities(ByVal sEnt As String) As Object
    ChildEntities = Nothing
  End Function

  Public Overridable Function AvailableChildEntities(ByRef sEnt As Object) As Object
    AvailableChildEntities = Nothing
  End Function

  Public Overridable Function ParentEntities(ByRef objPar As Object) As Object
    ParentEntities = Nothing
  End Function

  Public Overridable Function AvilableParentEntities(ByRef objPar As Object) As Object
    AvilableParentEntities = Nothing
  End Function

End Class
