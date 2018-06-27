Option Explicit On
Imports PA_Framework_Lib
Imports PA_Framework_OM.OMGlobals
Partial Public Class PAProjects
  Public Shared mObjectForJoinTable As Object

  Public Shared Function GetChildPropertyName(ByVal sPar As String, ByVal sChld As String) As String
    'should be used with care only for join tables. One parent can have the same child with different
    'propterty names
    'Dim eR As ChildRelationship
    '] Dim eD As EntityDataItem
    Dim pr As PAProject
    Dim ei As ProjectEntity
    Dim pei As ProjectEntityItem
    GetChildPropertyName = ""
    For Each pr In cPAProjects
      For Each ei In pr.ChildprojectEntities
        If ei.EntityName = sPar Then
          For Each pei In ei.mChildProjectEntityItems
            If pei.ChildName = sChld Then
              GetChildPropertyName = pei.PropertyName
              Exit Function
            End If
          Next pei
        End If
      Next ei
    Next pr
  End Function

  Public Shared Function GetAssociation(ByVal ParentEntity As String, ByVal ChldEntity As String, ByVal sChildPropertyName As String, ByRef sJoinTable As String) As String
    'Gets the Parent-child relationship given a parent and a child returns 1-M.M-M or nil
    'sParentEntity: parent entity Input
    'sChildEntity : Child Entity Input
    'sChildPropertyName: Property name of the child entity. Input
    'sJoinTable:Join Table is output
    'Returns association (1-M, M-M etc)
    GetAssociation = ""
    sJoinTable = "CSR"
    Dim pr As PAProject
    Dim pe As ProjectEntity
    Dim pei As ProjectEntityItem
    For Each pr In cPAProjects
      For Each pe In pr.ChildprojectEntities
        If pe.EntityName = ParentEntity Then
          For Each pei In pe.mChildProjectEntityItems
            If (pei.PropertyName = sChildPropertyName) Then 'And (pei.ChildName = ChldEntity)
              If pei.ChildRelationship = "M-M" Then sJoinTable = pei.RelationshipType
              GetAssociation = pei.ChildRelationship
              Exit Function
            End If
          Next pei
        End If
      Next pe
    Next pr
    GetAssociation = ""
  End Function

  Function IsJoinTable(ByVal ent As String) As Boolean
    Dim pr As PAProject
    Dim pe As ProjectEntity
    Dim pei As ProjectEntityItem
    For Each pr In cPAProjects
      For Each pe In pr.ChildprojectEntities
        For Each pei In pe.mChildProjectEntityItems
          If pei.RelationshipType = ent Then
            IsJoinTable = True
            Exit Function
          End If
        Next pei
      Next pe
    Next pr
    IsJoinTable = False
  End Function

  Public Shared Function GetJoinTable(ByVal sPar As String, ByVal sChldProperty As String) As String
    GetJoinTable = "CSR"
    If cPAProjects Is Nothing Then Exit Function 'happens before the three pa project entities are loaded. csr is ok for them.    
    Dim pr As PAProject
    Dim pe As ProjectEntity
    Dim pei As ProjectEntityItem
    For Each pr In cPAProjects
      For Each pe In pr.ChildprojectEntities
        If pe.EntityName = sPar Then
          For Each pei In pe.mChildProjectEntityItems
            If pei.PropertyName = sChldProperty Then
              If pei.ChildRelationship = "M-M" Then GetJoinTable = pei.RelationshipType
              Exit Function
            End If
          Next pei
        End If
      Next pe
    Next pr
  End Function

  Public Function GetEntityDependencies(ByVal ent As Object, ByRef strParDetails As String, ByRef strChldDetails As String) As Integer
    'Returns total parent and child dependencies
    Dim pr As PAProject
    Dim pe As ProjectEntity
    Dim pei As ProjectEntityItem

    Dim ParObj As New Object
    Dim ChldObj As New Object
    strParDetails = ""
    strChldDetails = ""
    Dim nCount As Integer = 0
    Dim nCountForEachParent As Integer = 0
    Dim nCountForEachChild As Integer = 0
    For Each pr In cPAProjects
      For Each pe In pr.ChildprojectEntities
        If pe.EntityName = TypeName(ent) Then 'Get all children
          For Each pei In pe.mChildProjectEntityItems
            Try
              ChldObj = CreateObjectFromString(pei.ChildName)
              If ChldObj Is Nothing Then Continue For
              Call UIFunctions.FillChildEntities(cPAProjects, ent, ChldObj, pei.PropertyName)
              nCountForEachChild = ent.childentities(pei.PropertyName).Count
            Catch ex As Exception
              MsgBox(ex.Message & vbCrLf & "Error from GetEntityDependencies (FillChildEntities): Parent:" & TypeName(ent) & " ChildName: " & TypeName(ChldObj))
              Call AppSettings.WriteToErrorLog("Error from GetEntityDependencies (FillChildEntities): Parent:" & TypeName(ent) & "; ChildName:" & TypeName(ChldObj))
            End Try
            strChldDetails = strChldDetails & pei.ChildName & ": " & nCountForEachChild & vbCrLf
          Next pei
        End If
        For Each pei In pe.mChildProjectEntityItems
          If pei.ChildName = TypeName(ent) Then  'Get parent dependencies
            Try
              ParObj = CreateObjectFromString(pe.EntityName)
              Call UIFunctions.FillParentEntities(cPAProjects, ParObj, ent, pei.PropertyName)
              nCountForEachParent = ent.parententities(ParObj).count
            Catch
              MsgBox("Error from GerEntityDepencies (FillParentEntities): Parent:" & TypeName(ParObj) & " ChildName: " & TypeName(ent))
              Call AppSettings.WriteToErrorLog("Error from GerEntityDepencies (FillParentEntities): Parent:" & TypeName(ParObj) & "; ChildName:" & TypeName(ent))
            End Try
            'If nCountForEachParent = 0 Then Continue For
            nCount = nCount + nCountForEachParent
            strParDetails = strParDetails & pe.EntityName & ": " & nCountForEachParent & vbCrLf
          End If
        Next pei
        'strParDetails = strParDetails & vbCrLf
      Next pe
    Next pr
    Return nCount
  End Function

  Public Function GetProjectByName(ByVal sName As String) As PAProject
    GetProjectByName = Nothing
    Dim pr As PAProject

    For Each pr In cPAProjects
      If sName.Equals(pr.ProjectName.ToString) Then
        'If String.Equals(pr.ProjectName.ToString, sName.ToString) Then
        GetProjectByName = pr
        Exit Function
      End If
    Next
  End Function
  
End Class

