'<ComClass(BLFunctions.ClassId, BLFunctions.InterfaceId, BLFunctions.EventsId)> _
Public Class BLFunctions

#Region "COM GUIDs"
  ' These  GUIDs provide the COM identity for this class 
  ' and its COM interfaces. If you change them, existing 
  ' clients will no longer be able to access the class.
  Public Const ClassId As String = "9ef03a74-7bf3-41dd-9202-4b26b5944259"
  Public Const InterfaceId As String = "b2591d0a-91d8-40e6-9def-af8b6642de33"
  Public Const EventsId As String = "7a23ca71-9e36-4454-b5a2-c5a61568aef1"
#End Region

  ' A creatable COM class must have a Public Sub New() 
  ' with no parameters, otherwise, the class will not be 
  ' registered in the COM registry and cannot be created 
  ' via CreateObject.
  Public Sub New()
    MyBase.New()
  End Sub

  Private Shared nUpdate As Integer
  Private Shared strAssoc As String

  Public Shared Function AddChildEntityViaJoin(ByRef objEDI As Object, ByVal sEntJt As String, ByRef enParent As Object, ByRef enChld As Object, ByVal sChildPropertyName As String) As Integer
    Dim sParentPropertyName As String
    sParentPropertyName = objEDI.GetChildPropertyName(sEntJt, TypeName(enParent))

    objEDI.mObjectForJoinTable.ChildEntityString(sParentPropertyName, enParent.ID)
    objEDI.mObjectForJoinTable.ChildEntityString(sChildPropertyName, enChld.ID)

    Call objEDI.mObjectForJoinTable.ChildEntities(sParentPropertyName).Add(enParent)
    Call objEDI.mObjectForJoinTable.ChildEntities(sChildPropertyName).Add(enChld)

    Call enParent.ChildEntities(sChildPropertyName).Add(enChld)

    nUpdate = objEDI.mObjectForJoinTable.mContainer.AddtoDB(objEDI.mObjectForJoinTable)
    If nUpdate = 0 Then  'revert
      Call enParent.ChildEntities(sChildPropertyName).Remove(enChld)
      MsgBox("Join Table AddtoDB Error: Parent ID:" & enParent.ID)
      'Else
      'objEDI.mObjectForJoinTable.mContainer.Add(objEDI.mObjectForJoinTable) ' add to the container
    End If

    AddChildEntityViaJoin = nUpdate
  End Function
  '
  Public Shared Function RemoveChildEntityViaJoin(ByRef objEDI As Object, ByVal sEntJt As String, ByRef enParent As Object, ByRef enChld As Object, ByVal sChildPropertyName As String) As Integer
    Dim entJt As Object
    Dim entCont As New Object
    entJt = GetItemForJoinTable(objEDI, sEntJt, enParent, enChld, sChildPropertyName)

    Call enParent.ChildEntities(sChildPropertyName).Remove(enChld)
    entCont = entJt.mContainer
    nUpdate = entCont.DeleteFromDB(entJt)
    If nUpdate = 0 Then  'revert
      Call enParent.ChildEntities(sChildPropertyName).Add(enChld)
      MsgBox("Update Error: Parent ID:" & enParent.ID)
    End If
    RemoveChildEntityViaJoin = nUpdate
  End Function

  Public Shared Function AddChildEntity(ByRef objEDI As Object, ByRef enParent As Object, ByRef enChld As Object, ByVal sChildPropertyName As String) As Integer
    'returns 0 if successful, 1 if the record is updated since last read, 2 if the entity attempted to add has already been added (1-M)
    'Adds a child entity ehchld to entity enParent
    Dim sJoinTable As String = String.Empty
    Dim strAssoc As String
    Dim bAddEntity As Boolean
    bAddEntity = True

    strAssoc = objEDI.GetAssociation(TypeName(enParent), TypeName(enChld), sChildPropertyName, sJoinTable)
    If strAssoc = "1-M" Then    ' check to see if the Child entity is still available.
      bAddEntity = Not UIFunctions.CheckAllParentsForChildPresence(enParent.mContainer, enChld, sChildPropertyName)
    End If
    If strAssoc = "M-1" And enParent.ChildEntities(sChildPropertyName).Count <> 0 Then      ' only one child allowed for M-1.
      AddChildEntity = 3
      Exit Function
    End If
    If bAddEntity = False Then
      MsgBox("The entity you are attempting to Add is not available any more.It has been added to another entity", vbOKOnly)
      AddChildEntity = 2  'entity is no more available because it has been added to another entity (1-M)
      Exit Function
    End If

    If strAssoc = "M-M" And sJoinTable <> "CSR" Then  'JTFK's
      objEDI.mObjectForJoinTable = objEDI.CreateObjectFromString(sJoinTable)
      If sJoinTable = String.Empty Then
        MsgBox("CreateObjectFromString() does not have the case for entity " & sJoinTable)
        AddChildEntity = 0
        Exit Function
      End If
      AddChildEntity = AddChildEntityViaJoin(objEDI, sJoinTable, enParent, enChld, sChildPropertyName)
      Exit Function
    End If
    ' start of CSFK's
    Call enParent.ChildEntities(sChildPropertyName).Add(enChld)
    Call BuildChildEntityString(enParent, sChildPropertyName)
    nUpdate = enParent.mContainer.UpdateDB(enParent)
    If nUpdate = 0 Then  'revert the child addition
      Call enParent.ChildEntities(enChld).Remove(enChld)

      Call BuildChildEntityString(enParent, enParent.ChildEntities(enChld))
      MsgBox("This record is not present or has been updated since you last refreshed. Please Load Data again and then update.")
      AddChildEntity = 0
      Exit Function   'db not updated because the record has been updated since last refresh.
    End If
    AddChildEntity = 1  'success
  End Function

  Public Shared Function RemoveChildEntity(ByRef objEDI As Object, ByRef enParent As Object, ByRef enChld As Object, ByVal sChildPropertyName As String) As Integer
    'Removes a child entity enChld from the Parent entity enParent
    Dim sJoinTable As String = String.Empty
    strAssoc = objEDI.GetAssociation(TypeName(enParent), TypeName(enChld), sChildPropertyName, sJoinTable)
    If strAssoc = "M-M" And sJoinTable <> "CSR" Then  'JTFK's
      objEDI.mObjectForJoinTable = objEDI.CreateObjectFromString(sJoinTable)
      RemoveChildEntity = RemoveChildEntityViaJoin(objEDI, sJoinTable, enParent, enChld, sChildPropertyName)
      Exit Function
    End If
    If enParent.ChildEntities(sChildPropertyName).Count = 0 Then Exit Function
    Call enParent.ChildEntities(sChildPropertyName).Remove(enChld)

    Call BuildChildEntityString(enParent, sChildPropertyName)
    nUpdate = enParent.mContainer.UpdateDB(enParent)
    If nUpdate = 1 Then
      RemoveChildEntity = nUpdate
      Exit Function
    End If
    Call enParent.ChildEntities(sChildPropertyName).Add(enChld)  'revert

    Call BuildChildEntityString(enParent, sChildPropertyName)
    MsgBox("This record has been updated since you last refreshed. Please refresh again and then update.")
    RemoveChildEntity = nUpdate
  End Function

  Public Shared Function ReOrderChildEntities(ByRef enParent As Object, ByRef enChlds As Object, ByVal strChild As String) As Integer
    'can only be used for CSR childs.
    Dim strChildString As String
    strChildString = enParent.ChildEntityString(strChild)  'preserve in case needing to revert

    Call BuildChildEntityString(enParent, strChild)
    nUpdate = enParent.mContainer.UpdateDB(enParent)
    If nUpdate = 0 Then  'revert back the entity string, the childentities need not be reverted

      Call BuildChildEntityString(enParent, strChildString)
    End If
    ReOrderChildEntities = nUpdate
  End Function

  Public Shared Sub BuildChildEntityObjects(ByRef objEDI As Object, ByRef objEnt As Object, ByVal sChildProperty As String, ByVal strCSR As String)     'ok
    'builds child entities from the CSV's of the child property s
    Dim ent As Object
    Dim n As Integer, i As Integer
    objEnt.ChildEntities(sChildProperty).removeall()
    'Dim sJoinTable As String
    Dim sArray() As String
    sArray = Split(strCSR, ";")
    n = UBound(sArray)
    For i = 0 To n  'removed ';' from both ends.
      If (strCSR = "") Or (strCSR = "0") Then Exit For
      ent = objEDI.CreateObjectFromString(sChildProperty)
      ent = ent.mContainer.item(sArray(i))
      Call objEnt.ChildEntities(sChildProperty).Add(ent)
    Next i
    objEnt.ChildEntityString(sChildProperty, strCSR)
  End Sub

  Public Shared Sub BuildChildEntityString(ByRef entPar As Object, ByVal enChlds As String)
    Dim en As Object
    Dim sSp As String
    sSp = ""
    For Each en In entPar.ChildEntities(enChlds)
      sSp = sSp & en.ID & ";"
    Next en
    If Len(sSp) > 0 Then sSp = Mid(sSp, 1, Len(sSp) - 1)
    entPar.ChildEntityString(enChlds, sSp)
  End Sub

  Shared Function GetItemForJoinTable(ByRef objEDI As Object, ByVal enJT As String, ByRef enPar As Object, ByRef enChld As Object, ByVal strChldPropertyName As String) As Object
    'enPar: parent entity
    'enChld: Child Entity
    Dim sParentPropertyName
    sParentPropertyName = objEDI.GetChildPropertyName(enJT, TypeName(enPar))
    Dim oJT As Object
    oJT = objEDI.CreateObjectFromString(enJT)
    Dim eP As Object
    For Each eP In oJT.mContainer
      If eP.ChildEntityString(sParentPropertyName) = enPar.ID And eP.ChildEntityString(strChldPropertyName) = enChld.ID Then
        GetItemForJoinTable = eP
        Exit Function
      End If
    Next eP
    GetItemForJoinTable = Nothing
  End Function

End Class


