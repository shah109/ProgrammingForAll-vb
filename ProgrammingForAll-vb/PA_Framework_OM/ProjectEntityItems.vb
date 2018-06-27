Option Explicit On
Imports PA_Framework_Lib
Imports PA_Framework_OM.OMGlobals
Imports System.Runtime.InteropServices

<ClassInterface(ClassInterfaceType.AutoDual)> _
Public Class ProjectEntityItems
  Inherits PAEnts

  Overrides Sub LoadEntityItemsForThisEntity(ByRef ent As Object)

    ent.ID = "" & AppSettings.rst.Fields("ID").Value
    ent.DBColName = "" & AppSettings.rst.Fields("DBColName").Value
    ent.DBColNameType = "" & AppSettings.rst.Fields("DBColNameType").Value
    ent.ChildName = "" & AppSettings.rst.Fields("ChildName").Value
    ent.ChildRelationship = "" & AppSettings.rst.Fields("ChildRelationship").Value
    ent.RelationshipType = "" & AppSettings.rst.Fields("RelationshipType").Value
    'ent.SQLName = "" & AppSettings.rst.Fields("SQLName").Value
    ent.InternalName = "" & AppSettings.rst.Fields("InternalName").Value
    ent.InternalNameType = "" & AppSettings.rst.Fields("InternalNameType").Value
    ent.PropertyName = "" & AppSettings.rst.Fields("PropertyName").Value
    ent.PropertyNameType = "" & AppSettings.rst.Fields("PropertyNameType").Value
    ent.ULName = "" & AppSettings.rst.Fields("ULName").Value
    ent.SheetDisplayOrder = "" & AppSettings.rst.Fields("SheetDisplayOrder").Value
    ent.FormControlName = "" & AppSettings.rst.Fields("FormControlName").Value
    'ent.ChildListSetNeeded = "" & AppSettings.rst.Fields("ChildListSetNeeded").Value
    'ent.ParentListSetNeeded = "" & AppSettings.rst.Fields("ParentListSetNeeded").Value
    'ent.ControlReference = "" & AppSettings.rst.Fields("ControlReference").Value
    ent.GenerateFlag = "" & AppSettings.rst.Fields("GenerateFlag").Value
  End Sub

  Overrides Sub Recordchanges(ByVal sOperation As String, ByRef ent As Object)
    Call gRecordChanges(sOperation, "ChildName", ent.ChildName, "chldNme")
    Call gRecordChanges(sOperation, "ChildRelationship", ent.ChildRelationship, "Chldrel")
    Call gRecordChanges(sOperation, "RelationshipType", ent.RelationshipType, "rType")
    Call gRecordChanges(sOperation, "PropertyName", ent.PropertyName, "PrpNme")
    Call gRecordChanges(sOperation, "PropertyNameType", ent.PropertyNameType, "pnt")
    Call gRecordChanges(sOperation, "InternalName", ent.InternalName, "iNme")
    Call gRecordChanges(sOperation, "InternalNameType", ent.InternalNameType, "iNTpe")
    Call gRecordChanges(sOperation, "DBColName", ent.DBColName, "DBColNme")
    Call gRecordChanges(sOperation, "DBColNameType", ent.DBColNameType, "DBColTpe")
    Call gRecordChanges(sOperation, "ULName", ent.ULName, "uln")
    Call gRecordChanges(sOperation, "GenerateFlag", ent.GenerateFlag, "gf")
    'Call gRecordChanges(sOperation, "SQLName", ent.SQLName, "ei2:")
    Call gRecordChanges(sOperation, "SheetDisplayOrder", ent.SheetDisplayOrder, "sdo")
    Call gRecordChanges(sOperation, "FormControlName", ent.FormControlName, "ei2:")
    'Call gRecordChanges(sOperation, "ChildListSetNeeded", ent.ChildListSetNeeded, "ei2:")
    'Call gRecordChanges(sOperation, "ParentListSetNeeded", ent.ParentListSetNeeded, "ei2:")
    'Call gRecordChanges(sOperation, "ControlReference", ent.ControlReference, "ei2:")
  End Sub

  Overrides Function GetSelectDBItems() As String
    GetSelectDBItems = _
              strDBO & " ID " & _
              strDBO & ",ChildName " & _
              strDBO & ",ChildRelationship " & _
              strDBO & ",RelationshipType " & _
              strDBO & ",PropertyName " & _
              strDBO & ",PropertyNameType " & _
              strDBO & ",InternalName " & _
              strDBO & ",InternalNameType " & _
              strDBO & ",DBColName " & _
              strDBO & ",DBColNameType " & _
              strDBO & ",ULName " & _
              strDBO & ",SheetDisplayOrder " & _
              strDBO & ",FormControlName " & _
              strDBO & ",GenerateFlag " & _
              strDBO & ",LastUpdate "

    'strDBO & ",SQLName " & _
    'strDBO & ",ULName " & _

    'strDBO & ",FormControlName " & _
    'strDBO & ",ChildListSetNeeded " & _
    'strDBO & ",ParentListSetNeeded " & _
    'strDBO & ",ControlReference " & _
  End Function

  Overrides Function GetSelectFromTable() As String
    GetSelectFromTable = " FROM " & strDBO & "ProjectEntityItems "
  End Function

  Overrides Function CreateNewEntity() As PAEnt
    CreateNewEntity = New ProjectEntityItem
  End Function

  Overrides Sub SetCurrentEntity(ByRef ent As PAEnt)
    currePrjEntItm_ = ent
  End Sub

  Overrides Sub BuildDomainModel()

  End Sub
End Class
