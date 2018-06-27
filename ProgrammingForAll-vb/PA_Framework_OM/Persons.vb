Option Explicit On
Imports PA_Framework_Lib
Imports PA_Framework_OM.OMGlobals
Imports System.Runtime.InteropServices

<ClassInterface(ClassInterfaceType.AutoDual)> _
Public Class Persons
  Inherits PAEnts

  Overrides Sub LoadEntityItemsForThisEntity(ByRef ent As Object)
    ent.ID = "" & AppSettings.rst.Fields("ID").Value
    ent.FirstName = "" & AppSettings.rst.Fields("FirstName").Value
    ent.MiddleName = "" & AppSettings.rst.Fields("MiddleName").Value
    ent.LastName = "" & AppSettings.rst.Fields("LastName").Value
    ent.LoginID = "" & AppSettings.rst.Fields("LoginID").Value
    ent.Email = "" & AppSettings.rst.Fields("Email").Value
    ent.Phone = "" & AppSettings.rst.Fields("Phone").Value
    ent.AccessRight = "" & AppSettings.rst.Fields("AccessRight").Value
    ent.DateJoined = "" & AppSettings.rst.Fields("DateJoined").Value
    ent.Remarks = "" & AppSettings.rst.Fields("Remarks").Value
  End Sub

  Overrides Sub RecordChanges(ByVal sOperation As String, ByRef ent As Object)
    Call gRecordChanges(sOperation, "FirstName", ent.FirstName, "fn:")
    Call gRecordChanges(sOperation, "MiddleName", ent.MiddleName, "mn:")
    Call gRecordChanges(sOperation, "LastName", ent.LastName, "ln:")
    Call gRecordChanges(sOperation, "LoginID", ent.LoginID, "logid:")
    Call gRecordChanges(sOperation, "Email", ent.Email, "em:")
    Call gRecordChanges(sOperation, "Phone", ent.Phone, "phn:")
    Call gRecordChanges(sOperation, "AccessRight", ent.AccessRight, "ar:")
    Call gRecordChanges(sOperation, "DateJoined", ent.DateJoined, "dj:")
    Call gRecordChanges(sOperation, "Remarks", ent.Remarks, "rmks:")
  End Sub

  Overrides Function GetSelectDBItems() As String
    GetSelectDBItems = _
      strDBO & " ID " & _
      strDBO & ",FirstName " & _
      strDBO & ",MiddleName " & _
      strDBO & ",LastName " & _
      strDBO & ",LoginID " & _
      strDBO & ",Email " & _
      strDBO & ",Phone " & _
      strDBO & ",AccessRight " & _
      strDBO & ",DateJoined " & _
      strDBO & ",Remarks " & _
      strDBO & ",LastUpdate "
  End Function

  Overrides Function GetSelectFromTable() As String
    GetSelectFromTable = " FROM " & strDBO & "Persons "
  End Function

  Overrides Function CreateNewEntity() As PAEnt
    CreateNewEntity = New Person
  End Function

  Overrides Sub SetCurrentEntity(ByRef ent As PAEnt)
    currePrsn_ = ent
  End Sub

  Overrides Sub BuildDomainModel()

  End Sub
End Class
