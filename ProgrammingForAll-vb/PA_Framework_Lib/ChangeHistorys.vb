Option Explicit On
Public Class ChangeHistorys
  Inherits System.Collections.CollectionBase
  Private _sNoHashTable As New Hashtable

  Public Sub Add(ByRef val As ChangeHistory)
    Me.List.Add(val)
    _sNoHashTable.Add(val.ID, val)
  End Sub

  Public Sub RemoveAll()
    Me.List.Clear()
    _sNoHashTable.Clear()
  End Sub

  Public Function Load()
    Dim eHist_ As ChangeHistory
    Dim nLoadorder As Integer
    gStrSqlCall = GetSQLQuery("Load", "EntityItem_1")

    cnn.Open(MGlobals.GetAppConnString)
    rst.CursorLocation = adUseClient
    rst.Open(gStrSqlCall, cnn, ADODB.CursorTypeEnum.adOpenStatic, 1)
    Me.RemoveAll()
    rst.MoveFirst()
    'if this entity has a M-1 rel with any other entity, then a blank entity must be present for cbos.
    'If cEntityDataItems.ParentHasM1Relationship("Course") Then
    '  eHist_ = New ChangeHistory   'this member is loaded for 0 index entries cbo's for M-1 relationships
    '  eHist_.Loadorder = 0
    '  eHist_.ID = ""
    '  Call cChangeHistorys.Add(eHist_)
    'End If
    nLoadorder = 1
    Do While Not rst.EOF
      eHist_ = New ChangeHistory

      Call Me.LoadEntityItemsForThisEntity(eHist_)

      'eHist_.Lastupdate = rst.Fields("sLastUpdate").Value
      eHist_.Loadorder = nLoadorder
      'eHist_.mContainer = Me
      Call cChangeHistorys.Add(eHist_)
      nLoadorder = nLoadorder + 1
      rst.MoveNext()
    Loop
    rst.Close()
    cnn.Close()
    'Build the domain model now
    'Call Me.BuildDomainModel
  End Function
  Function GetSQLQuery(ByVal strFunc As String, ByVal strID As String) As String
    'This SQL query is used in 3 functions: LoadSingleEntity(). Load() and UpdateDB()
    Dim strTemp As String
    'PH: For Each Entity Item, add to the Select
    ' strDBO & "EntityItem_2 as sEntityItem_2, " & _
    'PH: End
    'sPH_col_Sql:
    Dim Date1 As DateTime
    Date1 = "2011-11-11"
    'DateTime.TryParse("12/10/2011", Date1)
    strTemp = " Select " & _
              strDBO & "ID as sNo, " & _
                  strDBO & "DateTime, " & _
                  strDBO & "LoginID, " & _
                  strDBO & "TableID, " & _
                  strDBO & "KeyFieldNo, " & _
                  strDBO & "Changes, " & _
                  strDBO & "Operation " & _
                  "FROM " & strDBO & "UpdateLog" & _
                  " where " & strDBO & "DateTime >= " & Date1 & _
                  " ORDER BY " & strDBO & "DateTime desc"
    'sPH_col_Sql:End

    'If strFunc = "LoadSingleEntity" Or strFunc = "UpdateDB" Then
    '  strTemp = strTemp & " WHERE " & strDBO & "ID like " & strID
    'ElseIf strFunc = "Load" Then
    '  strTemp = strTemp & " ORDER BY " & strDBO & strID
    'End If
    GetSQLQuery = strTemp
  End Function

  Function LoadEntityItemsForThisEntity(ByVal ent As ChangeHistory)
    ''PH:For Each Entity Item. Add it for loading
    'eCrse_.EntityItem_1 = "" & rst.Fields("sEntityItem_1")
    ''PH End
    'sPH_col_Load:
    'ent.ID = "" & rst.Fields("sID").Value
    'ent.EntityItem_1 = "" & rst.Fields("sEntityItem_1").Value
    'ent.EntityItem_2 = "" & rst.Fields("sEntityItem_2").Value
    'ent.Students = "" & rst.Fields("sEntityBs").Value
    'sPH_col_Load:End
    ent.ID = "" & rst.Fields("sNo").Value
    ent.DateTime = rst.Fields("DateTime").Value
    ent.KeyField = "" & rst.Fields("KeyFieldNo").Value
    ent.User = "" & rst.Fields("LoginID").Value
    ent.Table = "" & rst.Fields("TableID").Value
    ent.Changes = "" & rst.Fields("Changes").Value
    ent.Operation = "" & rst.Fields("Operation").Value
    ''PH: For Each Child Entity, add the two lines for building child entities for this object.
    ' Call ent.BuildChildEntityObjects(e3_Par, "" & rst.Fields("sEntity3s"))
    ''PH: End
    'sPH_col_LoadChildEntities:
    'Call ent.BuildChildEntityObjects("Students", "" & rst.Fields("sEntityCs").Value)
    'sPH_col_LoadChildEntities:End
  End Function
End Class
