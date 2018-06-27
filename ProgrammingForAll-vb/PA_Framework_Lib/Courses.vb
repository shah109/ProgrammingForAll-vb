Option Explicit On
Imports System
Imports System.Collections.Generic
Imports System.Runtime.InteropServices

Public Class Courses
  Inherits System.Collections.CollectionBase

  Private _sNoHashTable As New Hashtable

  Dim nDecLoadOrder As Integer  'comment
  Dim strChanges As String

  Public Event Changed()
  Public Sub Add(ByRef val As Course)
    Me.List.Add(val)
    _sNoHashTable.Add(val.ID, val)
  End Sub


  'Public ReadOnly Property Item(ByVal index As Integer) As Course
  '  Get
  '    Return Me.List.Item(index)
  '  End Get
  'End Property

  Public Property Item(ByVal sid As String) As Course
    Get
      Return _sNoHashTable.Item(sid)
    End Get
    Set(ByVal value As Course)
      _sNoHashTable.Item(sid) = value
    End Set
  End Property

  Public Sub Remove(ByRef val As Course)
    Me.List.Remove(val)
    _sNoHashTable.Remove(val.ID)
  End Sub

  Public Sub RemoveAll()
    Me.List.Clear()    'alway remove the first item
    _sNoHashTable.Clear()
  End Sub

  Public Function LoadSingleEntity(ByVal sId As String, ByVal sOperation As String) As Integer
    Dim eCrse_ As Course

    gStrSqlCall = GetSQLQuery("LoadSingleEntity", sId)
    cnn.Open(gStrConnGenEBA)
    rst.CursorLocation = adUseClient
    rst.Open(gStrSqlCall, cnn, _
             adOpenStatic, 1)
    If rst.RecordCount <> 1 Then
      LoadSingleEntity = 0  '
      rst.Close()
      cnn.Close()
      Exit Function
    End If
    If sOperation = "N" Then
      eCrse_ = New Course
    ElseIf sOperation = "U" Then
      eCrse_ = Me.Item(sId)
    End If
    Call Me.LoadEntityItemsForThisEntity(eCrse_)

    eCrse_.Lastupdate = rst.Fields("LastUpdate").Value
    eCrse_.mContainer = Me

    eCrse_.Loadorder = Me.Count
    If sOperation = "N" Then
      Call cCourses.Add(eCrse_)
    End If
    rst.Close()
    cnn.Close()
    LoadSingleEntity = 1
  End Function

  Public Function Load()
    Dim eCrse_ As Course
    Dim nLoadorder As Integer
    gStrSqlCall = GetSQLQuery("Load", "ID")

    cnn.Open(MGlobals.GetAppConnString)
    rst.CursorLocation = adUseClient
    rst.Open(gStrSqlCall, cnn, ADODB.CursorTypeEnum.adOpenStatic, 1)
    Call Me.RemoveAll()
    rst.MoveFirst()
    'if this entity has a M-1 rel with any other entity, then a blank entity must be present for cbos.
    If cEntityDataItems.ParentHasM1Relationship("Course") Then
      eCrse_ = New Course   'this member is loaded for 0 index entries cbo's for M-1 relationships
      eCrse_.Loadorder = 0
      eCrse_.ID = ""
      Call cCourses.Add(eCrse_)
    End If
    nLoadorder = 1
    Do While Not rst.EOF
      eCrse_ = New Course

      Call Me.LoadEntityItemsForThisEntity(eCrse_)

      eCrse_.Lastupdate = rst.Fields("LastUpdate").Value
      eCrse_.Loadorder = nLoadorder
      eCrse_.mContainer = Me
      Call cCourses.Add(eCrse_)
      nLoadorder = nLoadorder + 1
      rst.MoveNext()
    Loop
    rst.Close()
    cnn.Close()
    'Build the domain model now
    'Call Me.BuildDomainModel
  End Function

  Sub BuildDomainModel()
    '  Dim eb_ As Course
    '  Dim e1_ As Entity1
    '  'now add eCrse_s to e1_s
    '  For Each e1_ In cEntity1s.Items  'assumed cEntity1s have been loaded
    '    Set e1_.ChildCourses = e1_.BuildChildEntityObjects(e1_.ChildCourses, e1_.ChildEntityString(CurreCrse_))
    '    'For Each eCrse_ In e1_.ChildCourses.Items
    '    ' If Not eCrse_.ContainsParentEntity(e1_) Then  'BuildObjectModel is called from AddtoDB() too. hence to check for already existing.
    '    '  Call eCrse_.ParentEntity1s.Add(e1_)   'add parents to Course
    '    'End If
    '    'Next eCrse_
    '  Next e1_
  End Sub

  Public Function UpdateDB(ByRef eCrse_ As Course) As Integer
    'Updates the database table for the entity'.
    'Return 0 if the db record has been updated by other user.
    'Return the loadorder if successfull in update.
    Dim nCount As Integer
    gbChangesMade = False
    gStrSqlCall = GetSQLQuery("UpdateDB", eCrse_.ID)
    If cnn.State = ADODB.ObjectStateEnum.adStateClosed Then
      cnn.Open(MGlobals.GetAppConnString)
    End If

    rst.Open(Source:=gStrSqlCall, ActiveConnection:=cnn, _
             CursorType:=adOpenKeySet, LockType:=adLockOptimistic)
    nCount = rst.RecordCount
    If nCount <> 1 Then
      rst.Close()
      cnn.Close()
      UpdateDB = 0
      Exit Function
    End If
    If rst.Fields("LastUpdate").Value <> eCrse_.Lastupdate Then
      rst.Close()
      cnn.Close()
      UpdateDB = 0
      Exit Function
    Else
      rst.Fields("LastUpdate").Value = Now
      eCrse_.Lastupdate = rst.Fields("LastUpdate").Value  'update the local last update so you can mame changes without refreshing.
    End If
    strChanges = ""

    Call RecordChanges("Update", eCrse_)

    rst.Update()
    rst.Close()
    MGlobals.curreCrse_ = eCrse_

    'No need to enter a record in update log if no change has been made
    If gbChangesMade = True Then
      AppSettings.LastUpdateID = gUpdateLogTable("eCrse_", "U", eCrse_.ID)
    End If
    cnn.Close()
    Call DoChanged()
    UpdateDB = nCount
  End Function

  Public Function AddtoDB(ByRef eCrse_ As Course) As Integer
    Dim strSQL As String
    Dim nCount As Integer
    strSQL = "select Id from Courses"
    Try
      If cnn.State = ADODB.ObjectStateEnum.adStateClosed Then
        cnn.Open(MGlobals.GetAppConnString)
      End If
    Catch
      MsgBox(Err.Description)
    End Try

    rst.Open(Source:=strSQL, ActiveConnection:=cnn, CursorType:=adOpenKeySet, LockType:=adLockOptimistic)
    nCount = rst.RecordCount
    If cEntityDataItems.ParentHasM1Relationship("Course") Then nCount = nCount + 1
    If nCount <> cCourses.Count - 1 Then  'other user has added a record
      rst.Close()
      cnn.Close()
      AddtoDB = 0
      Exit Function
    End If
    rst.Close()
    strSQL = "select max(ID) as MaxNo from Courses"
    Dim nMax As Integer
    rst.CursorLocation = adUseClient
    rst.Open(strSQL, cnn, ADODB.CursorTypeEnum.adOpenStatic, adLockOptimistic)
    nMax = rst.Fields("MaxNo").Value
    rst.Close()
    eCrse_.ID = nMax + 1
    'now add the record
    rst.CursorType = adOpenKeySet
    rst.LockType = adLockOptimistic
    rst.Open("Courses", cnn, , , adCmdTable)
    rst.AddNew()

    Call BLFunctions.gRecordChanges1("Add", "ID", eCrse_.ID, "id:") 'ID needs to be directly done because recordchanges does not cover ID
    Call RecordChanges("Add", eCrse_)

    rst("LastUpdate").Value = Now
    eCrse_.Lastupdate = rst("LastUpdate").Value
    eCrse_.mContainer = Me
    rst.Update()
    rst.Close()

    Call Me.Add(eCrse_)
    eCrse_.Loadorder = Me.Count - nDecLoadOrder

    Call Me.BuildDomainModel()
    MGlobals.curreCrse_ = eCrse_
    'now update the log
    AppSettings.LastUpdateID = gUpdateLogTable("eCrse_", "N", eCrse_.ID)
    cnn.Close()
    Call DoChanged()
    AddtoDB = nCount + 1
  End Function

  Public Function DeleteFromDB(ByVal eCrse_ As Course) As Integer
    Dim lAffected As Long
    gStrSqlCall = " Delete " & _
                  "FROM " & strDBO & "Courses " & _
                  " where " & strDBO & "ID LIKE '" & eCrse_.ID & "'"

    MGlobals.cmd.ActiveConnection = cnn
    MGlobals.cmd.CommandText = gStrSqlCall
    MGlobals.cmd.CommandType = ADODB.CommandTypeEnum.adCmdText
    rst = MGlobals.cmd.Execute(lAffected, ADODB.ExecuteOptionEnum.adExecuteNoRecords)
    If lAffected = 1 Then
      Dim sEbID As String
      sEbID = eCrse_.ID
      MGlobals.gstrChanges = "Deleted: " & "ID" & _
                            "fn:" & eCrse_.EntityItem_1 & _
                            ", ln:" & eCrse_.EntityItem_2
      Call cCourses.Remove(eCrse_)
      cnn.Open(gStrConnGenEBA)
      AppSettings.LastUpdateID = gUpdateLogTable("eCrse_", "D", sEbID)
      cnn.Close()
      Call DoChanged()
    End If
    MGlobals.cmd = Nothing
    DeleteFromDB = lAffected
  End Function

  Function GetSQLQuery(ByVal strFunc As String, ByVal strID As String) As String
    'This SQL query is used in 3 functions: LoadSingleEntity(). Load() and UpdateDB()
    Dim strTemp As String

    strTemp = " Select " & _
              strDBO & "ID , " & _
              strDBO & "EntityItem_1 , " & _
              strDBO & "EntityItem_2 , " & _
              strDBO & "EntityCs , " & _
              strDBO & "EntityUs , " & _
              strDBO & "LastUpdate " & _
              "FROM " & strDBO & "Courses "

    If strFunc = "LoadSingleEntity" Or strFunc = "UpdateDB" Then
      strTemp = strTemp & " WHERE " & strDBO & "ID like " & strID
    ElseIf strFunc = "Load" Then
      strTemp = strTemp & " ORDER BY " & strDBO & strID
    End If
    GetSQLQuery = strTemp
  End Function

  Sub LoadEntityItemsForThisEntity(ByVal ent As Course)

    ent.ID = "" & rst.Fields("ID").Value
    ent.EntityItem_1 = "" & rst.Fields("EntityItem_1").Value
    ent.EntityItem_2 = "" & rst.Fields("EntityItem_2").Value

    Call ent.BuildChildEntityObjects("Students", "" & rst.Fields("EntityCs").Value)
    Call ent.BuildChildEntityObjects("Persons", "" & rst.Fields("EntityUs").Value)
  End Sub

  Public Sub RaiseChanged()
    RaiseEvent Changed()
  End Sub

  Private Sub RecordChanges(ByVal sOperation As String, ByVal ent As Course)
    Call BLFunctions.gRecordChanges1(sOperation, "EntityItem_1", ent.EntityItem_1, "ei1:")
    Call BLFunctions.gRecordChanges1(sOperation, "EntityItem_2", ent.EntityItem_2, "ei2:")
    Call BLFunctions.gRecordChanges1(sOperation, "EntityCs", ent.ChildEntityString("Students"), "Stds:")
    Call BLFunctions.gRecordChanges1(sOperation, "EntityUs", ent.ChildEntityString("Persons"), "Stds:")
  End Sub

  Public Function GetMe() As Courses
    Return cCourses
  End Function
  Public Function ShowHello()
    MsgBox("Hello")
  End Function

End Class
