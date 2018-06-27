Option Explicit On
Imports System
Imports System.Collections.Generic
Imports System.Runtime.InteropServices

Public Class Instructors
  Inherits System.Collections.CollectionBase

  Private _sNoHashTable As New Hashtable

  Dim nDecLoadOrder As Integer  'comment
  Dim strChanges As String

  Public Event Changed()
  Public Sub Add(ByRef val As Instructor)
    Me.List.Add(val)
    _sNoHashTable.Add(val.ID, val)
  End Sub

  Public Property Item(ByVal sid As String) As Instructor
    Get
      Return _sNoHashTable.Item(sid)
    End Get
    Set(ByVal value As Instructor)
      _sNoHashTable.Item(sid) = value
    End Set
  End Property

  Public Sub Remove(ByRef val As Instructor)
    Me.List.Remove(val)
    _sNoHashTable.Remove(val.ID)
  End Sub

  Public Sub RemoveAll()
    Me.List.Clear()    'alway remove the first item
    _sNoHashTable.Clear()
  End Sub

  Public Function LoadSingleEntity(ByVal sId As String, ByVal sOperation As String) As Integer
    Dim eInst_ As Instructor

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
      eInst_ = New Instructor
    ElseIf sOperation = "U" Then
      eInst_ = Me.Item(sId)
    End If
    Call Me.LoadEntityItemsForThisEntity(eInst_)

    eInst_.Lastupdate = rst.Fields("LastUpdate").Value
    eInst_.mContainer = Me

    eInst_.Loadorder = Me.Count
    If sOperation = "N" Then
      Call cInstructors.Add(eInst_)
    End If
    rst.Close()
    cnn.Close()
    LoadSingleEntity = 1
  End Function

  Public Function Load()
    Dim eInst_ As Instructor
    Dim nLoadorder As Integer
    gStrSqlCall = GetSQLQuery("Load", "ID")

    cnn.Open(MGlobals.GetAppConnString)
    rst.CursorLocation = adUseClient
    rst.Open(gStrSqlCall, cnn, ADODB.CursorTypeEnum.adOpenStatic, 1)
    Call Me.RemoveAll()
    rst.MoveFirst()
    'if this entity has a M-1 rel with any other entity, then a blank entity must be present for cbos.
    If cEntityDataItems.ParentHasM1Relationship("Instructor") Then
      eInst_ = New Instructor   'this member is loaded for 0 index entries cbo's for M-1 relationships
      eInst_.Loadorder = 0
      eInst_.ID = ""
      Call cInstructors.Add(eInst_)
    End If
    nLoadorder = 1
    Do While Not rst.EOF
      eInst_ = New Instructor

      Call Me.LoadEntityItemsForThisEntity(eInst_)
      eInst_.Lastupdate = rst.Fields("LastUpdate").Value
      eInst_.Loadorder = nLoadorder
      eInst_.mContainer = Me
      Call cInstructors.Add(eInst_)
      nLoadorder = nLoadorder + 1
      rst.MoveNext()
    Loop
    rst.Close()
    cnn.Close()
    'Build the domain model now
    'Call Me.BuildDomainModel
  End Function

  Sub BuildDomainModel()
    '  Dim eb_ As Instructor
    '  Dim e1_ As Entity1
    '  'now add eCrse_s to e1_s
    '  For Each e1_ In cEntity1s.Items  'assumed cEntity1s have been loaded
    '    Set e1_.ChildCourses = e1_.BuildChildEntityObjects(e1_.ChildCourses, e1_.ChildEntityString(curreInst_))
    '    'For Each eInst_ In e1_.ChildCourses.Items
    '    ' If Not eInst_.ContainsParentEntity(e1_) Then  'BuildObjectModel is called from AddtoDB() too. hence to check for already existing.
    '    '  Call eInst_.ParentEntity1s.Add(e1_)   'add parents to Instructor
    '    'End If
    '    'Next eInst_
    '  Next e1_
  End Sub

  Public Function UpdateDB(ByRef eInst_ As Instructor) As Integer
    'Updates the database table for the entity'.
    'Return 0 if the db record has been updated by other user.
    'Return the loadorder if successfull in update.
    Dim nCount As Integer
    gbChangesMade = False
    gStrSqlCall = GetSQLQuery("UpdateDB", eInst_.ID)
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
    If rst.Fields("LastUpdate").Value <> eInst_.Lastupdate Then
      rst.Close()
      cnn.Close()
      UpdateDB = 0
      Exit Function
    Else
      rst.Fields("LastUpdate").Value = Now
      eInst_.Lastupdate = rst.Fields("LastUpdate").Value  'update the local last update so you can mame changes without refreshing.
    End If
    strChanges = ""

    Call RecordChanges("Update", eInst_)

    rst.Update()
    rst.Close()
    MGlobals.curreInst_ = eInst_

    'No need to enter a record in update log if no change has been made
    If gbChangesMade = True Then
      AppSettings.LastUpdateID = gUpdateLogTable("eInst_", "U", eInst_.ID)
    End If
    cnn.Close()
    Call DoChanged()
    UpdateDB = nCount
  End Function

  Public Function AddtoDB(ByRef eInst_ As Instructor) As Integer
    Dim strSQL As String
    Dim nCount As Integer
    strSQL = "select Id from Instructors"
    Try
      If cnn.State = ADODB.ObjectStateEnum.adStateClosed Then
        cnn.Open(MGlobals.GetAppConnString)
      End If
    Catch
      MsgBox(Err.Description)
    End Try

    rst.Open(Source:=strSQL, ActiveConnection:=cnn, CursorType:=adOpenKeySet, LockType:=adLockOptimistic)
    nCount = rst.RecordCount
    If cEntityDataItems.ParentHasM1Relationship("Instructor") Then nCount = nCount + 1
    If nCount <> cInstructors.Count - 1 Then  'other user has added a record
      rst.Close()
      cnn.Close()
      AddtoDB = 0
      Exit Function
    End If
    rst.Close()
    strSQL = "select max(ID) as MaxNo from Instructors"
    Dim nMax As Integer
    rst.CursorLocation = adUseClient
    rst.Open(strSQL, cnn, ADODB.CursorTypeEnum.adOpenStatic, adLockOptimistic)
    nMax = rst.Fields("MaxNo").Value
    rst.Close()
    eInst_.ID = nMax + 1
    'now add the record
    rst.CursorType = adOpenKeySet
    rst.LockType = adLockOptimistic
    rst.Open("Instructors", cnn, , , adCmdTable)
    rst.AddNew()

    Call BLFunctions.gRecordChanges1("Add", "ID", eInst_.ID, "id:") 'ID needs to be directly done because recordchanges does not cover ID
    Call RecordChanges("Add", eInst_)

    rst("LastUpdate").Value = Now
    eInst_.Lastupdate = rst("LastUpdate").Value
    eInst_.mContainer = Me
    rst.Update()
    rst.Close()

    Call Me.Add(eInst_)
    eInst_.Loadorder = Me.Count - nDecLoadOrder

    Call Me.BuildDomainModel()
    MGlobals.curreInst_ = eInst_
    'now update the log
    AppSettings.LastUpdateID = gUpdateLogTable("eInst_", "N", eInst_.ID)
    cnn.Close()
    Call DoChanged()
    AddtoDB = nCount + 1
  End Function

  Public Function DeleteFromDB(ByVal eInst_ As Instructor) As Integer
    Dim lAffected As Long
    gStrSqlCall = " Delete " & _
                  "FROM " & strDBO & "Instructors " & _
                  " where " & strDBO & "ID LIKE '" & eInst_.ID & "'"

    MGlobals.cmd.ActiveConnection = cnn
    MGlobals.cmd.CommandText = gStrSqlCall
    MGlobals.cmd.CommandType = ADODB.CommandTypeEnum.adCmdText
    rst = MGlobals.cmd.Execute(lAffected, ADODB.ExecuteOptionEnum.adExecuteNoRecords)
    If lAffected = 1 Then
      Dim sEbID As String
      sEbID = eInst_.ID
      MGlobals.gstrChanges = "Deleted: " & "ID" & _
                            "fn:" & eInst_.Comments 
      Call cInstructors.Remove(eInst_)
      cnn.Open(gStrConnGenEBA)
      AppSettings.LastUpdateID = gUpdateLogTable("eInst_", "D", sEbID)
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
              strDBO & "Persons , " & _
              strDBO & "Comments, " & _
              strDBO & "LastUpdate " & _
              "FROM " & strDBO & "Instructors "

    If strFunc = "LoadSingleEntity" Or strFunc = "UpdateDB" Then
      strTemp = strTemp & " WHERE " & strDBO & "ID like " & strID
    ElseIf strFunc = "Load" Then
      strTemp = strTemp & " ORDER BY " & strDBO & strID
    End If
    GetSQLQuery = strTemp
  End Function

  Sub LoadEntityItemsForThisEntity(ByVal ent As Instructor)

    ent.ID = "" & rst.Fields("ID").Value
    ent.Comments = "" & rst.Fields("Comments").Value

    Call ent.BuildChildEntityObjects("Persons", "" & rst.Fields("Persons").Value)
  End Sub

  Public Sub RaiseChanged()
    RaiseEvent Changed()
  End Sub

  Private Sub RecordChanges(ByVal sOperation As String, ByVal ent As Instructor)
    Call BLFunctions.gRecordChanges1(sOperation, "Comments", ent.Comments, "cmnts:")
    'Call BLFunctions.gRecordChanges1(sOperation, "EntityItem_2", ent.EntityItem_2, "ei2:")
    'Call BLFunctions.gRecordChanges1(sOperation, "EntityCs", ent.ChildEntityString("Students"), "Stds:")
    Call BLFunctions.gRecordChanges1(sOperation, "Persons", ent.ChildEntityString("Persons"), "Prsns:")
  End Sub

End Class
