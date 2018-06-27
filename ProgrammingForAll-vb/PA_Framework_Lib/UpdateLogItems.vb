Option Explicit On
Imports System

Public Class UpdateLogItem

  Public Loadorder As Integer  'Load order

  Public sUpdateID As String
  Public sLoginID As String
  Public sTableID As String
  Public sKeyFieldNumber As String
  Public sChanges As String
  Public sOperation As String
  Public dDateTime As Date
End Class

Public Class UpdateLogItems
  Inherits System.Collections.CollectionBase
  
  Public Sub Add(ByRef val As UpdateLogItem)
    Call Me.List.Add(val)
  End Sub

  Public ReadOnly Property item(ByVal val As Integer) As UpdateLogItem
    Get
      item = Me.List(val)
    End Get

  End Property

  Public Sub Remove(ByRef val As UpdateLogItem)
    Me.List.Remove(val)
  End Sub

  Function LoadFromLastUpdate(ByRef PAsettings As Object, ByVal nPrevUpdateID As Integer) As Integer
    Dim uLi_ As UpdateLogItem = Nothing
    Me.Clear()
    'PH: SQL Call for load.
    Dim nLastUpdate As String
    If nPrevUpdateID = 0 Then
      nLastUpdate = PAsettings.GetSetting("LastUpdateID")
    Else
      nLastUpdate = nPrevUpdateID
    End If

    gStrSqlCallLib = " Select " & _
                      MGlobals.strDBO & "ID, " & _
                      MGlobals.strDBO & "DateTime, " & _
                      MGlobals.strDBO & "LoginID, " & _
                      MGlobals.strDBO & "TableID, " & _
                      MGlobals.strDBO & "KeyFieldNo, " & _
                      MGlobals.strDBO & "Changes, " & _
                      MGlobals.strDBO & "Operation " & _
                      " FROM " & MGlobals.strDBO & "UpdateLog " & _
                      "WHERE " & MGlobals.strDBO & "ID > " & nLastUpdate & _
                      " ORDER BY " & MGlobals.strDBO & "ID" & " Asc"
    'PH: SQL Call for load. end
    PASettings_Lib.cnn.Open(PASettings_Lib.GetAppConnString)
    PASettings_Lib.rst.CursorLocation = adUseClient
    PASettings_Lib.rst.Open(gStrSqlCallLib, PASettings_Lib.cnn, adOpenStatic, 1)
    If PASettings_Lib.rst.RecordCount = 0 Then
      LoadFromLastUpdate = nLastUpdate
      PASettings_Lib.rst.Close()
      PASettings_Lib.cnn.Close()
      Exit Function
    End If
    PASettings_Lib.rst.MoveFirst()
    Do While Not PASettings_Lib.rst.EOF
      uLi_ = New UpdateLogItem
      uLi_.sUpdateID = "" & PASettings_Lib.rst.Fields("ID").Value
      'PH:For Each Entity Item. Add all entity items to load
      'uLi_.dDateTime = "" & PASettings_Lib.rst.Fields("sEntityItem_1")
      uLi_.sLoginID = "" & PASettings_Lib.rst.Fields("LoginID").Value
      uLi_.sTableID = "" & PASettings_Lib.rst.Fields("TableID").Value
      uLi_.sKeyFieldNumber = "" & PASettings_Lib.rst.Fields("KeyFieldNo").Value
      uLi_.sChanges = "" & PASettings_Lib.rst.Fields("Changes").Value
      uLi_.sOperation = "" & PASettings_Lib.rst.Fields("Operation").Value

      'PH:For Each Entity Item. Add all entity items to load. end

      'uLi_.Lastupdate = PASettings_Lib.rst.Fields("sLastUpdate")
      Call Me.Add(uLi_)
      uLi_.Loadorder = Me.Count
      PASettings_Lib.rst.MoveNext()
    Loop
    LoadFromLastUpdate = CInt(uLi_.sUpdateID)
    PASettings_Lib.rst.Close()
    PASettings_Lib.cnn.Close()
  End Function

  Public Function GetLastUpdateID(ByRef setts As Object) As Integer
    Dim strSQL As String
    PASettings_Lib = setts
    strSQL = "select max(ID) as MaxNo from updatelog"

    PASettings_Lib.cnn.Open(PASettings_Lib.GetAppConnString)
    'cnn.Open(gStrConnGenEBA)
    PASettings_Lib.rst.CursorLocation = adUseClient
    PASettings_Lib.rst.Open(strSQL, PASettings_Lib.cnn, adOpenStatic, adLockOptimistic)
    'Now set the two global variables
    GetLastUpdateID = PASettings_Lib.rst.Fields("MaxNo").Value
    PASettings_Lib.rst.Close()
    PASettings_Lib.cnn.Close()

  End Function

  Function UpdateLogTable(ByVal sTableID As String, ByVal sOperation As String, ByVal recID As Integer, ByVal strChanges As String) As Integer
    'updates the log table 'UpdateLog' for recording history of changes made  by users.
    'nTableID to be assigned to each entity
    'sOperation="U" for update of existing entity, "N" for creation of new entity.
    'recID= ID of the record in the table that was updated or created
    'In addition to the above info, it also logs the current time and the user ID of the user who made the change.
    Dim strSQL As String

    PASettings_Lib.rst.Open("UpdateLog", PASettings_Lib.cnn, , , adCmdTable)
    PASettings_Lib.rst.AddNew()
    PASettings_Lib.rst.Fields("DateTime").Value = Now
    'PASettings_Lib.rst.Fields("LoginID").Value = MGlobals.LoggedInUser.ID    'frmSettings.LoginName  'strLoginID
    PASettings_Lib.rst.Fields("TableID").Value = sTableID  'id of the table
    PASettings_Lib.rst.Fields("KeyFieldNo").Value = recID
    PASettings_Lib.rst.Fields("Changes").Value = strChanges
    PASettings_Lib.rst.Fields("Operation").Value = sOperation  '
    PASettings_Lib.rst.Update()
    strChanges = ""
    PASettings_Lib.rst.Close()
    strSQL = "select max(ID) as MaxNo from Updatelog"
    PASettings_Lib.rst.CursorLocation = adUseClient
    PASettings_Lib.rst.Open(strSQL, PASettings_Lib.cnn, adOpenStatic, adLockOptimistic)
    UpdateLogTable = PASettings_Lib.rst.Fields("MaxNo").Value
    PASettings_Lib.rst.Close()
  End Function

End Class

