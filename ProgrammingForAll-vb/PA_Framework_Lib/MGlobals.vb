Option Explicit On

Public Module MGlobals

  Public PASettings_Lib As Object

  Public gStrAssoc As String  'Association string with other child and parent entities
  Public gStrSqlCallLib As String  'sql call string

  'Public gsChildPropertyName As String  ' to get from Entity Data items
  Public cUpdateLogItems As New UpdateLogItems
  Public Const APPLongName = "The Framework"
  Public Const APPName = "FrmWrk"

  Public Const adLockOptimistic = 3
  Public Const adUseClient = 3
  Public Const adOpenKeySet = 1
  Public Const adCmdTable = 2
  Public Const adOpenStatic = 3

  Public gStrConnGenEBA As String  'Connection string for the GenEBA db.
  Public strDBO As String

  'DB Update related variables
  Public PrevUpdateID As Integer
  Public NewUpdateID As Integer

  Function UpdateLogTable(ByVal sTableID As String, ByVal sOperation As String, ByVal recID As String, ByRef strChanges As String) As Integer
    'updates the log table 'UpdateLog' for recording history of changes made  by users.
    'nTableID to be assigned to each entity
    'sOperation="U" for update of existing entity, "N" for creation of new entity.
    'recID= ID of the record in the table that was updated or created
    'In addition to the above info, it also logs the current time and the user ID of the user who made the change.
    Dim strSQL As String
    Dim nMax As Integer
    strSQL = "select max(ID) as MaxNo from Updatelog"
    PASettings_Lib.rst.CursorLocation = adUseClient
    PASettings_Lib.rst.Open(strSQL, PASettings_Lib.cnn, adOpenStatic, adLockOptimistic)
    nMax = PASettings_Lib.rst.Fields("MaxNo").Value
    PASettings_Lib.rst.Close()

    PASettings_Lib.rst.CursorType = adOpenKeySet
    PASettings_Lib.rst.LockType = adLockOptimistic
    PASettings_Lib.rst.Open("UpdateLog", PASettings_Lib.cnn, , , adCmdTable)
    PASettings_Lib.rst.AddNew()
    PASettings_Lib.rst.Fields("ID").Value = nMax + 1
    PASettings_Lib.rst.Fields("DateTime").Value = Now
    PASettings_Lib.rst.Fields("LoginID").Value = PASettings_Lib.GetSetting(PASettings_Lib.loginname)    'frmSettings.LoginName  'strLoginID
    PASettings_Lib.rst.Fields("TableID").Value = sTableID  'id of the table
    PASettings_Lib.rst.Fields("KeyFieldNo").Value = recID
    PASettings_Lib.rst.Fields("Changes").Value = strChanges
    PASettings_Lib.rst.Fields("Operation").Value = sOperation  '
    PASettings_Lib.rst.Update()
    strChanges = ""
    PASettings_Lib.rst.Close()
    UpdateLogTable = nMax + 1
    PASettings_Lib.SetSetting("LastUpdateID", UpdateLogTable.ToString)
  End Function

  Public Function GetDBUpdates(ByRef objEDI As Object, ByVal nPrevUpdateID As Integer) As Integer
    GetDBUpdates = 0
    PrevUpdateID = nPrevUpdateID
    'PrevUpdateID = CInt(PASettings_Lib.GetSetting("LastUpdateID"))
    Call cUpdateLogItems.LoadFromLastUpdate(PASettings_Lib, PrevUpdateID)
    If cUpdateLogItems.Count > 0 Then
      Call objEDI.DBLoadWithUpdateLogItems(cUpdateLogItems)
      GetDBUpdates = cUpdateLogItems.Count
      PASettings_Lib.SetSetting("LastUpdateID", (PrevUpdateID + GetDBUpdates).ToString) 'updated after the db update is done
    End If
    Return PASettings_Lib.GetSetting("LastUpdateID")
  End Function
End Module
