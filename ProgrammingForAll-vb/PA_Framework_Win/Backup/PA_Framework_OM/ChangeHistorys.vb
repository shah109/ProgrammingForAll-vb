Option Explicit On
Imports PA_Framework_OM.OMGlobals
Imports System.Runtime.InteropServices

<ClassInterface(ClassInterfaceType.AutoDual)> _
Public Class ChangeHistorys
  Inherits System.Collections.CollectionBase
  Dim nMonthsToAccessHistory As Integer

  Private _sNoHashTable As New Hashtable

  Public Sub Add(ByRef val As ChangeHistory)
    Me.List.Add(val)
    _sNoHashTable.Add(val.ID, val)
  End Sub

  Public Sub RemoveAll()
    Me.List.Clear()
    _sNoHashTable.Clear()
  End Sub

  Public Sub Load()
    Dim eHist_ As ChangeHistory
    Dim nLoadorder As Integer
    nMonthsToAccessHistory = CInt(AppSettings.GetSetting("MonthsToAccessHistory"))
    gStrSqlCall = GetSQLQuery("Load", "EntityItem_1")

    'AppSettings.cnn.Open(AppSettings.GetAppConnString)
    If AppSettings.cnn.State = ADODB.ObjectStateEnum.adStateClosed Then
      AppSettings.cnn.Open(AppSettings.GetAppConnString)
    End If

    AppSettings.rst.CursorLocation = ADODB.CursorLocationEnum.adUseClient
    AppSettings.rst.Open(gStrSqlCall, AppSettings.cnn, ADODB.CursorTypeEnum.adOpenStatic, 1)
    Me.RemoveAll()
    AppSettings.rst.MoveFirst()

    nLoadorder = 1
    Do While Not AppSettings.rst.EOF
      eHist_ = New ChangeHistory

      Call Me.LoadEntityItemsForThisEntity(eHist_)

      'eHist_.Lastupdate = AppSettings.rst.Fields("sLastUpdate").Value
      eHist_.Loadorder = nLoadorder
      'eHist_.mContainer = Me
      Call Me.Add(eHist_)
      nLoadorder = nLoadorder + 1
      AppSettings.rst.MoveNext()
    Loop
    AppSettings.rst.Close()
    AppSettings.cnn.Close()
    'Build the domain model now
    'Call Me.BuildDomainModel
  End Sub
  Function GetSQLQuery(ByVal strFunc As String, ByVal strID As String) As String
    'This SQL query is used in 3 functions: LoadSingleEntity(). Load() and UpdateDB()
    Dim strTemp As String

    Dim Date1 As DateTime
    Date1 = Today.AddMonths(-nMonthsToAccessHistory)
    strTemp = " Select " & _
              strDBO & "ID as sNo, " & _
                  strDBO & "DateTime, " & _
                  strDBO & "LoginID, " & _
                  strDBO & "TableID, " & _
                  strDBO & "KeyFieldNo, " & _
                  strDBO & "Changes, " & _
                  strDBO & "Operation " & _
                  "FROM " & strDBO & "UpdateLog" & _
                  " where " & strDBO & "DateTime >= #" & Date1 & "#" & _
                  " ORDER BY " & strDBO & "DateTime desc"

    GetSQLQuery = strTemp
  End Function

  Sub LoadEntityItemsForThisEntity(ByVal ent As ChangeHistory)

    ent.ID = "" & AppSettings.rst.Fields("sNo").Value
    ent.DateTime = AppSettings.rst.Fields("DateTime").Value
    ent.KeyField = "" & AppSettings.rst.Fields("KeyFieldNo").Value
    ent.User = "" & AppSettings.rst.Fields("LoginID").Value
    ent.Table = "" & AppSettings.rst.Fields("TableID").Value
    ent.Changes = "" & AppSettings.rst.Fields("Changes").Value
    ent.Operation = "" & AppSettings.rst.Fields("Operation").Value

  End Sub
End Class
