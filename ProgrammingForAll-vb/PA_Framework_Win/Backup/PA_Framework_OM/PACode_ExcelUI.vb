Option Explicit On
Imports System.IO
Imports System.Text
Imports PA_Framework_OM.OMGlobals
Imports System.Runtime.InteropServices

<ClassInterface(ClassInterfaceType.AutoDual)> _
Public Class PACode_ExcelUI

  Public Shared Function GenerateExcelUI(ByVal sProjectName As String) As String
    GenerateExcelUI = String.Empty

    Dim oPAProject As New PAProject
    Dim oPAEntity As New ProjectEntity
    Dim oPAEntityItem As ProjectEntityItem = Nothing

    For Each pap As PAProject In cPAProjects
      If pap.ProjectName = sProjectName Then
        oPAProject = pap
        Exit For
      End If
    Next

    For Each pae As ProjectEntity In oPAProject.ChildprojectEntities
      oPAEntity = pae
      Dim fs1 As FileStream = New FileStream("c:\Framework\" & oPAEntity.EntityName & "forExcel.txt", FileMode.Create, FileAccess.Write)
      Dim s1 As StreamWriter = New StreamWriter(fs1)

      s1.WriteLine("  'PA_Code Generator Code")
      s1.WriteLine("  '" & Now & ":, " & sProjectName)
      s1.WriteLine()
      s1.WriteLine("Option Explicit")
      s1.WriteLine("Dim {0:C} As {1:C}", pae.EntityShortName, pae.EntityName)

      s1.WriteLine()
      s1.WriteLine("Private Enum eColumnName")
      s1.WriteLine("  HypDetails = 2")
      s1.WriteLine("  HypEdit = 3")
      s1.WriteLine("  Loadorder = 4")
      's1.WriteLine("  ID = 5")

      For Each paei As ProjectEntityItem In pae.mChildProjectEntityItems
        If Trim(paei.SheetDisplayOrder) <> "" Then
          s1.WriteLine("  {0:C} = {1:C} ", paei.PropertyName, CStr(4 + CInt(paei.SheetDisplayOrder)))
        End If
      Next paei
      s1.WriteLine("End Enum")
      s1.WriteLine("'    '////////////////////////////////////////////////////////////////////")
      s1.WriteLine("'    'Standard Functions")
      s1.WriteLine("  Private Sub btnLoadData_Click()")
      s1.WriteLine("    Call LoadAppData()")
      s1.WriteLine("  End Sub")
      s1.WriteLine("")
      s1.WriteLine("  Private Sub btnSettings_Click()")
      s1.WriteLine("    Call frmSettings.Show()")
      s1.WriteLine("  End Sub")
      s1.WriteLine("")
      s1.WriteLine("  Private Sub lblGenEBA_Click()")
      s1.WriteLine("    frmAboutFrmWrk.btnClose.Visible = True")
      s1.WriteLine("    Call frmAboutFrmWrk.Show()")
      s1.WriteLine("  End Sub")
      s1.WriteLine("")
      s1.WriteLine("  Public Sub Worksheet_Activate()")
      s1.WriteLine("    Call CallDBLoadIfNeeded()")
      s1.WriteLine("    Call Me.LoadEntitiesInSheet()")
      s1.WriteLine("    lblLoginInfo.Caption = Globals.GetCurrentLoginStatus")
      s1.WriteLine("    currRow = ActiveCell.Row")
      s1.WriteLine("    Set currWks = Me")
      s1.WriteLine("    Set Globals.currObjCollection = Globals.omg.get{0:C}", pae.EntityCollectionName)
      s1.WriteLine("    Set Globals.currForm = {0:C}", pae.ExcelFormName)
      s1.WriteLine("    Set Globals.currEnt = Globals.Curr{0:C}", pae.EntityShortName)
      s1.WriteLine("    Call UIFunctions.PutAssociation(Me.CodeName, ""LastRefresh"", frmSettings.LastUpdateID)")
      s1.WriteLine("  End Sub")
      s1.WriteLine("")
      s1.WriteLine("  Public Sub ShowEntityDetails()")
      s1.WriteLine("    'sets the current Department and shows up the entity details form")
      s1.WriteLine("    Set Globals.Curr{0:C} = omg.get{1:C}.Item(CStr(Me.Cells(ActiveCell.Row, eColumnName.ID)))", pae.EntityShortName, pae.EntityCollectionName)
      s1.WriteLine("    Call {0:C}.Show()", pae.ExcelFormName)
      s1.WriteLine("  End Sub")
      s1.WriteLine("")
      s1.WriteLine("  Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)")
      s1.WriteLine("    Dim sTarget As String")
      s1.WriteLine("    sTarget = Target.Range.Text")
      s1.WriteLine("    If sTarget = ""New"" Then")
      s1.WriteLine("      Set Curr{0:C} = New {1:C}", pae.EntityShortName, pae.EntityName)
      s1.WriteLine("    End If")
      s1.WriteLine("    Call UIFunctions.wksFollowHyperlink(Target, Me, Globals.Curr{0:C})", pae.EntityShortName)
      s1.WriteLine("  End Sub")

      s1.WriteLine("'  '//////////////////////////////////////////////////////////////////////////////////////////////////")
      s1.WriteLine("'  'PlaceHolder Functions")
      s1.WriteLine()

      s1.WriteLine("  Public Sub LoadSheetintoEntity(ByVal feditmode As FormEditMode, ByVal n As Integer)")
      s1.WriteLine("    {0:C}.Loadorder = Me.Cells(n, eColumnName.Loadorder)  ", pae.EntityShortName)
      For Each paei As ProjectEntityItem In pae.mChildProjectEntityItems
        If Trim(paei.SheetDisplayOrder) <> "" Then 'only load those which are displayed on sheet
          s1.WriteLine("    Curr{0:C}.{1:C} = Me.Cells(n, eColumnName.{1:C}) ", pae.EntityShortName, paei.PropertyName)
        End If

      Next paei
      s1.WriteLine("  End Sub")
      s1.WriteLine()

      s1.WriteLine("  Public Sub LoadSingleEntityInSheet(ByVal n As Integer, ByVal ent As Object)")
      s1.WriteLine("    Me.Cells(n, eColumnName.Loadorder) = ent.Loadorder")
      For Each paei As ProjectEntityItem In pae.mChildProjectEntityItems
        If Trim(paei.SheetDisplayOrder) <> "" Then
          s1.WriteLine("    Me.Cells(n, eColumnName.{0:C}) = ent.{0:C}", paei.PropertyName)
        End If
        '    Me.Cells(n, eA_Column.ID) = ent.ID
        '    Me.Cells(n, eA_Column.EntityItem_1) = ent.EntityItem_1
        '    Me.Cells(n, eA_Column.EntityItem_2) = ent.EntityItem_2
      Next paei
      s1.WriteLine("  End Sub")
      s1.WriteLine()

      s1.WriteLine("  Sub LoadEntitiesInSheet()")
      s1.WriteLine("    Call Globals.CallDBLoadIfNeeded()")
      s1.WriteLine("    Call UIFunctions.ClearWorkSheet(Me)")
      s1.WriteLine("    nrow = NSTARTROW")
      s1.WriteLine("    For Each {0:C} In omg.get{1:C}", pae.EntityShortName, pae.EntityCollectionName)
      s1.WriteLine("      If {0:C}.ID = """" Then GoTo nextloop", pae.EntityShortName)
      s1.WriteLine("      Me.Hyperlinks.Add Anchor:=Me.Cells(nrow, eColumnName.HypEdit), Address:="""", SubAddress:= _")
      s1.WriteLine("                       Me.Cells(nrow, eColumnName.HypEdit).Address, TextToDisplay:=""Edit""")
      s1.WriteLine("      Me.Hyperlinks.Add Anchor:=Me.Cells(nrow, eColumnName.HypDetails), Address:="""", SubAddress:= _")
      s1.WriteLine("                        Me.Cells(nrow, eColumnName.HypDetails).Address, TextToDisplay:=""Details""")
      s1.WriteLine("")
      s1.WriteLine("      Call Me.LoadSingleEntityInSheet(nrow, {0:C})", pae.EntityShortName)
      s1.WriteLine("")
      s1.WriteLine("      nrow = nrow + 1")
      s1.WriteLine("nextloop:")
      s1.WriteLine("    Next {0:C}", pae.EntityShortName)
      s1.WriteLine("    'Another row for providing links to add a new entity")
      s1.WriteLine("    Me.Hyperlinks.Add Anchor:=Me.Cells(nrow, eColumnName.HypDetails), Address:="""", SubAddress:= _")
      s1.WriteLine("                      Me.Cells(nrow, eColumnName.HypDetails).Address, TextToDisplay:="" """)
      s1.WriteLine("    Me.Hyperlinks.Add Anchor:=Me.Cells(nrow, eColumnName.HypEdit), Address:="""", SubAddress:= _")
      s1.WriteLine("                      Me.Cells(nrow, eColumnName.HypEdit).Address, TextToDisplay:=""New""")
      s1.WriteLine("  End Sub")
      s1.WriteLine()

      s1.WriteLine("  Private Sub Worksheet_SelectionChange(ByVal Target As Range)")
      s1.WriteLine("    'Keeps the selected entity in the sheet in sync with the entity shown in the details form (if open)")
      s1.WriteLine("    Call Globals.CallDBLoadIfNeeded()")
      s1.WriteLine("    Set Globals.currForm = {0:C}", pae.ExcelFormName)
      s1.WriteLine("    Set Globals.currObjCollection = omg.get{0:C}", pae.EntityCollectionName)
      s1.WriteLine("    Call UIFunctions.WorksheetSelectionChange(Target, Me)")
      s1.WriteLine("    Set Globals.Curr{0:C} = Globals.currEnt", pae.EntityShortName)
      s1.WriteLine("  End Sub")

      s1.WriteLine("")
      s1.WriteLine("")

      s1.Close()
      fs1.Close()
      s1 = Nothing
      fs1 = Nothing

    Next pae
  End Function

End Class
