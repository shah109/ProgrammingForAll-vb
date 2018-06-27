Option Explicit On
Imports Microsoft.Win32
Imports PA_Framework_OM
Imports PA_Framework_OM.OMGlobals
Imports System.Drawing

Partial Public Class frmPAInstitute
  Dim cbobindingsource As New BindingSource

  Private Sub TabDepartments_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabDepartments.Enter
    nPrevUpdateID = omg.GetDBUpdates(cPAProjects, nPrevUpdateID)
    DepartmentBindingSource.DataSource = cDepartments
    DepartmentBindingSourceChanged()
    Me.btnUpdateDepartment.Enabled = False
    Me.btnCancelUpdateDepartment.Enabled = False
    DepartmentBindingSource.ResetBindings(False)
    Me.BackColor = Color.Yellow
  End Sub

  Private Sub DepartmentBindingSource_CurrentChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DepartmentBindingSource.CurrentChanged
    curreDpt_ = DepartmentBindingSource.Current
    DepartmentBindingSourceChanged()
  End Sub

  Sub DepartmentBindingSourceChanged()
    If curreDpt_ Is Nothing Then Exit Sub
    Dim crs As New Course
    nPrevUpdateID = omg.GetDBUpdates(cPAProjects, nPrevUpdateID)
    Dim nCount As Integer, strParDetails As String = "", strChldDetails = ""
    nCount = cPAProjects.GetEntityDependencies(curreDpt_, strParDetails, strChldDetails)
    Me.txtDepartmentDependencies.Text = CStr(nCount)
    If nCount <> 0 Then
      Me.btnDeleteDepartment.Enabled = False
    Else
      Me.btnDeleteDepartment.Enabled = True
    End If

    Me.txtDepartmentParDeps.Text = strParDetails
    Me.txtDepartmentChldDeps.Text = strChldDetails

    If curreDpt_.IsDirty = True Then
      Me.btnUpdateDepartment.Enabled = True
      Me.btnCancelUpdateDepartment.Enabled = True
    Else
      Me.btnUpdateDepartment.Enabled = False
      Me.btnCancelUpdateDepartment.Enabled = False
    End If
    'Department parent Courses
    DepartmentParCourseBindingSource.DataSource = curreDpt_.ParentEntities(crs)
    Me.dgvDepartmentParCourse.DataSource = DepartmentParCourseBindingSource
    DepartmentParCourseBindingSource.ResetBindings(True)

    'Department available parent courses
    Call omg.FillAvailableParentEntities(cPAProjects, crs, curreDpt_, "Departments")
    AvDepartmentParCourseBindingSource.DataSource = curreDpt_.AvailableParentEntities(crs)
    Me.dgvAvDepartmentParCourse.DataSource = AvDepartmentParCourseBindingSource
    AvDepartmentParCourseBindingSource.ResetBindings(True)

  End Sub

  Private Sub btnAddDepartment_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNewDepartment.Click
    Dim np = New Department
    np.EntityItem_1 = "new"
    Call omg.DBUpdate("Add", cDepartments, np)
    nPrevUpdateID = omg.GetSetting("LastUpdateID")
    DepartmentBindingSource.ResetBindings(True)
    DepartmentBindingSource.MoveLast()

  End Sub

  Private Sub btnDeleteDepartment_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDeleteDepartment.Click
    Dim nCount As Integer, strParDetails As String = "", strChldDetails As String = ""
    curreDpt_ = DepartmentBindingSource.Current
    nCount = cPAProjects.GetEntityDependencies(curreDpt_, strParDetails, strChldDetails)
    If nCount = 0 Then
      omg.DBUpdate("Delete", cDepartments, curreDpt_)
      nPrevUpdateID = omg.GetSetting("LastUpdateID")
    Else
      MsgBox("Can not delete because of dependencies ")
    End If
    DepartmentBindingSource.ResetBindings(True)
  End Sub

  Private Sub btnUpdateDepartment_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdateDepartment.Click
    Dim bResult As Boolean
    Me.dgvDepartments.EndEdit()

    curreDpt_ = DepartmentBindingSource.Current
    bResult = omg.DBUpdate("Update", cDepartments, curreDpt_)
    nPrevUpdateID = omg.GetSetting("LastUpdateID")
    If bResult = True Then
      Me.btnCancelUpdateDepartment.Enabled = False
      Me.btnUpdateDepartment.Enabled = False
      Me.LockDGVRows(dgvDepartments, False)
    End If
  End Sub

  Private Sub btnRemDepartmentParCourse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemDepartmentParCourse.Click
    Dim crs As Course
    curreDpt_ = DepartmentBindingSource.Current
    crs = DepartmentParCourseBindingSource.Current
    If crs Is Nothing Then Exit Sub
    omg.RemoveChildEntity(cPAProjects, crs, curreDpt_, "Departments")
    nPrevUpdateID = omg.GetSetting("LastUpdateID")
    'refresh dgv's
    Call Me.DepartmentBindingSourceChanged()
  End Sub

  Private Sub btbnAddAvDepartmentParCourse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btbnAddAvDepartmentParCourse.Click
    Dim crs As Course
    curreDpt_ = DepartmentBindingSource.Current()
    crs = AvDepartmentParCourseBindingSource.Current()
    If crs Is Nothing Then Exit Sub
    omg.AddChildEntity(cPAProjects, crs, curreDpt_, "Departments")
    nPrevUpdateID = omg.GetSetting("LastUpdateID")
    'refresh dgv's
    Call Me.DepartmentBindingSourceChanged()
  End Sub

  Private Sub dgvDepartments_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvDepartments.CellBeginEdit

    Me.sCurrentCellText = dgvDepartments.CurrentCell.Value
    'Me.dgvDepartments.CurrentCell.Style.BackColor = Drawing.Color.Aqua

    Me.LockDGVRows(dgvDepartments, True)
    Me.dgvDepartments.Rows(Me.dgvDepartments.CurrentRow.Index).ReadOnly = False
  End Sub

  Private Sub dgvDepartments_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvDepartments.CellEndEdit
    'MessageBox.Show("CellEndEdit")
    'Dim cst As DataGridViewCellStyle
    'cst = Me.dgvDepartments.CurrentCell.Style
    'If dgvDepartments.CurrentCell.EditedFormattedValue = Me.sCurrentCellText And cst.BackColor = Color.Aqua Then
    '  cst.BackColor = Drawing.Color.Empty
    'End If


  End Sub

  Private Sub dgvDepartments_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvDepartments.CellValueChanged
    'MessageBox.Show(e.RowIndex)
    '  MessageBox.Show("CellValueChanged")
    '  '  'Me.dgvDepartments.CurrentCell.Style.BackColor = Drawing.Color.Aqua
    '  '  'Dim CellStyle = New DataGridViewCellStyle()
    '  '  'CellStyle.BackColor = Drawing.Color.Aqua
    '  '  ''dgvDepartments.Rows[e.RowIndex].Cells[e.ColumnIndex].Style = CellStyle
    '  '  'dgvDepartments.CurrentCell.Style = CellStyle
    '  '  'dgvDepartments.CurrentCell.Style.ApplyStyle(CellStyle)
    '  '  'dgvDepartments.InvalidateCell(dgvDepartments.CurrentCell)
    '  If Me.bEditBegines = False Then Exit Sub
    '  '  If txtCurrentCellText <> dgvDepartments.CurrentCell.Value Then
    '  Me.dgvDepartments.CurrentCell.Style.BackColor = Drawing.Color.Aqua
    '  '  End If
    'End Sub

    'Private Sub dgvDepartments_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvDepartments.CurrentCellChanged
    '  MessageBox.Show("CurrentCellChanged")
    '  If Me.bEditBegines = False Then Exit Sub
    '  If txtCurrentCellText <> dgvDepartments.CurrentCell.Value Then
    '    Me.dgvDepartments.CurrentCell.Style.BackColor = Drawing.Color.Aqua
    '  End If
    '  'Dim oldrowind As Integer
    '  'If oldrowind Then
    '  'If Me.dgvDepartments.CurrentCell Is Nothing Then Exit Sub
    '  'Me.dgvDepartments.CurrentCell.Style.BackColor = Drawing.Color.Aqua
  End Sub

  Private Sub dgvDepartments_CurrentCellDirtyStateChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvDepartments.CurrentCellDirtyStateChanged
    'MessageBox.Show("CurrentCellDirtyStateChanged")
    'If Me.dgvDepartments.CurrentCell.IsInEditMode Then
    '  MessageBox.Show("IsInEditMode")
    '  Me.dgvDepartments.CurrentCell.Style.BackColor = Drawing.Color.Aqua
    'End If
    Me.dgvDepartments.CurrentCell.Style.BackColor = Drawing.Color.Aqua
    'If Me.dgvDepartments.IsCurrentCellDirty Then
    'Me.dgvDepartments.CommitEdit(DataGridViewDataErrorContexts.Commit)
    'End If

    Me.btnUpdateDepartment.Enabled = True
    Me.btnCancelUpdateDepartment.Enabled = True
    'Me.dgvDepartments.InvalidateCell(Me.dgvDepartments.CurrentCell)

    'Dim CellStyle = New DataGridViewCellStyle()
    'CellStyle.BackColor = Drawing.Color.Aqua
    'dgvDepartments.CurrentCell.Style.ApplyStyle(CellStyle)

    'CellStyle.ForeColor = Drawing.Color.Red
    ''dgvDepartments.Rows[e.RowIndex].Cells[e.ColumnIndex].Style = CellStyle
    'dgvDepartments.CurrentCell.Style = CellStyle
    'dgvDepartments.InvalidateRow(dgvDepartments.CurrentCellAddress.Y)
    'dgvDepartments.CommitEdit(DataGridViewDataErrorContexts.Parsing)
    'e()
    'If Me.dgvDepartments.IsCurrentCellDirty = True Then
    'Dim cll As DataGridViewCell
    'cll = Me.dgvDepartments.CurrentCell
    'cll.Style.BackColor = Drawing.Color.Aqua
    'If Me.dgvDepartments.CurrentCell.IsInEditMode = True Then
    'Me.dgvDepartments.CurrentCell.Style.BackColor = Drawing.Color.Aqua
    'Me.dgvDepartments.CurrentCell.Style.BackColor = Drawing.Color.Aqua
    'Me.dgvDepartments.InvalidateRow(cll.RowIndex)
    'End If
    'End If
    'Me.dgvDepartments.CurrentCell

    '''''Me.dgvDepartments.CurrentCell.Style.BackColor = Drawing.Color.Aqua
    '''''Me.dgvDepartments.InvalidateCell(dgvDepartments.CurrentCell)
  End Sub

  Private Sub btnCancelUpdateDepartment_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancelUpdateDepartment.Click
    Dim bResult As Boolean
    Me.dgvDepartments.EndEdit()

    curreDpt_ = DepartmentBindingSource.Current
    bResult = omg.DBUpdate("Load", cDepartments, curreDpt_)
    nPrevUpdateID = omg.GetSetting("LastUpdateID")
    If bResult = True Then
      Me.btnCancelUpdateDepartment.Enabled = False
      Me.btnUpdateDepartment.Enabled = False
      Me.LockDGVRows(dgvDepartments, False)
    End If
  End Sub

  Private Sub LockDGVRows(ByRef dgv As DataGridView, ByVal bLock As Boolean)
    For RCnt As Integer = 0 To dgv.Rows.Count - 1
      dgv.Rows(RCnt).ReadOnly = bLock
    Next RCnt
    If bLock = False Then 'if clearing the lock
      For CellCnt As Integer = 0 To dgv.CurrentRow.Cells.Count - 1
        dgv.CurrentRow.Cells(CellCnt).Style.BackColor = Color.Empty
      Next
      curreDpt_.IsDirty = False
    ElseIf bLock = True Then
      curreDpt_.IsDirty = True

    End If
  End Sub

  Private Sub TabDepartments_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabDepartments.Leave
    Me.BackColor = Color.White
  End Sub
End Class



