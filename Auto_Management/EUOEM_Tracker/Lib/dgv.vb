Imports System.IO

Module dgv
    Public Sub init(ByRef drvcontrol As DataGridView, ByRef bs As BindingSource)
        drvcontrol.DataSource = bs
        drvcontrol.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        drvcontrol.CellBorderStyle = DataGridViewCellBorderStyle.None
        drvcontrol.Columns.Clear()
        drvcontrol.Columns.Add(New DataGridViewTextBoxColumn With {.Width = 100, .Name = "Project", .DataPropertyName = "Project"})
        'drvcontrol.Columns.Add(New DataGridViewTextBoxColumn With {.Width = 60, .Name = "Model", .DataPropertyName = "Model"})
        'drvcontrol.Columns.Add(New DataGridViewTextBoxColumn With {.Width = 70, .Name = "Release", .DataPropertyName = "Release"})
        drvcontrol.Columns.Add(New DataGridViewTextBoxColumn With {.Width = 120, .Name = "Task_ID", .DataPropertyName = "Task_ID"})
        drvcontrol.Columns.Add(New DataGridViewTextBoxColumn With {.Width = 180, .Name = "Module", .DataPropertyName = "Module"})
        drvcontrol.Columns.Add(New DataGridViewTextBoxColumn With {.Width = 50, .Name = "Eloc", .DataPropertyName = "Eloc"})
        Dim cmb As New DataGridViewComboBoxColumn()
        cmb.Width = 150
        cmb.Name = "Status"
        cmb.Items.Add("Started")
        cmb.Items.Add("Requirement Analyzing")
        cmb.Items.Add("OPL sending")
        cmb.Items.Add("Testcase Designing")
        cmb.Items.Add("Reviewing")
        cmb.Items.Add("Review DONE")
        cmb.Items.Add("PDC waiting")
        cmb.Items.Add("Delivered")
        cmb.Items.Add("")
        cmb.DataPropertyName = "Status"
        cmb.DataSource = MainF.ds.Tables("Status")
        cmb.ValueMember = "Status"
        cmb.DisplayMember = cmb.ValueMember
        cmb.FlatStyle = FlatStyle.Flat
        drvcontrol.Columns.Add(cmb)
        drvcontrol.Columns.Add(New DataGridViewTextBoxColumn With {.Width = 230, .Name = "Remark", .DataPropertyName = "Remark"})

        drvcontrol.Columns.Add(New DataGridViewTextBoxColumn With {.Width = 100, .Name = "Date", .DataPropertyName = "Date_Task"})
    End Sub


    Public Sub highlight(ByRef dgvControl As MetroFramework.Controls.MetroGrid, ByVal record As String)
        Dim FindRow As DataGridViewRow = dgvControl.Rows.Cast(Of DataGridViewRow)().Where(Function(r) r.Cells("Task_ID").Value.ToString().Equals(record)).FirstOrDefault
        If Not FindRow Is Nothing Then
            Dim FindRowIndex As String = FindRow.Index
            dgvControl.CurrentCell = dgvControl.Rows(FindRowIndex).Cells("Task_ID")
            dgvControl.CurrentCell.Selected = True
        End If
    End Sub

    Public Function add(ByRef txt As MetroFramework.Controls.MetroTextBox, ByRef drv As MetroFramework.Controls.MetroGrid, ByRef dt As DataTable, ByVal isRev As Boolean, ByVal isMul As Boolean, ByRef errorMsg As String) As Boolean
        Dim errorFlag As Boolean = False
        If My.Settings.LastFullPath <> "" And System.IO.File.Exists(My.Settings.LastFullPath) = True And Path.GetExtension(My.Settings.LastFullPath) = ".accdb" Then
            Dim newtask As String = Trim(txt.Text) : txt.Text = ""
            If newtask <> "" Then
                Dim findrow As DataRow() = dt.[Select]("Task_ID = '" & newtask & "'")

                If findrow.Length = 0 Then
                    Dim newrow As DataRow = ace.add(dt)
                    If isRev = True Then
                        newrow.Item("isRev") = "R"
                        newrow.Item("Status") = "Reviewing"
                    Else
                        newrow.Item("Status") = "Started"
                    End If
                    If isMul = True Then newrow.Item("isMul") = "M"
                    newrow.Item("Task_ID") = newtask
                    newrow.Item("Date_Task") = CStr(Now().Date)
                Else
                    errorFlag = True : errorMsg = "Task is existed"
                End If
                highlight(drv, newtask)
            Else
                errorFlag = True : errorMsg = "Textbox is empty"
            End If
        End If

        Return errorFlag
    End Function

    Public Function find(ByRef txt As MetroFramework.Controls.MetroTextBox, ByRef drv As MetroFramework.Controls.MetroGrid, ByRef dt As DataTable, ByRef errorMsg As String) As Boolean

        Dim errorFlag As Boolean = False
        If My.Settings.LastFullPath <> "" And System.IO.File.Exists(My.Settings.LastFullPath) = True And Path.GetExtension(My.Settings.LastFullPath) = ".accdb" Then
            Dim findtask As String = Trim(txt.Text) : txt.Text = ""
            If findtask <> "" Then
                Dim foundtask As DataRow() = dt.[Select]("Task_ID = '" & findtask & "'")
                If foundtask.Length = 0 Then
                    errorFlag = True : errorMsg = "Task is not existed"
                Else
                    highlight(drv, findtask)
                End If
            Else
                errorFlag = True : errorMsg = "Textbox is empty"
            End If
        End If

        Return errorFlag
    End Function

    Public Function multiselect(ByVal dgvcontrol As MetroFramework.Controls.MetroGrid, ByRef selrow() As DataRow) As Integer
        Dim SelRowIndex As Integer = 0
        For Each EachSelRow As DataGridViewRow In dgvcontrol.SelectedRows
            ReDim Preserve selrow(dgvcontrol.SelectedRows.Count - 1)
            selrow(SelRowIndex) = CType(EachSelRow.DataBoundItem, DataRowView).Row
            SelRowIndex = SelRowIndex + 1
        Next
        Return SelRowIndex
    End Function

    Public Function removeA(ByRef dgv As DataGridView, ByRef dt As DataTable, ByRef errorMsg As String) As Boolean
        Dim errorFlag As Boolean = False
        If My.Settings.LastFullPath <> "" And System.IO.File.Exists(My.Settings.LastFullPath) = True And Path.GetExtension(My.Settings.LastFullPath) = ".accdb" Then
            Dim selrow() As DataRow = Nothing
            Dim NumSelRow As Integer = multiselect(dgv, selrow)
            If selrow IsNot Nothing Then
                For Each row As DataRow In selrow
                    Dim taskid As String = row.Item("Task_ID")
                    Dim findtask As DataRow() = dt.[Select]("Task_ID = '" & taskid & "'")
                    If findtask.Length <> 0 Then
                        For Each r In findtask
                            ace.remove(dt, r)
                        Next
                    End If
                Next
            End If
        End If
        Return errorFlag
    End Function

    Public Function remove(ByRef dgv As DataGridView, ByRef dt As DataTable, ByRef errorMsg As String) As Boolean
        Dim errorFlag As Boolean = False
        If My.Settings.LastFullPath <> "" And System.IO.File.Exists(My.Settings.LastFullPath) = True And Path.GetExtension(My.Settings.LastFullPath) = ".accdb" Then
            Dim selrow() As DataRow = Nothing
            Dim NumSelRow As Integer = multiselect(dgv, selrow)
            If selrow IsNot Nothing Then
                For Each row As DataRow In selrow
                    ace.remove(dt, row)
                Next
            End If
        End If
        Return errorFlag
    End Function

    Public Function move(ByRef drv As DataGridView, ByRef dt As DataTable, ByVal isStored As Boolean, ByRef errorMsg As String) As Boolean
        Dim errorFlag As Boolean = False
        If My.Settings.LastFullPath <> "" And System.IO.File.Exists(My.Settings.LastFullPath) = True And Path.GetExtension(My.Settings.LastFullPath) = ".accdb" Then
            Dim selrow() As DataRow = Nothing
            Dim numrow = multiselect(drv, selrow)
            Dim storedchar As String = Nothing
            If isStored = True Then
                storedchar = "S"
            End If
            If selrow Is Nothing Then
                MainF.OutputTextBox.Text = "No row selected"
                Return Nothing
                Exit Function
            End If
            For Each row As DataRow In selrow
                Dim taskid As String = row.Item("Task_ID")
                Dim findtask As DataRow() = Nothing
                findtask = dt.[Select]("Task_ID = '" & taskid & "'")
                If findtask.Length <> 0 Then
                    For Each r In findtask
                        r.Item("isStored") = storedchar
                    Next
                End If
            Next
        End If
        Return errorFlag
    End Function

End Module
