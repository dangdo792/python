Imports System.ComponentModel
Imports System.IO
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports Scripting

Public Class ConfigF
    Inherits MetroFramework.Forms.MetroForm

    Private dsConfig As New DataSet
    Private bs As New BindingSource

    Private Sub ConfigF_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        MetroStyleManager1.Style = MetroFramework.MetroColorStyle.Teal
        Me.Style = MetroFramework.MetroColorStyle.Teal
        TabControl1.Appearance = TabAppearance.FlatButtons
        TabControl1.ItemSize = New Size(0, 1)
        TabControl1.SizeMode = TabSizeMode.Fixed
        MaximizeBox = False : MinimizeBox = False : ControlBox = False

        Dim parentnode As TreeNode = TreeView1.Nodes.Add("Configuration")
        Dim tn1 = parentnode.Nodes.Add("User Config")
        Dim tn2 = parentnode.Nodes.Add("Project Config")

        parentnode.ExpandAll()

        If My.Settings.LastFullPath <> "" And IO.File.Exists(My.Settings.LastFullPath) = True And Path.GetExtension(My.Settings.LastFullPath) = ".accdb" Then
            dsConfig = MainF.ds
            bs.DataSource = New DataView(dsConfig.Tables("User_Config"))
            tex.init(GroupBox1, bs)
            tex.init(GroupBox2, bs)
            TreeView1.SelectedNode = tn1
            If MainF.IsProject_ModelToConfig = True Then
                TreeView1.SelectedNode = tn2
                ProModTextBox.Text = MainF.Project_ModelToConfig
            End If
            TreeView1.Select()
        End If
    End Sub

    Private JsonFileName As String

    Private Sub LoadJsonFile()
        Dim userinfo As TContainer.UserInfo
        userinfo = New TContainer.UserInfo(dsConfig.Tables("User_Config").Rows(0))
        Dim JsonFilePath As New Information.CheckNull("Json File Path", userinfo.JsonPath)
        Dim LoadJson As New WriteJson.LoadJson(JsonFilePath, JsonFileName)
        If JsonFileName = "Project_Config.json" Then
            Grid1.DataSource = LoadJson.execute()
            Grid1.Columns.Clear()
            Grid1.Columns.Add(New DataGridViewTextBoxColumn With {.Width = 140, .HeaderText = "Project", .Name = "Project_Model", .DataPropertyName = "Project_Model"})
            Grid1.Columns.Add(New DataGridViewTextBoxColumn With {.Width = 150, .HeaderText = "Coordinator", .Name = "Task_Coor", .DataPropertyName = "Task_Coor"})
            Grid1.Columns.Add(New DataGridViewTextBoxColumn With {.Width = 300, .HeaderText = "Result Path", .Name = "Result_Path", .DataPropertyName = "Result_Path"})
            'Grid1.Columns.Add(New DataGridViewTextBoxColumn With {.Width = 120, .HeaderText = "System Type", .Name = "System_Type", .DataPropertyName = "System_Type"})
        End If
        If Not String.IsNullOrEmpty(LoadJson.ErrorMsg) Then
            Dim result As MsgBoxResult = Nothing
            result = MsgBox(LoadJson.ErrorMsg & vbNewLine & vbNewLine & "Do you want create new file", vbYesNo)
            If result = MsgBoxResult.Yes Then
                Dim fso As Object
                fso = CreateObject("scripting.filesystemobject")
                Try
                    Dim Fileout As Object
                    Fileout = fso.CreateTextFile(LoadJson.jsonFilePath, True, True)
                    Dim content1 As String = Nothing
                    If JsonFileName = "Project_Config.json" Then : content1 = "{" & Chr(34) & "fields" & Chr(34) & ":[]}"
                    ElseIf JsonFileName = "Review_Config.json" Then : content1 = "{" & Chr(34) & "fieldsReview" & Chr(34) & ":[]}"
                    End If
                    If content1 IsNot Nothing Then
                        Fileout.Write(content1)
                        Fileout.Close()
                        Fileout = Nothing
                        fso = Nothing
                        GC.Collect()
                    End If

                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
            Else
                Exit Sub
            End If

        End If
    End Sub

    Private Sub AddProBtn_Click(sender As Object, e As EventArgs) Handles AddProBtn.Click
        Dim userinfo As TContainer.UserInfo
        userinfo = New TContainer.UserInfo(dsConfig.Tables("User_Config").Rows(0))
        Dim JsonFilePath As New Information.CheckNull("Json File Path", userinfo.JsonPath)
        If String.IsNullOrEmpty(ProModTextBox.Text) Then
            MsgBox("Project_Model textbox is empty.")
            Exit Sub
        Else
            ProModTextBox.Text = ProModTextBox.Text.Trim
        End If
        If Not String.IsNullOrEmpty(ProCoorTextBox.Text) Then ProCoorTextBox.Text = ProCoorTextBox.Text.Trim
        If Not String.IsNullOrEmpty(ProResultPathTextBox.Text) Then ProResultPathTextBox.Text = ProResultPathTextBox.Text.Trim

        Dim AddJsonFile As New WriteJson.AddJson(JsonFilePath,
                                                 JsonFileName,
                                                 ProModTextBox.Text,
                                                 ProCoorTextBox.Text,
                                                 ProResultPathTextBox.Text)
        AddJsonFile.execute()
        ProModTextBox.Text = ""
        ProCoorTextBox.Text = ""
        ProResultPathTextBox.Text = ""
        LoadJsonFile()

        If Grid1.RowCount <> 0 Then
            Grid1.CurrentCell = Grid1.Rows(Grid1.RowCount - 1).Cells("Project_Model")
            Grid1.CurrentCell.Selected = True
        End If

    End Sub

    Private Sub RemoveProBtn_Click(sender As Object, e As EventArgs) Handles RemoveProBtn.Click
        Dim userinfo As TContainer.UserInfo
        userinfo = New TContainer.UserInfo(dsConfig.Tables("User_Config").Rows(0))
        Dim JsonFilePath As New Information.CheckNull("Json File Path", userinfo.JsonPath)
        Dim RemoveJsonFile As New WriteJson.RemoveJson(JsonFilePath, JsonFileName, Grid1)
        RemoveJsonFile.execute()
        LoadJsonFile()
    End Sub

    Private Sub ConfigF_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Enter) Then : savebtn.PerformClick()
        ElseIf (e.KeyCode = Keys.Escape) Then : cancelbtn.PerformClick()
        End If
    End Sub

    Private Sub TreeView1_AfterSelect(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles TreeView1.AfterSelect
        If TreeView1.SelectedNode.Text = "User Config" Then
            Me.TabControl1.SelectedTab = TabPage1
            Me.TabPage1.Select()
        End If
        If TreeView1.SelectedNode.Text = "Project Config" Then
            Me.TabControl1.SelectedTab = TabPage2
            Me.TabPage2.Select()
            JsonFileName = "Project_Config.json"
            LoadJsonFile()
        End If
    End Sub

    Private Sub savebtn_Click(sender As Object, e As EventArgs) Handles savebtn.Click
        If My.Settings.LastFullPath <> "" And System.IO.File.Exists(My.Settings.LastFullPath) = True And Path.GetExtension(My.Settings.LastFullPath) = ".accdb" Then
            MainF.ds = dsConfig
            bs.EndEdit() : Grid1.EndEdit()
            save(MainF.ds, MainF.dbpath)
            MainF.UpJsonFile.Visible = False
            Close() : Dispose()
        End If

    End Sub

    Private Sub cancelbtn_Click(sender As Object, e As EventArgs) Handles cancelbtn.Click
        MainF.UpJsonFile.Visible = False
        Close() : Dispose()
    End Sub

    Private Sub FindProBtn_Click(sender As Object, e As EventArgs) Handles FindProBtn.Click
        Dim FindRow As DataGridViewRow = Grid1.Rows.Cast(Of DataGridViewRow)().Where(Function(r) LCase(r.Cells("Project_Model").Value.ToString()).Equals(LCase(ProModTextBox.Text))).FirstOrDefault
        If Not FindRow Is Nothing Then
            Dim FindRowIndex As String = FindRow.Index
            Grid1.CurrentCell = Grid1.Rows(FindRowIndex).Cells("Project_Model")
            Grid1.CurrentCell.Selected = True
        Else
            MsgBox("Not Found")
        End If
    End Sub
End Class


