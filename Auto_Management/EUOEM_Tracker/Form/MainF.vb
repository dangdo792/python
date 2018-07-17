Imports System.Reflection

Imports System.ComponentModel
Imports System.IO
Imports System.Threading
Imports Microsoft.Office.Interop
Imports Newtonsoft.Json.Linq
Imports System.Data.OleDb

Public Class MainF
    Inherits MetroFramework.Forms.MetroForm

    Public dbpath As String
    Public ds As New DataSet
    Private bs As New BindingSource
    Private bs2 As New BindingSource
    Private bs3 As New BindingSource

    '*********************************-Form Event Handler-*********************************************

    Private Sub init()
        bs.DataSource = New DataView(ds.Tables("Task"))
        bs2.DataSource = New DataView(ds.Tables("Task"))
        bs3.DataSource = New DataView(ds.Tables("Task"))

        TabControl1.SelectedTab = TabControl1.TabPages.Item(1)
        TabControl1.SelectedTab = TabControl1.TabPages.Item(2)
        TabControl1.SelectedTab = TabControl1.TabPages.Item(0)

        bs.Filter = " isRev is Null AND isMul is Null AND isStored is Null" : bs.Sort = "ID DESC" : dgv.init(Grid1, bs)
        bs2.Filter = "isRev = 'R' AND isMul is Null AND isStored is Null" : bs2.Sort = "ID DESC" : dgv.init(Grid2, bs2)
        bs3.Filter = "isMul is Null AND isStored = 'S'" : bs3.Sort = "ID DESC" : dgv.init(Grid3, bs3)

        tex.init(Panel1, bs)

        AddHandler ds.Tables("Task").ColumnChanging, New DataColumnChangeEventHandler(AddressOf Column_Changing)
        AddHandler ds.Tables("Task").RowDeleted, New DataRowChangeEventHandler(AddressOf Row_Deleted)

    End Sub

    Private Sub Column_Changing(sender As Object, e As DataColumnChangeEventArgs)
        SaveToolStrip.Enabled = True
    End Sub

    Private Sub Row_Deleted(ByVal sender As Object, ByVal e As DataRowChangeEventArgs)
        SaveToolStrip.Enabled = True
    End Sub

    Private Sub DataGridView1_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles Grid1.DataError

    End Sub

    Dim rs As New Resizer

    Private Sub MainF_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        rs.FindAllControls(Me)
        If My.Settings.LastFullPath <> "" And System.IO.File.Exists(My.Settings.LastFullPath) = True And Path.GetExtension(My.Settings.LastFullPath) = ".accdb" Then
            dbpath = My.Settings.LastFullPath
            ace.load(dbpath, ds)
            init()
        Else
            OutputTextBox.Text = "There is no database file"
        End If
    End Sub

    Private Sub Form1_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        rs.ResizeAllControls(Me)
    End Sub


    Private Sub MainF_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        If My.Settings.LastFullPath <> "" And System.IO.File.Exists(My.Settings.LastFullPath) = True And Path.GetExtension(My.Settings.LastFullPath) = ".accdb" Then
            If ds.HasChanges = True Then
                Dim result As Integer = MessageBox.Show("Would you like to save it?", "caption", MessageBoxButtons.YesNoCancel)
                Select Case result
                    Case DialogResult.Yes : SaveToolStrip.PerformClick()
                    Case DialogResult.No : Exit Sub
                    Case DialogResult.Cancel : e.Cancel = True
                End Select
            End If
        End If
    End Sub

    Private Sub MainF_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.S AndAlso e.Modifiers = Keys.Control) Then : SaveToolStrip.PerformClick()
        ElseIf (e.KeyCode = Keys.O AndAlso e.Modifiers = Keys.Control) Then : OpenToolStrip.PerformClick()
        ElseIf (e.KeyCode = Keys.F AndAlso e.Modifiers = Keys.Control) Then : ConfigToolStrip.PerformClick()
        ElseIf (e.KeyCode = Keys.Enter) Then
            If TabControl1.SelectedTab.Name = "TabPage1" Then : addbtn.PerformClick()
            ElseIf TabControl1.SelectedTab.Name = "TabPage2" Then : addRbtn.PerformClick()
            ElseIf TabControl1.SelectedTab.Name = "TabPage3" Then : findHbtn.PerformClick()
            End If
        End If
    End Sub

    '*********************************-ToolStrip Event Handler-*********************************************
    Private Sub CreateNewDatabaseFileBtn_Click(sender As Object, e As EventArgs) Handles CreateNewDatabaseFileBtn.Click
        CreateNewDatabaseFileBtn.Enabled = False
        Dim appPath As String = Application.StartupPath() & "\YourDatabase.accdb"

        Try
            ' Create a Catalog object
            Dim ct As New ADOX.Catalog()
            ct.Create("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & appPath)

            Dim oConn As OleDbConnection
            Dim oComm, oComm2, oComm3 As OleDbCommand
            Dim oConnect, oQuery, oQuery2, oQuery3 As String

            oConnect = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & appPath

            oQuery = "CREATE TABLE Task ( ID NUMBER," &
                                         "Status TEXT(255)," &
                                          "isRev TEXT(255)," &
                                          "isMul TEXT(255) ," &
                                          "isStored TEXT(255) ," &
                                          "Remark TEXT(255) ," &
                                          "Project TEXT(255)," &
                                          "Model TEXT(255) ," &
                                          "Release TEXT(255) ," &
                                          "Task_ID TEXT(255) ," &
                                          "Eloc TEXT(255) ," &
                                          "[Module] TEXT(255) ," &
                                          "Revision TEXT(255)," &
                                          "Sandbox TEXT(255) ," &
                                          "Feature_Branch TEXT(255) ," &
                                          "M_Path TEXT(255) ," &
                                          "M_Owner TEXT(255)," &
                                          "Old_Task TEXT(255) ," &
                                          "Reviewer TEXT(255) ," &
                                          "Defect_ID TEXT(255) ," &
                                          "Sub_Rev TEXT(255) ," &
                                          "ILM_Link TEXT(255) ," &
                                          "OPL_Link TEXT(255) ," &
                                          "RS TEXT(255) ," &
                                          "RS_BL TEXT(255) ," &
                                          "TS TEXT(255) ," &
                                          "TS_BL TEXT(255) ," &
                                          "SD TEXT(255) ," &
                                          "SD_BL TEXT(255) ," &
                                          "Statement TEXT(255) ," &
                                          "Decisions TEXT(255) ," &
                                          "Cover_Path TEXT(255) ," &
                                          "Result_Path TEXT(255) ," &
                                          "Package_ID TEXT(255) ," &
                                          "Date_Task TEXT(255) ," &
                                          "PRIMARY KEY(ID) )"
            oQuery2 = "CREATE TABLE User_Config ( ID COUNTER," &
                                         "MyName TEXT(255)," &
                                          "wfd TEXT(255)," &
                                          "sfd TEXT(255) ," &
                                          "rfd TEXT(255) ," &
                                          "tfd TEXT(255) ," &
                                          "explorer TEXT(255)," &
                                          "mks_checker TEXT(255) ," &
                                          "teamlead TEXT(255) ," &
                                          "myPM TEXT(255) ," &
                                          "CMacro TEXT(255) ," &
                                          "CppMacro TEXT(255)," &
                                          "SysBPlusMacro TEXT(255) ," &
                                          "CQUser TEXT(255) ," &
                                          "CQPassword TEXT(255) ," &
                                          "Style TEXT(255)," &
                                          "FilesTemplateDir TEXT(255) ," &
                                          "JsonPath TEXT(255) ," &
                                          "PRIMARY KEY(ID) )"

            oQuery3 = "INSERT INTO User_Config([explorer],[teamlead],[myPM],[FilesTemplateDir],[JsonPath])VALUES(@explorer,@teamlead,@myPM,@FilesTemplateDir,@JsonPath)"


            ' Instantiate the connectors
            oConn = New OleDbConnection(oConnect)
            oComm = New OleDbCommand(oQuery, oConn)
            oComm2 = New OleDbCommand(oQuery2, oConn)
            oComm3 = New OleDbCommand(oQuery3, oConn)

            oComm3.Parameters.AddWithValue("@explorer", "C:\Windows\explorer.exe")
            oComm3.Parameters.AddWithValue("@teamlead", "Nguyen Thi Thanh")
            oComm3.Parameters.AddWithValue("@myPM", "Hoang Xuan Thanh Khiet")
            oComm3.Parameters.AddWithValue("@FilesTemplateDir", "\\bosch.com\dfsRB\DfsVN\LOC\Hc\RBVH\20_ESS\04_Projects\10_CC_DA\EI-300010\EUOEM\Template_Document")
            oComm3.Parameters.AddWithValue("@JsonPath", "\\bosch.com\dfsRB\DfsVN\LOC\Hc\RBVH\20_ESS\04_Projects\10_CC_DA\EI-300010\EUOEM\Shared_Folder")

            oConn.Open()
            oComm.ExecuteNonQuery()
            oComm2.ExecuteNonQuery()
            oComm3.ExecuteNonQuery()
            oConn.Close()

            dbpath = appPath
            My.Settings.LastFullPath = dbpath
            My.Settings.Save()
            ace.load(dbpath, ds)

            init()
            OutputTextBox.Text = "Create database file successful"
        Catch exp As Exception
            OutputTextBox.Text = exp.Message.ToString()
        End Try
        CreateNewDatabaseFileBtn.Enabled = True
    End Sub

    Private Sub OpenToolStrip_Click(sender As Object, e As EventArgs) Handles OpenToolStrip.Click
        OpenToolStrip.Enabled = False
        Dim DlgOpenFile As New OpenFileDialog()
        DlgOpenFile.Filter = "Access files (*.accdb)|*.accdb|All files (*.*)|*.*"
        If DlgOpenFile.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            dbpath = DlgOpenFile.FileName
            ace.load(dbpath, ds)
            My.Settings.LastFullPath = DlgOpenFile.FileName
            My.Settings.Save()
            init()
            OutputTextBox.Text = "Load database file is succesful"
        End If
        OpenToolStrip.Enabled = True
    End Sub

    Private Sub SaveToolStrip_Click(sender As Object, e As EventArgs) Handles SaveToolStrip.Click
        SaveToolStrip.Enabled = False
        If Not dbpath Is Nothing Then
            bsEnd() : gridEnd()
            ace.save(ds, dbpath)
        End If
    End Sub

    Private Sub ConfigToolStrip_Click(sender As Object, e As EventArgs) Handles ConfigToolStrip.Click
        Dim configf As New ConfigF
        ConfigToolStrip.Enabled = False
        'IsProject_ModelToConfig = False
        'IsReviewerToConfig = False
        configf.ShowDialog()
        ConfigToolStrip.Enabled = True
    End Sub

    '*********************************-Datagridview Event Handler-*********************************************
    Private Sub Grid1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid1.CellClick
        tex.init(Panel1, bs)
        RS_InfoTextBox.Text = Search_f(RSTextBox.Text, "[O|V][-]+\d+")
        TS_InfoTextBox.Text = Search_f(TSTextBox.Text, "[O|V][-]+\d+")
        SD_InfoTextBox.Text = Search_f(SDTextBox.Text, "[O|V][-]+\d+")
    End Sub

    Private Sub Grid2_CellClick(sender As Object, e As DataGridViewCellEventArgs)
        tex.init(Panel1, bs2)
        RS_InfoTextBox.Text = Search_f(RSTextBox.Text, "[O|V][-]+\d+")
        TS_InfoTextBox.Text = Search_f(TSTextBox.Text, "[O|V][-]+\d+")
        SD_InfoTextBox.Text = Search_f(SDTextBox.Text, "[O|V][-]+\d+")
    End Sub

    Private Sub Grid3_CellClick(sender As Object, e As DataGridViewCellEventArgs)
        tex.init(Panel1, bs3)
        RS_InfoTextBox.Text = Search_f(RSTextBox.Text, "[O|V][-]+\d+")
        TS_InfoTextBox.Text = Search_f(TSTextBox.Text, "[O|V][-]+\d+")
        SD_InfoTextBox.Text = Search_f(SDTextBox.Text, "[O|V][-]+\d+")
    End Sub

    '*********************************-Textbox Event Handler-*********************************************
    Private Sub bsEnd()
        bs.EndEdit()
        bs2.EndEdit()
        bs3.EndEdit()
    End Sub

    Private Sub gridEnd()
        Grid1.EndEdit()
        Grid2.EndEdit()
        Grid3.EndEdit()
    End Sub

    Private Sub ProjectTextBox_Leave(sender As Object, e As EventArgs) Handles ProjectTextBox.Leave
        bsEnd()
    End Sub

    Private Sub ModelTextBox_Leave(sender As Object, e As EventArgs) Handles ModelTextBox.Leave
        bsEnd()
    End Sub

    Private Sub ReleaseTextBox_Leave(sender As Object, e As EventArgs) Handles ReleaseTextBox.Leave
        bsEnd()
    End Sub

    Private Sub Task_IDTextBox_Leave(sender As Object, e As EventArgs) Handles Task_IDTextBox.Leave
        bsEnd()
    End Sub

    Private Sub ModulTextBox_Leave(sender As Object, e As EventArgs) Handles ModuleTextBox.Leave
        If Not String.IsNullOrEmpty(ModuleTextBox.Text) Then
            If InStr(ModuleTextBox.Text, ".cpp") Then : ModuleTextBox.Text = ModuleTextBox.Text.Replace(".cpp", "")
            ElseIf InStr(ModuleTextBox.Text, ".c") Then : ModuleTextBox.Text = ModuleTextBox.Text.Replace(".c", "")
            End If
        End If
        bsEnd()
    End Sub

    Private Sub ElocTextBox_Leave(sender As Object, e As EventArgs) Handles ElocTextBox.Leave
        bsEnd()
    End Sub

    Private Sub M_PathTextBox_Leave(sender As Object, e As EventArgs) Handles M_PathTextBox.Leave
        If Not String.IsNullOrEmpty(M_PathTextBox.Text) Then
            M_PathTextBox.Text = M_PathTextBox.Text.Replace("/", "\")
            If InStr(Mid(M_PathTextBox.Text, 1, 1), "\") = 0 Then M_PathTextBox.Text = "\" & M_PathTextBox.Text
        End If
        bsEnd()
    End Sub



    Private Sub RSTextBox_Leave(sender As Object, e As EventArgs) Handles RSTextBox.Leave
        RS_InfoTextBox.Text = Search_f(RSTextBox.Text, "[O|V][-]+\d+")
        TS_InfoTextBox.Text = Search_f(TSTextBox.Text, "[O|V][-]+\d+")
        SD_InfoTextBox.Text = Search_f(SDTextBox.Text, "[O|V][-]+\d+")
    End Sub

    Private Sub TSTextBox_Leave(sender As Object, e As EventArgs) Handles TSTextBox.Leave
        RS_InfoTextBox.Text = Search_f(RSTextBox.Text, "[O|V][-]+\d+")
        TS_InfoTextBox.Text = Search_f(TSTextBox.Text, "[O|V][-]+\d+")
        SD_InfoTextBox.Text = Search_f(SDTextBox.Text, "[O|V][-]+\d+")
    End Sub

    Private Sub SDTextBox_Leave(sender As Object, e As EventArgs) Handles SDTextBox.Leave
        RS_InfoTextBox.Text = Search_f(RSTextBox.Text, "[O|V][-]+\d+")
        TS_InfoTextBox.Text = Search_f(TSTextBox.Text, "[O|V][-]+\d+")
        SD_InfoTextBox.Text = Search_f(SDTextBox.Text, "[O|V][-]+\d+")
    End Sub


    '*********************************-Button Event Handler-*********************************************
    Public Task_ID, ModuleName, Eloc As String

    Private Sub addbtn_Click(sender As Object, e As EventArgs) Handles addbtn.Click
        Dim errorMsg As String = Nothing
        If dgv.add(AddTextBox, Grid1, ds.Tables("Task"), False, False, errorMsg) Then
            OutputTextBox.Text = errorMsg
        End If
        tex.init(Panel1, bs)
    End Sub

    Private Sub findbtn_Click(sender As Object, e As EventArgs) Handles findbtn.Click
        Dim errorMsg As String = Nothing
        If dgv.find(AddTextBox, Grid1, ds.Tables("Task"), errorMsg) Then
            OutputTextBox.Text = errorMsg
        End If
    End Sub

    Private Sub removebtn_Click(sender As Object, e As EventArgs) Handles removebtn.Click
        Dim errorMsg As String = Nothing
        Dim result As Integer = MessageBox.Show("Do you want to remove this task", "caption", MessageBoxButtons.OKCancel)
        If result = DialogResult.OK Then
            If dgv.removeA(Grid1, ds.Tables("Task"), errorMsg) Then
                OutputTextBox.Text = errorMsg
            End If
        End If
    End Sub

    Private Sub movebtn_Click(sender As Object, e As EventArgs) Handles movebtn.Click
        bsEnd() : gridEnd()
        Dim errorMsg As String = Nothing
        If dgv.move(Grid1, ds.Tables("Task"), True, errorMsg) Then
            OutputTextBox.Text = errorMsg
        End If
    End Sub

    Private Sub addRbtn_Click(sender As Object, e As EventArgs) Handles addRbtn.Click
        Dim errorMsg As String = Nothing
        If dgv.add(AddRTextBox, Grid2, ds.Tables("Task"), True, False, errorMsg) Then
            OutputTextBox.Text = errorMsg
        End If
        tex.init(Panel1, bs2)
    End Sub

    Private Sub findRbtn_Click(sender As Object, e As EventArgs) Handles findRbtn.Click
        Dim errorMsg As String = Nothing
        If dgv.find(AddRTextBox, Grid2, ds.Tables("Task"), errorMsg) Then
            OutputTextBox.Text = errorMsg
        End If
    End Sub

    Private Sub removeRbtn_Click(sender As Object, e As EventArgs) Handles removeRbtn.Click
        Dim errorMsg As String = Nothing
        Dim result As Integer = MessageBox.Show("Do you want to remove this task", "caption", MessageBoxButtons.OKCancel)
        If result = DialogResult.OK Then
            If dgv.removeA(Grid2, ds.Tables("Task"), errorMsg) Then
                OutputTextBox.Text = errorMsg
            End If
        End If
    End Sub

    Private Sub moveRbtn_Click(sender As Object, e As EventArgs) Handles moveRbtn.Click
        bsEnd() : gridEnd()
        Dim errorMsg As String = Nothing
        bs.EndEdit()

        If dgv.move(Grid2, ds.Tables("Task"), True, errorMsg) Then
            OutputTextBox.Text = errorMsg
        End If
    End Sub

    Private Sub findHbtn_Click(sender As Object, e As EventArgs) Handles findHbtn.Click
        Dim errorMsg As String = Nothing
        If dgv.find(hisTextBox, Grid3, ds.Tables("Task"), errorMsg) Then
            OutputTextBox.Text = errorMsg
        End If
    End Sub

    Private Sub removeHbtn_Click(sender As Object, e As EventArgs) Handles removeHbtn.Click
        Dim errorMsg As String = Nothing
        Dim result As Integer = MessageBox.Show("Do you want to remove this task", "caption", MessageBoxButtons.OKCancel)
        If result = DialogResult.OK Then
            If dgv.removeA(Grid3, ds.Tables("Task"), errorMsg) Then
                OutputTextBox.Text = errorMsg
            End If
        End If
    End Sub

    Private Sub recoverHbtn_Click(sender As Object, e As EventArgs) Handles recoverHbtn.Click
        bsEnd() : gridEnd()
        Dim errorMsg As String = Nothing
        If dgv.move(Grid3, ds.Tables("Task"), False, errorMsg) Then
            OutputTextBox.Text = errorMsg
        End If
    End Sub

    '*****************************************************************************************************************************************************
    '               
    '                                                               -Tracking Features-
    '
    '******************************************************************************************************************************************************
    Private Sub Get_Info(ByRef taskinfo() As TContainer.TaskInfo, ByRef userinfo As TContainer.UserInfo)
        Dim selrow() As DataRow = Nothing
        If TabControl1.SelectedTab.Name = "TabPage1" Then : multiselect(Grid1, selrow)
        ElseIf TabControl1.SelectedTab.Name = "TabPage2" Then : multiselect(Grid2, selrow)
        ElseIf TabControl1.SelectedTab.Name = "TabPage3" Then : multiselect(Grid3, selrow)
        End If
        Dim index As Integer = 0
        If selrow IsNot Nothing Then
            For Each r As DataRow In selrow
                ReDim Preserve taskinfo(selrow.Count - 1)
                taskinfo(index) = New TContainer.TaskInfo(r)
                index = index + 1
            Next
            userinfo = New TContainer.UserInfo(ds.Tables("User_Config").Rows(0))
        End If
        If taskinfo Is Nothing Or userinfo Is Nothing Then
            OutputTextBox.Text = "- No row selected."
        End If
    End Sub



    Private Delegate Sub _Set_Hash_Commit(ByRef taskinfo As TContainer.TaskInfo, ByRef GetHashCommit As Button.GetHashCommit)
    Public Sub Set_Hash_Commit(ByRef taskinfo As TContainer.TaskInfo, ByRef GetHashCommit As Button.GetHashCommit)
        If ToolStrip1.InvokeRequired Then
            Dim d As New _Set_Hash_Commit(AddressOf Set_Hash_Commit)
            Invoke(d, New Object() {taskinfo, GetHashCommit})
        Else
            bsEnd()
            If String.IsNullOrEmpty(taskinfo.Revision) Then
                taskinfo.Revision = GetHashCommit.hash
            Else
                If taskinfo.Revision <> GetHashCommit.hash Then
                    Dim result = MessageBox.Show("Your current Hash is: " & taskinfo.Revision & vbNewLine &
                           "New Hash is: " & GetHashCommit.hash & vbNewLine &
                           "Would you like to replace it?", "Override Confirmation", MessageBoxButtons.YesNo)
                    If result = DialogResult.Yes Then
                        taskinfo.Revision = GetHashCommit.hash
                    End If
                End If
            End If

            If String.IsNullOrEmpty(taskinfo.Feature_Branch) Then
                taskinfo.Feature_Branch = GetHashCommit.commit
            Else
                If taskinfo.Feature_Branch <> GetHashCommit.commit Then
                    Dim result = MessageBox.Show("Your current Commit is: " & taskinfo.Feature_Branch & vbNewLine &
                           "New commit is: " & GetHashCommit.commit & vbNewLine &
                           "Would you like to replace it?", "Override Confirmation", MessageBoxButtons.YesNo)
                    If result = DialogResult.Yes Then
                        taskinfo.Feature_Branch = GetHashCommit.commit
                    End If
                End If
            End If
            bsEnd()
        End If
    End Sub

    Private Delegate Sub _Set_Status(ByVal btn As Object, ByVal state As Boolean)
    Public Sub Set_Status(ByVal btn As Object, ByVal state As Boolean)
        If ToolStrip1.InvokeRequired Then
            Dim d As New _Set_Status(AddressOf Set_Status)
            Invoke(d, New Object() {btn, state})
        Else
            btn.Enabled = state
        End If
    End Sub

    Delegate Sub _Set_Content(ByRef tb As MetroFramework.Controls.MetroTextBox, ByRef text As String)
    Public Sub Set_Content(ByRef tb As MetroFramework.Controls.MetroTextBox, ByRef text As String)
        If tb.InvokeRequired Then
            Dim d As New _Set_Content(AddressOf Set_Content)
            Invoke(d, New Object() {tb, text})
        Else
            tb.Text = text
        End If
    End Sub

    Private Delegate Sub _Set_Visible(ByVal btn As Object, ByVal state As Boolean)
    Public Sub Set_Visibled(ByVal btn As Object, ByVal state As Boolean)
        If ToolStrip1.InvokeRequired Then
            Dim d As New _Set_Visible(AddressOf Set_Visibled)
            Invoke(d, New Object() {btn, state})
        Else
            btn.visible = state
        End If
    End Sub

    '*********************************************-Get CQ Infomation-***************************************************
    Public IsReviewTab As Boolean

    Private Delegate Sub _Set_JIRA_Info(ByRef taskinfo As TContainer.TaskInfo, ByRef GetCQorJIRA As Button.GetCQorJIRA)
    Public Sub Set_JIRA_Info(ByRef taskinfo As TContainer.TaskInfo, ByRef GetCQorJIRA As Button.GetCQorJIRA)
        If ToolStrip1.InvokeRequired Then
            Dim d As New _Set_JIRA_Info(AddressOf Set_JIRA_Info)
            Invoke(d, New Object() {taskinfo, GetCQorJIRA})
        Else
            taskinfo.Project = GetCQorJIRA.project
            taskinfo.ModuleName = GetCQorJIRA.modulename
            taskinfo.M_Path = GetCQorJIRA.modulepath
            taskinfo.M_Owner = GetCQorJIRA.moduleowner
            If IsReviewTab = True Then
                taskinfo.Reviewer = GetCQorJIRA.taskauthor
            End If
            bsEnd()
        End If
    End Sub

    Private Sub GetCQInfo(taskinfo As TContainer.TaskInfo, userinfo As TContainer.UserInfo)
        Dim errormsg As String = Nothing
        Dim taskid As New Information.TaskID(taskinfo.Task_ID)

        Set_Status(GetJIRAInfoBtn, False)
        Set_Content(OutputTextBox, "The tool is running, please be patient!" & vbNewLine)
        Dim getCQorJIRA As New Button.GetCQorJIRA(taskid)
        errormsg = getCQorJIRA.Execute()
        If String.IsNullOrEmpty(errormsg) Then
            Set_JIRA_Info(taskinfo, getCQorJIRA)
            Set_Content(OutputTextBox, OutputTextBox.Text & vbNewLine & vbNewLine & "Finished!")
        Else
            Set_Content(OutputTextBox, errormsg)
        End If
        Set_Status(GetJIRAInfoBtn, True)
    End Sub

    Private Sub GetCQInfoBtn_Click(sender As Object, e As EventArgs) Handles GetJIRAInfoBtn.Click
        Dim taskinfo() As TContainer.TaskInfo = Nothing
        Dim userinfo As TContainer.UserInfo = Nothing
        Get_Info(taskinfo, userinfo)

        IsReviewTab = False
        If TabControl1.SelectedTab.Name = "TabPage2" Then
            IsReviewTab = True
        End If

        If taskinfo IsNot Nothing And userinfo IsNot Nothing Then
            For Each t As TContainer.TaskInfo In taskinfo
                Dim _GetCQInfo As New Thread(Sub() GetCQInfo(t, userinfo))
                If _GetCQInfo.IsAlive = False Then
                    _GetCQInfo = New Thread(Sub() GetCQInfo(t, userinfo))
                End If
                _GetCQInfo.Start()
            Next
        End If
    End Sub

    '*********************************************-Get Hash and Commit-***************************************************
    Private Sub GetHashCommit(taskinfo As TContainer.TaskInfo, userinfo As TContainer.UserInfo)
        Dim errormsg As String = Nothing
        Dim sfd As New Information.CheckNull("Sandbox Folder Dir", userinfo.Sfd)
        Dim tfd As New Information.CheckNull("Tool Folder Dir", userinfo.Tfd)
        Dim sandbox As New Information.CheckNull("Sandbox", taskinfo.Sandbox)
        Dim modulepath As New Information.ModulePath(taskinfo.M_Path)

        Set_Status(GetHashCommitBtn, False)
        Set_Content(OutputTextBox, "The tool is running, please be patient!" & vbNewLine)
        Dim getHashCommit As New Button.GetHashCommit(sfd, sandbox, modulepath)
        errormsg = getHashCommit.Execute()
        If String.IsNullOrEmpty(errormsg) Then
            Set_Hash_Commit(taskinfo, getHashCommit)
            Set_Content(OutputTextBox, OutputTextBox.Text & vbNewLine & vbNewLine & "Finished!")
        Else
            Set_Content(OutputTextBox, errormsg)
        End If
        Set_Status(GetHashCommitBtn, True)
    End Sub

    Private Sub GetHashCommitBtn_Click(sender As Object, e As EventArgs) Handles GetHashCommitBtn.Click
        Dim taskinfo() As TContainer.TaskInfo = Nothing
        Dim userinfo As TContainer.UserInfo = Nothing
        Get_Info(taskinfo, userinfo)

        If taskinfo IsNot Nothing And userinfo IsNot Nothing Then
            For Each t As TContainer.TaskInfo In taskinfo
                Dim _GetHashCommit As New Thread(Sub() GetHashCommit(t, userinfo))
                If _GetHashCommit.IsAlive = False Then
                    _GetHashCommit = New Thread(Sub() GetHashCommit(t, userinfo))
                End If
                _GetHashCommit.Start()
            Next
        End If
    End Sub

    '*********************************************-Get Document files-***************************************************
    Private Sub GetDocFiles(taskinfo As TContainer.TaskInfo, userinfo As TContainer.UserInfo)
        Dim errormsg As String = Nothing

        Dim project As New Information.CheckNull("Project", taskinfo.Project)
        Dim taskid As New Information.TaskID(taskinfo.Task_ID)
        Dim modulename As New Information.CheckNull("Module Name", taskinfo.ModuleName)
        Dim oldtask As New Information.NoCheck(taskinfo.Old_Task)

        Dim Wfd As New Information.CheckNull("Working Folder Dir", userinfo.Wfd)
        Dim MyName As New Information.CheckNull("Your Name", userinfo.MyName)
        Dim FilesTemplateDir As New Information.CheckNull("Document Files Template Dir", userinfo.FilesTemplateDir)

        Set_Status(GetDocFilesBtn, False)
        Set_Content(OutputTextBox, "The tool is running, please be patient!" & vbNewLine)
        Dim getdocbtn As New Button.GetDocument(Wfd,
                                                project,
                                                taskid,
                                                modulename,
                                                oldtask,
                                                MyName,
                                                FilesTemplateDir)
        errormsg = getdocbtn.Execute()
        If String.IsNullOrEmpty(errormsg) Then
            Set_Content(OutputTextBox, OutputTextBox.Text & vbNewLine & vbNewLine & "Finished!")
        Else
            Set_Content(OutputTextBox, errormsg)
        End If
        Set_Status(GetDocFilesBtn, True)
    End Sub

    Private Sub GetDocFilesBtn_Click(sender As Object, e As EventArgs) Handles GetDocFilesBtn.Click
        Dim taskinfo() As TContainer.TaskInfo = Nothing
        Dim userinfo As TContainer.UserInfo = Nothing
        Get_Info(taskinfo, userinfo)

        If taskinfo IsNot Nothing And userinfo IsNot Nothing Then
            For Each t As TContainer.TaskInfo In taskinfo
                Dim _GetDocFiles As New Thread(Sub() GetDocFiles(t, userinfo))
                If _GetDocFiles.IsAlive = False Then
                    _GetDocFiles = New Thread(Sub() GetDocFiles(t, userinfo))
                End If
                _GetDocFiles.Start()
            Next
        End If
    End Sub

    '*********************************************-Auto Execution-***************************************************
    Private Sub AutoExe(taskinfo As TContainer.TaskInfo, userinfo As TContainer.UserInfo)
        Dim errormsg As String = Nothing

        Dim sandbox As New Information.CheckNull("Sandbox", taskinfo.Sandbox)
        Dim modulepath As New Information.ModulePath(taskinfo.M_Path)

        Dim sfd As New Information.CheckNull("Sandbox Folder Dir", userinfo.Sfd)
        Dim tfd As New Information.CheckNull("Tool Folder Dir", userinfo.Tfd)
        Dim CMacro As New Information.CheckNull("C Macro", userinfo.CMacro)
        Dim CppMacro As New Information.CheckNull("Cpp Macro", userinfo.CppMacro)
        Dim SysPMacro As New Information.CheckNull("SystemBPlus Macro", userinfo.SysBPlusMacro)

        Set_Status(CantataAutoExeBtn, False)
        Set_Content(OutputTextBox, "The tool is running, please be patient!" & vbNewLine)
        Dim autoexe As New Button.AutoExecution(sfd, tfd, sandbox, modulepath, CMacro, CppMacro, SysPMacro)
        errormsg = autoexe.Execute()
        If String.IsNullOrEmpty(errormsg) Then
            Set_Content(OutputTextBox, OutputTextBox.Text & vbNewLine & vbNewLine & "Finished!")
        Else
            Set_Content(OutputTextBox, errormsg)
        End If
        Set_Status(CantataAutoExeBtn, True)
    End Sub

    Private Sub CantataAutoExeBtn_Click(sender As Object, e As EventArgs) Handles CantataAutoExeBtn.Click
        Dim taskinfo() As TContainer.TaskInfo = Nothing
        Dim userinfo As TContainer.UserInfo = Nothing
        Get_Info(taskinfo, userinfo)

        If taskinfo Is Nothing Or userinfo Is Nothing Then
            OutputTextBox.Text = "- No row selected."
            Exit Sub
        End If

        If taskinfo IsNot Nothing Then
            For Each t As TContainer.TaskInfo In taskinfo
                Dim _AutoExe As New Thread(Sub() AutoExe(t, userinfo))
                If _AutoExe.IsAlive = False Then
                    _AutoExe = New Thread(Sub() AutoExe(t, userinfo))
                End If
                _AutoExe.Start()
            Next
        End If
    End Sub

    '*******************************************************-Go To-*************************************************************
    '****Go to Document
    Private Sub GoDoc(taskinfo As TContainer.TaskInfo, userinfo As TContainer.UserInfo, IsRev As Boolean)
        Dim errormsg As String = Nothing

        Dim project As New Information.CheckNull("Project", taskinfo.Project)
        Dim model As New Information.CheckNull("Model", taskinfo.Model)
        Dim taskid As New Information.TaskID(taskinfo.Task_ID)
        Dim modulename As New Information.CheckNull("Module Name", taskinfo.ModuleName)

        Dim Wfd As New Information.CheckNull("Working Folder Dir", userinfo.Wfd)
        Dim rfd As New Information.CheckNull("Review Folder Dir", userinfo.Rfd)
        Dim ExplorerPath As New Information.CheckNull("Explorer Dir", userinfo.ExplorerPath)

        Set_Status(GoDocBtn, False)
        If IsRev = True Then
            Dim gorevbtn As New Button.GotoReview(ExplorerPath, rfd, taskid, modulename)
            errormsg = gorevbtn.Execute()
        Else
            Dim godocbtn As New Button.GotoDocument(ExplorerPath, Wfd, project, taskid, modulename)
            errormsg = godocbtn.Execute()
        End If
        If Not String.IsNullOrEmpty(errormsg) Then
            Set_Content(OutputTextBox, errormsg)
        End If
        Set_Status(GoDocBtn, True)
    End Sub
    Private Sub GoDocBtn_Click(sender As Object, e As EventArgs) Handles GoDocBtn.Click
        Dim IsRev As Boolean = False

        Dim taskinfo() As TContainer.TaskInfo = Nothing
        Dim userinfo As TContainer.UserInfo = Nothing
        Get_Info(taskinfo, userinfo)

        If taskinfo Is Nothing Or userinfo Is Nothing Then
            OutputTextBox.Text = "- No row selected."
            Exit Sub
        End If

        If TabControl1.SelectedTab.Name = "TabPage2" Then
            IsRev = True
        End If

        If taskinfo IsNot Nothing Then
            For Each t As TContainer.TaskInfo In taskinfo
                Dim _GoDoc As New Thread(Sub() GoDoc(t, userinfo, IsRev))
                If _GoDoc.IsAlive = False Then
                    _GoDoc = New Thread(Sub() GoDoc(t, userinfo, IsRev))
                End If
                _GoDoc.Start()
            Next
        End If


    End Sub
    '****Go to test folder
    Private Sub GoTestFolder(taskinfo As TContainer.TaskInfo, userinfo As TContainer.UserInfo)
        Dim errormsg As String = Nothing

        Dim sandbox As New Information.CheckNull("Sandbox", taskinfo.Sandbox)
        Dim modulepath As New Information.ModulePath(taskinfo.M_Path)
        Dim modulename As New Information.CheckNull("Module Name", taskinfo.ModuleName)

        Dim ExplorerPath As New Information.CheckNull("Explorer Dir", userinfo.ExplorerPath)
        Dim sfd As New Information.CheckNull("Sandbox Folder Dir", userinfo.Sfd)

        Set_Status(GoTestBtn, False)
        Dim gotest As New Button.GotoTest(ExplorerPath, sfd, sandbox, modulepath, modulename)
        errormsg = gotest.Execute()
        If Not String.IsNullOrEmpty(errormsg) Then
            Set_Content(OutputTextBox, errormsg)
        End If
        Set_Status(GoTestBtn, True)
    End Sub

    Private Sub GoTest_Click(sender As Object, e As EventArgs) Handles GoTestBtn.Click
        Dim taskinfo() As TContainer.TaskInfo = Nothing
        Dim userinfo As TContainer.UserInfo = Nothing
        Get_Info(taskinfo, userinfo)

        If taskinfo Is Nothing Or userinfo Is Nothing Then
            OutputTextBox.Text = "- No row selected."
            Exit Sub
        End If

        If taskinfo IsNot Nothing Then
            For Each t As TContainer.TaskInfo In taskinfo
                Dim _GoTest As New Thread(Sub() GoTestFolder(t, userinfo))
                If _GoTest.IsAlive = False Then
                    _GoTest = New Thread(Sub() GoTestFolder(t, userinfo))
                End If
                _GoTest.Start()
            Next
        End If
    End Sub

    '****Go to Sandbox
    Private Sub GoSand(taskinfo As TContainer.TaskInfo, userinfo As TContainer.UserInfo)
        Dim errormsg As String = Nothing

        Dim sandbox As New Information.CheckNull("Sandbox", taskinfo.Sandbox)
        Dim modulepath As New Information.ModulePath(taskinfo.M_Path)

        Dim ExplorerPath As New Information.CheckNull("Explorer Dir", userinfo.ExplorerPath)
        Dim sfd As New Information.CheckNull("Sandbox Folder Dir", userinfo.Sfd)

        Set_Status(GoSandBtn, False)
        Dim gosandboxbtn As New Button.GotoSandBox(ExplorerPath, sfd, sandbox, modulepath)
        errormsg = gosandboxbtn.Execute()
        If Not String.IsNullOrEmpty(errormsg) Then
            Set_Content(OutputTextBox, errormsg)
        End If
        Set_Status(GoSandBtn, True)
    End Sub

    Private Sub GoSandBtn_Click(sender As Object, e As EventArgs) Handles GoSandBtn.Click
        Dim taskinfo() As TContainer.TaskInfo = Nothing
        Dim userinfo As TContainer.UserInfo = Nothing
        Get_Info(taskinfo, userinfo)

        If taskinfo Is Nothing Or userinfo Is Nothing Then
            OutputTextBox.Text = "- No row selected."
            Exit Sub
        End If

        If taskinfo IsNot Nothing Then
            For Each t As TContainer.TaskInfo In taskinfo
                Dim _GoSand As New Thread(Sub() GoSand(t, userinfo))
                If _GoSand.IsAlive = False Then
                    _GoSand = New Thread(Sub() GoSand(t, userinfo))
                End If
                _GoSand.Start()
            Next
        End If
    End Sub
    '****Go to CQ or JIRA
    Private Sub GotoCQorJIRA(taskinfo As TContainer.TaskInfo, userinfo As TContainer.UserInfo)
        Dim errormsg As String = Nothing
        Dim taskid As New Information.TaskID(taskinfo.Task_ID)
        Dim cquser As New Information.CheckNull("CQ User", userinfo.CQUser)
        Dim cqpassword As New Information.CheckNull("CQ Password", userinfo.CQPassword)

        Set_Status(GoCQorJIRABtn, False)
        Dim GotoCQorJIRA As New Button.GotoCQOrJIRA(taskid, cquser, cqpassword)
        errormsg = GotoCQorJIRA.Execute()
        If String.IsNullOrEmpty(errormsg) Then
            Set_Content(OutputTextBox, OutputTextBox.Text & vbNewLine & vbNewLine & "Finished!")
        Else
            Set_Content(OutputTextBox, errormsg)
        End If
        Set_Status(GoCQorJIRABtn, True)
    End Sub

    Private Sub GoCQorJIRABtn_Click(sender As Object, e As EventArgs) Handles GoCQorJIRABtn.Click
        Dim taskinfo() As TContainer.TaskInfo = Nothing
        Dim userinfo As TContainer.UserInfo = Nothing
        Get_Info(taskinfo, userinfo)

        If taskinfo Is Nothing Or userinfo Is Nothing Then
            OutputTextBox.Text = "- No row selected."
            Exit Sub
        End If

        If taskinfo IsNot Nothing Then
            For Each t As TContainer.TaskInfo In taskinfo
                Dim _GoCQorJIRA As New Thread(Sub() GotoCQorJIRA(t, userinfo))
                If _GoCQorJIRA.IsAlive = False Then
                    _GoCQorJIRA = New Thread(Sub() GotoCQorJIRA(t, userinfo))
                End If
                _GoCQorJIRA.Start()
            Next
        End If
    End Sub

    '****Go to OPL
    Private Sub GoOPL(taskinfo As TContainer.TaskInfo, userinfo As TContainer.UserInfo)
        Dim errormsg As String = Nothing
        Dim opl_link As New Information.CheckNull("OPL Link", taskinfo.OPL_Link)
        Dim explorerPath As New Information.CheckNull("explorer Dir", userinfo.ExplorerPath)

        Set_Status(GoOPLBtn, False)

        Dim btnGoOPL As New Button.GotoOPL(explorerPath, opl_link)
        errormsg = btnGoOPL.Execute()
        If String.IsNullOrEmpty(errormsg) Then
            Set_Content(OutputTextBox, OutputTextBox.Text & vbNewLine & vbNewLine & "Finished!")
        Else
            Set_Content(OutputTextBox, errormsg)
        End If
        Set_Status(GoOPLBtn, True)
    End Sub

    Private Sub GoOPLBtn_Click(sender As Object, e As EventArgs) Handles GoOPLBtn.Click
        Dim taskinfo() As TContainer.TaskInfo = Nothing
        Dim userinfo As TContainer.UserInfo = Nothing
        Get_Info(taskinfo, userinfo)

        If taskinfo Is Nothing Or userinfo Is Nothing Then
            OutputTextBox.Text = "- No row selected."
            Exit Sub
        End If

        If taskinfo IsNot Nothing Then
            For Each t As TContainer.TaskInfo In taskinfo
                Dim _GoOPL As New Thread(Sub() GoOPL(t, userinfo))
                If _GoOPL.IsAlive = False Then
                    _GoOPL = New Thread(Sub() GoOPL(t, userinfo))
                End If
                _GoOPL.Start()
            Next
        End If
    End Sub

    '****Go to Result
    Private Sub GotoResult(taskinfo As TContainer.TaskInfo, userinfo As TContainer.UserInfo)
        Dim errormsg As String = Nothing
        Dim result_Path As New Information.CheckNull("Result Path", taskinfo.Result_Path)
        Dim explorerPath As New Information.CheckNull("explorer Dir", userinfo.ExplorerPath)

        Set_Status(GoResultBtn, False)

        Dim btnGoResult As New Button.GotoResult(explorerPath, result_Path)
        errormsg = btnGoResult.Execute()
        If String.IsNullOrEmpty(errormsg) Then
            Set_Content(OutputTextBox, OutputTextBox.Text & vbNewLine & vbNewLine & "Finished!")
        Else
            Set_Content(OutputTextBox, errormsg)
        End If
        Set_Status(GoResultBtn, True)
    End Sub

    Private Sub GoResultBtn_Click(sender As Object, e As EventArgs) Handles GoResultBtn.Click
        Dim taskinfo() As TContainer.TaskInfo = Nothing
        Dim userinfo As TContainer.UserInfo = Nothing
        Get_Info(taskinfo, userinfo)

        If taskinfo Is Nothing Or userinfo Is Nothing Then
            OutputTextBox.Text = "- No row selected."
            Exit Sub
        End If

        If taskinfo IsNot Nothing Then
            For Each t As TContainer.TaskInfo In taskinfo
                Dim _GoResult As New Thread(Sub() GotoResult(t, userinfo))
                If _GoResult.IsAlive = False Then
                    _GoResult = New Thread(Sub() GotoResult(t, userinfo))
                End If
                _GoResult.Start()
            Next
        End If
    End Sub

    '****Go to ILM Server
    Private Delegate Sub _SetILMLink(ByVal btn As MetroFramework.Controls.MetroButton, ByRef taskinfo As TContainer.TaskInfo, ByVal ILM_Link As String)
    Public Sub SetILMLink(ByVal btn As MetroFramework.Controls.MetroButton, ByRef taskinfo As TContainer.TaskInfo, ByVal ILM_Link As String)
        If ToolStrip1.InvokeRequired Then
            Dim d As New _SetILMLink(AddressOf SetILMLink)
            Invoke(d, New Object() {btn, taskinfo, ILM_Link})
        Else
            taskinfo.ILM_Link = ILM_Link
            bsEnd()
        End If
    End Sub

    Private Sub GotoILM(taskinfo As TContainer.TaskInfo, userinfo As TContainer.UserInfo)
        Dim ErrorMsg As String = Nothing
        Dim ILMLink As String = Nothing
        Dim ILMServerPath As String = Nothing
        Dim FireFoxPath As String = "C:\Program Files (x86)\Mozilla Firefox\firefox.exe"

        Dim project As New Information.CheckNull("Project", taskinfo.Project)
        Dim model As New Information.CheckNull("Model", taskinfo.Model)
        Dim taskid As New Information.TaskID(taskinfo.Task_ID)

        Set_Status(GoILMBtn, False)
        If String.IsNullOrEmpty(ILM_LinkTextBox.Text) Then
            Set_Content(OutputTextBox, "ILM is searching...")
            If Search_VSS_Link(taskid.GetValue, project.GetValue, ILMLink, ErrorMsg) = True Then
                Set_Content(OutputTextBox, ErrorMsg)
            Else
                ILMServerPath = "https://inside-ilm.bosch.com/irj/go/nui/nav/.../versions" + ILMLink
                Shell(FireFoxPath & " -url" & " " & ILMServerPath)
                Set_Content(ILM_LinkTextBox, ILMServerPath)
                Set_Content(OutputTextBox, OutputTextBox.Text & vbNewLine & "Go to ILM is Successul")
            End If
        Else
            ILMServerPath = ILM_LinkTextBox.Text
            Shell(FireFoxPath & " -url" & " " & ILMServerPath)
        End If

        Set_Status(GoILMBtn, True)
    End Sub

    Private Sub GoILMBtn_Click(sender As Object, e As EventArgs) Handles GoILMBtn.Click
        Dim taskinfo() As TContainer.TaskInfo = Nothing
        Dim userinfo As TContainer.UserInfo = Nothing
        Get_Info(taskinfo, userinfo)

        If taskinfo Is Nothing Or userinfo Is Nothing Then
            OutputTextBox.Text = "- No row selected."
            Exit Sub
        End If

        If taskinfo IsNot Nothing Then
            For Each t As TContainer.TaskInfo In taskinfo
                Dim _GoILM As New Thread(Sub() GotoILM(t, userinfo))
                If _GoILM.IsAlive = False Then
                    _GoILM = New Thread(Sub() GotoILM(t, userinfo))
                End If
                _GoILM.Start()
            Next
        End If
    End Sub




    '************************-Download Task Review
    Private Sub DownTaskReview(taskid As String, ByVal Dest As String, project As String)
        Dim LinkWorkingFolder As String = Dest

        Dim LinkILM As String = Nothing
        Dim ErrorMsg As String = Nothing

        Set_Status(TaskReviewBtn, False)
        Set_Content(OutputTextBox, "Link Searching...")
        Search_VSS_Link(taskid, project, LinkILM, ErrorMsg)
        If LinkILM <> "" Then
            Dim aPath() As String
            aPath = Split(LinkILM, "/")
            LinkWorkingFolder = LinkWorkingFolder + "\" + aPath(UBound(aPath))

            Try
                If (Not Directory.Exists(LinkWorkingFolder)) Then
                    Directory.CreateDirectory(LinkWorkingFolder)
                End If
            Catch ex As Exception
                Set_Content(OutputTextBox, OutputTextBox.Text & vbNewLine & ex.Message)
                Exit Sub
            End Try

            Set_Content(OutputTextBox, OutputTextBox.Text & vbNewLine & "Downloading...")
            Call ILM_DownloadItemInFolder(LinkILM, LinkWorkingFolder, OutputTextBox)
            Set_Content(OutputTextBox, OutputTextBox.Text & vbNewLine & "ILM download document is successful")
        Else
            Set_Content(OutputTextBox, OutputTextBox.Text & vbNewLine & "Can't find Old task")
        End If
        Set_Status(TaskReviewBtn, True)
    End Sub

    Private Sub DownloadReviewTaskToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DownloadReviewTaskToolStripMenuItem.Click
        Dim taskinfo() As TContainer.TaskInfo = Nothing
        Dim userinfo As TContainer.UserInfo = Nothing
        Get_Info(taskinfo, userinfo)

        If TabControl1.SelectedTab.Name <> "TabPage2" Then
            Set_Content(OutputTextBox, "You're not on 'Review' Tab" & vbNewLine & "Please go to tab 'Review'")
            Exit Sub
        End If

        If taskinfo IsNot Nothing And userinfo IsNot Nothing Then
            For Each t As TContainer.TaskInfo In taskinfo
                Dim _DownloadTaskReview As New Thread(Sub() DownTaskReview(t.Task_ID, userinfo.Rfd, t.Project))
                If _DownloadTaskReview.IsAlive = False Then
                    _DownloadTaskReview = New Thread(Sub() DownTaskReview(t.Task_ID, userinfo.Rfd, t.Project))
                End If
                _DownloadTaskReview.Start()
            Next
        End If
    End Sub

    Private Sub MoveToGitFolder(taskinfo As TContainer.TaskInfo, userinfo As TContainer.UserInfo, RevFlag As Boolean)
        Dim errormsg As String = Nothing

        Dim project As New Information.CheckNull("Project", taskinfo.Project)
        Dim taskid As New Information.TaskID(taskinfo.Task_ID)
        Dim modulename As New Information.CheckNull("Module Name", taskinfo.ModuleName)

        Dim sandbox As New Information.CheckNull("Sandbox", taskinfo.Sandbox)


        Dim ExplorerPath As New Information.CheckNull("explorer Dir", userinfo.ExplorerPath)
        Dim Wfd As New Information.CheckNull("Working Folder Dir", userinfo.Wfd)
        Dim rfd As New Information.CheckNull("Review Folder Dir", userinfo.Rfd)
        Dim sfd As New Information.CheckNull("Sandbox Folder Dir", userinfo.Sfd)


        Set_Status(TaskReviewBtn, False)
        Set_Content(OutputTextBox, "The tool is running, please be patient!" & vbNewLine)
        Dim btnMoveToGitFolder As New Button.MoveToGitFolder(ExplorerPath,
                                                             Wfd,
                                                             project,
                                                             taskid,
                                                             modulename,
                                                             sfd,
                                                             sandbox,
                                                             rfd,
                                                             RevFlag)
        errormsg = btnMoveToGitFolder.Execute()
        If String.IsNullOrEmpty(errormsg) Then
            Set_Content(OutputTextBox, OutputTextBox.Text & vbNewLine & vbNewLine & "Finished!")
        Else
            Set_Content(OutputTextBox, errormsg)
        End If
        Set_Status(TaskReviewBtn, True)
    End Sub

    Private Sub MoveToGitFolderToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MoveToGitFolderToolStripMenuItem.Click
        Dim taskinfo() As TContainer.TaskInfo = Nothing
        Dim userinfo As TContainer.UserInfo = Nothing
        Get_Info(taskinfo, userinfo)

        Dim IsRevTab As Boolean = False
        If TabControl1.SelectedTab.Name = "TabPage2" Then IsRevTab = True

        If taskinfo IsNot Nothing And userinfo IsNot Nothing Then
            For Each t As TContainer.TaskInfo In taskinfo
                Dim _MoveToGitFolder As New Thread(Sub() MoveToGitFolder(t, userinfo, IsRevTab))
                If _MoveToGitFolder.IsAlive = False Then
                    _MoveToGitFolder = New Thread(Sub() MoveToGitFolder(t, userinfo, IsRevTab))
                End If
                _MoveToGitFolder.Start()
            Next
        End If
    End Sub

    '*******************************************************-Download Old Task-*************************************************************
    Private Sub DownOldTask(taskid As String, ByVal Dest As String, project As String)
        Dim LinkWorkingFolder As String = Dest
        Dim ModuleName As String = Nothing

        Dim statusInfo As String = Nothing
        Dim LinkILM As String = Nothing
        Dim ErrorMsg As String = Nothing

        Set_Status(DownOldTaskBtn, False)
        Set_Content(OutputTextBox, "Link Searching...")
        Search_VSS_Link(taskid, project, LinkILM, ErrorMsg)
        If LinkILM <> "" Then
            LinkWorkingFolder = LinkWorkingFolder + "\" + taskid
            Try
                If (Not Directory.Exists(LinkWorkingFolder)) Then
                    Directory.CreateDirectory(LinkWorkingFolder)
                End If
            Catch ex As Exception
                Set_Content(OutputTextBox, OutputTextBox.Text & vbNewLine & ex.Message)
                Exit Sub
            End Try

            Set_Content(OutputTextBox, OutputTextBox.Text & vbNewLine & "Downloading...")
            Call ILM_DownloadItemInFolder(LinkILM, LinkWorkingFolder, OutputTextBox)
            Set_Content(OutputTextBox, OutputTextBox.Text & vbNewLine & "ILM download document is successful")
        Else
            Set_Content(OutputTextBox, OutputTextBox.Text & vbNewLine & "Can't find Old task")
        End If
        Set_Status(DownOldTaskBtn, True)
    End Sub

    Private Sub DownOldTaskBtn_Click(sender As Object, e As EventArgs) Handles DownOldTaskBtn.Click
        Dim taskinfo() As TContainer.TaskInfo = Nothing
        Dim userinfo As TContainer.UserInfo = Nothing
        Get_Info(taskinfo, userinfo)



        If taskinfo Is Nothing Or userinfo Is Nothing Then
            OutputTextBox.Text = "- No row selected."
            Exit Sub
        End If

        If taskinfo.Length > 1 Then OutputTextBox.Text = "Please chose only 1 task" : Exit Sub
        If String.IsNullOrEmpty(taskinfo(0).Old_Task) Then OutputTextBox.Text = "Old task field is empty" : Exit Sub

        Dim project As New Information.CheckNull("Project", taskinfo(0).Project)
        Dim model As New Information.CheckNull("Model", taskinfo(0).Model)
        Dim taskid As New Information.TaskID(taskinfo(0).Task_ID)
        Dim modulename As New Information.CheckNull("Module Name", taskinfo(0).ModuleName)

        Dim Wfd As New Information.CheckNull("Working Folder Dir", userinfo.Wfd)
        Dim rfd As New Information.CheckNull("Review Folder Dir", userinfo.Rfd)
        Dim ExplorerPath As New Information.CheckNull("Explorer Dir", userinfo.ExplorerPath)

        Dim OldProject = InputBox("Please Put Old Project for old task " & taskinfo(0).Old_Task & ":" & vbNewLine & "- Hint: Put 0 to Search Whole ILM Server", , taskinfo(0).Project)
        If Not String.IsNullOrEmpty(OldProject) Then
            If OldProject = "0" Then OldProject = ""
            Dim taskid_double As Double
            Double.TryParse(Search_f(taskinfo(0).Task_ID, "\d+"), taskid_double)
            Dim placeoldtask As String = Nothing
            If TabControl1.SelectedTab.Name = "TabPage1" Then
                Dim GotoDoc As New Button.GotoDocument(ExplorerPath, Wfd, project, taskid, modulename)
                placeoldtask = GotoDoc.GetFullPath
            ElseIf TabControl1.SelectedTab.Name = "TabPage2" Then
                Dim GotoRev As New Button.GotoReview(ExplorerPath, rfd, taskid, modulename)
                placeoldtask = GotoRev.GetFullPath
            End If

            If taskinfo IsNot Nothing Then
                For Each t As TContainer.TaskInfo In taskinfo
                    Dim _DownOldTask As New Thread(Sub() DownOldTask(t.Old_Task, placeoldtask, OldProject))
                    If _DownOldTask.IsAlive = False Then
                        _DownOldTask = New Thread(Sub() DownOldTask(t.Old_Task, placeoldtask, OldProject))
                    End If
                    _DownOldTask.Start()
                Next
            End If

        End If
    End Sub





    '*******************************************************-Fill Document-*************************************************************
    '******************-Fill OPL 
    Private Sub FillOPL(taskinfo As TContainer.TaskInfo, userinfo As TContainer.UserInfo)
        Dim errormsg As String = Nothing

        Dim project As New Information.CheckNull("Project", taskinfo.Project)
        Dim taskid As New Information.TaskID(taskinfo.Task_ID)
        Dim modulename As New Information.CheckNull("Module Name", taskinfo.ModuleName)
        Dim moduleowner As New Information.CheckNull("Module Owner", taskinfo.M_Owner)

        Dim Wfd As New Information.CheckNull("Working Folder Dir", userinfo.Wfd)
        Dim MyName As New Information.CheckNull("Your Name", userinfo.MyName)
        Dim ExplorerPath As New Information.CheckNull("explorer Dir", userinfo.ExplorerPath)

        Set_Status(FillDocumentBtn, False)
        Dim btnfillOPL As New Button.FillOPL(ExplorerPath,
                                                 Wfd,
                                                 project,
                                                 taskid,
                                                 modulename,
                                                 moduleowner,
                                                 MyName)
        errormsg = btnfillOPL.Execute()
        If String.IsNullOrEmpty(errormsg) Then
            Set_Content(OutputTextBox, OutputTextBox.Text & vbNewLine & vbNewLine & "Finished!")
        Else
            Set_Content(OutputTextBox, errormsg)
        End If
        Set_Status(FillDocumentBtn, True)
    End Sub

    Private Sub FillOPLToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FillOPLToolStripMenuItem.Click
        Dim taskinfo() As TContainer.TaskInfo = Nothing
        Dim userinfo As TContainer.UserInfo = Nothing
        Get_Info(taskinfo, userinfo)

        If taskinfo Is Nothing Or userinfo Is Nothing Then
            OutputTextBox.Text = "- No row selected."
            Exit Sub
        End If

        If taskinfo IsNot Nothing Then
            For Each t As TContainer.TaskInfo In taskinfo
                Dim _FillOPL As New Thread(Sub() FillOPL(t, userinfo))
                If _FillOPL.IsAlive = False Then
                    _FillOPL = New Thread(Sub() FillOPL(t, userinfo))
                End If
                _FillOPL.Start()
            Next
        End If
    End Sub

    '******************-Fill Coverage
    Private Sub FillCodeCoverage(taskinfo As TContainer.TaskInfo, userinfo As TContainer.UserInfo)
        Dim errormsg As String = Nothing
        Dim project As New Information.CheckNull("Project", taskinfo.Project)
        Dim taskid As New Information.TaskID(taskinfo.Task_ID)
        Dim modulename As New Information.CheckNull("Module Name", taskinfo.ModuleName)
        Dim modulepath As New Information.ModulePath(taskinfo.M_Path)
        Dim Statement As New Information.CheckNull("Statement", taskinfo.Statement)
        Dim Decisions As New Information.NoCheck(taskinfo.Decisions)
        Dim hashfile As New Information.CheckNull("Hash File", taskinfo.Revision)

        Dim Wfd As New Information.CheckNull("Working Folder Dir", userinfo.Wfd)
        Dim Rfd As New Information.CheckNull("Review Folder Dir", userinfo.Rfd)
        Dim MyName As New Information.CheckNull("Your Name", userinfo.MyName)
        Dim ExplorerPath As New Information.CheckNull("explorer Dir", userinfo.ExplorerPath)

        Set_Status(FillDocumentBtn, False)
        Set_Content(OutputTextBox, "The tool is running, please be patient!" & vbNewLine)
        Dim btnfillcoverage As New Button.FillCodeCoverage(ExplorerPath,
                                                           MyName,
                                                           Wfd,
                                                           taskid,
                                                           modulename,
                                                           Statement,
                                                           Decisions,
                                                           modulepath,
                                                           hashfile,
                                                           Rfd,
                                                          project)
        errormsg = btnfillcoverage.Execute()
        If String.IsNullOrEmpty(errormsg) Then
            Set_Content(OutputTextBox, OutputTextBox.Text & vbNewLine & vbNewLine & "Finished!")
        Else
            Set_Content(OutputTextBox, errormsg)
        End If
        Set_Status(FillDocumentBtn, True)
    End Sub
    Private Sub FillCodeCoverageToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FillCodeCoverageToolStripMenuItem.Click
        Dim taskinfo() As TContainer.TaskInfo = Nothing
        Dim userinfo As TContainer.UserInfo = Nothing
        Get_Info(taskinfo, userinfo)

        If taskinfo Is Nothing Or userinfo Is Nothing Then
            OutputTextBox.Text = "- No row selected."
            Exit Sub
        End If

        If taskinfo IsNot Nothing Then
            For Each t As TContainer.TaskInfo In taskinfo
                Dim _FillCodeCoverage As New Thread(Sub() FillCodeCoverage(t, userinfo))
                If _FillCodeCoverage.IsAlive = False Then
                    _FillCodeCoverage = New Thread(Sub() FillCodeCoverage(t, userinfo))
                End If
                _FillCodeCoverage.Start()
            Next
        End If
    End Sub
    '******************-Fill Checklists
    Private Sub FillChecklists(taskinfo As TContainer.TaskInfo, userinfo As TContainer.UserInfo)
        Dim errormsg As String = Nothing

        Dim project As New Information.CheckNull("Project", taskinfo.Project)
        Dim taskid As New Information.TaskID(taskinfo.Task_ID)
        Dim modulename As New Information.CheckNull("Module Name", taskinfo.ModuleName)
        Dim revision As New Information.NoCheck(taskinfo.Revision)
        Dim sandbox As New Information.CheckNull("Sandbox", taskinfo.Sandbox)
        Dim modulepath As New Information.ModulePath(taskinfo.M_Path)
        Dim Statement As New Information.CheckNull("Statement", taskinfo.Statement)
        Dim Decisions As New Information.NoCheck(taskinfo.Decisions)
        Dim opl As New Information.NoCheck(taskinfo.OPL_Link)
        Dim rs As New Information.NoCheck(taskinfo.RS)
        Dim rsbl As New Information.NoCheck(taskinfo.RS_BL)
        Dim ts As New Information.NoCheck(taskinfo.TS)
        Dim tsbl As New Information.NoCheck(taskinfo.TS_BL)
        Dim sd As New Information.NoCheck(taskinfo.SD)
        Dim sdbl As New Information.NoCheck(taskinfo.SD_BL)
        Dim reviewer As New Information.NoCheck(taskinfo.Reviewer)
        Dim defectid As New Information.NoCheck(taskinfo.Defect_ID)
        Dim oldtask As New Information.NoCheck(taskinfo.Old_Task)


        Dim Wfd As New Information.CheckNull("Working Folder Dir", userinfo.Wfd)
        Dim MyName As New Information.CheckNull("Your Name", userinfo.MyName)
        Dim ExplorerPath As New Information.CheckNull("explorer Dir", userinfo.ExplorerPath)
        Dim teamlead As New Information.CheckNull("Team Leader", userinfo.TeamLead)
        Dim sfd As New Information.CheckNull("Sandbox Folder Dir", userinfo.Sfd)

        Set_Status(FillDocumentBtn, False)
        Set_Content(OutputTextBox, "The tool is running, please be patient!" & vbNewLine)
        Dim btnfillchecklist As New Button.FillChecklist(ExplorerPath,
                                                         Wfd,
                                                         project,
                                                         taskid,
                                                         modulename,
                                                         MyName,
                                                         reviewer,
                                                         teamlead,
                                                         modulepath,
                                                         opl,
                                                         rs,
                                                         rsbl,
                                                         sd,
                                                         sdbl,
                                                         ts,
                                                         tsbl,
                                                         Statement,
                                                         Decisions,
                                                         defectid,
                                                         oldtask,
                                                         sfd,
                                                            sandbox)
        errormsg = btnfillchecklist.Execute()
        If String.IsNullOrEmpty(errormsg) Then
            Set_Content(OutputTextBox, OutputTextBox.Text & vbNewLine & vbNewLine & "Finished!")
        Else
            Set_Content(OutputTextBox, errormsg)
        End If
        Set_Status(FillDocumentBtn, True)
    End Sub

    Private Sub FillChecklistsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FillChecklistsToolStripMenuItem.Click
        Dim taskinfo() As TContainer.TaskInfo = Nothing
        Dim userinfo As TContainer.UserInfo = Nothing
        Get_Info(taskinfo, userinfo)

        If taskinfo Is Nothing Or userinfo Is Nothing Then
            OutputTextBox.Text = "- No row selected."
            Exit Sub
        End If

        If taskinfo IsNot Nothing Then
            For Each t As TContainer.TaskInfo In taskinfo
                Dim _FillChecklists As New Thread(Sub() FillChecklists(t, userinfo))
                If _FillChecklists.IsAlive = False Then
                    _FillChecklists = New Thread(Sub() FillChecklists(t, userinfo))
                End If
                _FillChecklists.Start()
            Next
        End If
    End Sub


    ''**********************-Needed Content Butoon --R
    Private Sub NeedContentBtn_Click(sender As Object, e As EventArgs) Handles NeedContentBtn.Click
        Dim errormsg As String = Nothing
        Dim taskinfo() As TContainer.TaskInfo = Nothing
        Dim userinfo As TContainer.UserInfo = Nothing
        Get_Info(taskinfo, userinfo)

        If taskinfo Is Nothing Or userinfo Is Nothing Then
            OutputTextBox.Text = "- No row selected."
            Exit Sub
        End If

        Dim project As New Information.CheckNull("Project", taskinfo(0).Project)
        Dim taskid As New Information.TaskID(taskinfo(0).Task_ID)
        Dim modulename As New Information.CheckNull("Module Name", taskinfo(0).ModuleName)
        Dim sandbox As New Information.CheckNull("Sandbox", taskinfo(0).Sandbox)
        Dim modulepath As New Information.ModulePath(taskinfo(0).M_Path)
        Dim hash As New Information.NoCheck(taskinfo(0).Revision)
        Dim commit As New Information.NoCheck(taskinfo(0).Feature_Branch)
        Dim statement As New Information.NoCheck(taskinfo(0).Statement)
        Dim decision As New Information.NoCheck(taskinfo(0).Decisions)
        Dim resultpath As New Information.NoCheck(taskinfo(0).Result_Path)
        Dim eloc As New Information.NoCheck(taskinfo(0).Eloc)

        Dim Tfd As New Information.CheckNull("Tool Folder Dir", userinfo.Tfd)
        Dim MyName As New Information.CheckNull("Your Name", userinfo.MyName)

        Dim btnNeededContent As New Button.NeededContent(Tfd,
                                                         project,
                                                         taskid,
                                                         modulename,
                                                         modulepath,
                                                         sandbox,
                                                         hash,
                                                         commit,
                                                         MyName,
                                                       statement,
                                                        decision,
                                                         resultpath,
                                                         eloc)
        errormsg = errormsg & btnNeededContent.Execute()

        OutputTextBox.Text = errormsg
    End Sub

    Private Sub JIRAStartTaskFill(taskinfo As TContainer.TaskInfo, userinfo As TContainer.UserInfo)
        Dim errormsg As String = Nothing

        Dim taskid As New Information.TaskID(taskinfo.Task_ID)
        Dim modulename As New Information.CheckNull("Module name", taskinfo.ModuleName)
        Dim toolpath As New Information.CheckNull("Tool folder path", userinfo.Tfd)
        Dim modulepath As New Information.ModulePath(taskinfo.M_Path)
        Dim cquser As New Information.CheckNull("CQ User", userinfo.CQUser)
        Dim cqpassword As New Information.CheckNull("CQ Password", userinfo.CQPassword)

        Set_Status(UploadJIRATaskInfoBtn, False)
        Set_Content(OutputTextBox, "The tool is running, please be patient!" & vbNewLine)
        Dim UploadStartTaskBtn As New Button.UploadJIRAStartTask(taskid, cquser, cqpassword, modulename, toolpath, modulepath)
        errormsg = UploadStartTaskBtn.Execute()
        If String.IsNullOrEmpty(errormsg) Then
            Set_Content(OutputTextBox, OutputTextBox.Text & vbNewLine & vbNewLine & "Finished!")
        Else
            Set_Content(OutputTextBox, errormsg)
        End If
        Set_Status(UploadJIRATaskInfoBtn, True)
    End Sub

    Private Sub StartTaskToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles StartTaskToolStripMenuItem.Click
        Dim taskinfo() As TContainer.TaskInfo = Nothing
        Dim userinfo As TContainer.UserInfo = Nothing
        Get_Info(taskinfo, userinfo)

        If taskinfo Is Nothing Or userinfo Is Nothing Then
            OutputTextBox.Text = "- No row selected."
            Exit Sub
        End If

        If taskinfo IsNot Nothing Then
            For Each t As TContainer.TaskInfo In taskinfo
                Dim _JIRAStartTaskFill As New Thread(Sub() JIRAStartTaskFill(t, userinfo))
                If _JIRAStartTaskFill.IsAlive = False Then
                    _JIRAStartTaskFill = New Thread(Sub() JIRAStartTaskFill(t, userinfo))
                End If
                _JIRAStartTaskFill.Start()
            Next
        End If
    End Sub

    Private Sub JIRADiliveryTaskFill(taskinfo As TContainer.TaskInfo, userinfo As TContainer.UserInfo)
        Dim errormsg As String = Nothing

        Dim taskid As New Information.TaskID(taskinfo.Task_ID)
        Dim modulename As New Information.CheckNull("Module name", taskinfo.ModuleName)
        Dim hash As New Information.NoCheck(taskinfo.Revision)
        Dim statement As New Information.NoCheck(taskinfo.Statement)
        Dim decision As New Information.NoCheck(taskinfo.Decisions)
        Dim eloc As New Information.NoCheck(taskinfo.Eloc)
        Dim commit As New Information.NoCheck(taskinfo.Feature_Branch)
        Dim resultpath As New Information.NoCheck(taskinfo.Result_Path)
        Dim toolpath As New Information.CheckNull("Tool folder path", userinfo.Tfd)
        Dim modulepath As New Information.ModulePath(taskinfo.M_Path)
        Dim cquser As New Information.CheckNull("CQ User", userinfo.CQUser)
        Dim cqpassword As New Information.CheckNull("CQ Password", userinfo.CQPassword)

        Set_Status(UploadJIRATaskInfoBtn, False)
        Set_Content(OutputTextBox, "The tool is running, please be patient!" & vbNewLine)
        Dim UploadDeliveryTaskBtn As New Button.UploadJIRADeliveryTask(taskid,
                                                                    cquser,
                                                                    cqpassword,
                                                                    modulename,
                                                                    toolpath,
                                                                    modulepath,
                                                                    hash,
                                                                    statement,
                                                                    decision,
                                                                    eloc,
                                                                    commit,
                                                                    resultpath)
        errormsg = UploadDeliveryTaskBtn.Execute()
        If String.IsNullOrEmpty(errormsg) Then
            Set_Content(OutputTextBox, OutputTextBox.Text & vbNewLine & vbNewLine & "Finished!")
        Else
            Set_Content(OutputTextBox, errormsg)
        End If
        Set_Status(UploadJIRATaskInfoBtn, True)
    End Sub

    Private Sub DeliveryTaskToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeliveryTaskToolStripMenuItem.Click
        Dim taskinfo() As TContainer.TaskInfo = Nothing
        Dim userinfo As TContainer.UserInfo = Nothing
        Get_Info(taskinfo, userinfo)

        If taskinfo Is Nothing Or userinfo Is Nothing Then
            OutputTextBox.Text = "- No row selected."
            Exit Sub
        End If

        If taskinfo IsNot Nothing Then
            For Each t As TContainer.TaskInfo In taskinfo
                Dim _JIRADiliveryTaskFill As New Thread(Sub() JIRADiliveryTaskFill(t, userinfo))
                If _JIRADiliveryTaskFill.IsAlive = False Then
                    _JIRADiliveryTaskFill = New Thread(Sub() JIRADiliveryTaskFill(t, userinfo))
                End If
                _JIRADiliveryTaskFill.Start()
            Next
        End If
    End Sub



    Private Sub CreateObservation(taskinfo As TContainer.TaskInfo, userinfo As TContainer.UserInfo)
        Dim errormsg As String = Nothing

        Dim taskid As New Information.TaskID(taskinfo.Task_ID)
        Dim modulename As New Information.CheckNull("Module name", taskinfo.ModuleName)
        Dim toolpath As New Information.CheckNull("Tool folder path", userinfo.Tfd)
        Dim modulepath As New Information.ModulePath(taskinfo.M_Path)
        Dim cquser As New Information.CheckNull("CQ User", userinfo.CQUser)
        Dim cqpassword As New Information.CheckNull("CQ Password", userinfo.CQPassword)

        Set_Status(CreateJIRAIssueBtn, False)
        Set_Content(OutputTextBox, "The tool is running, please be patient!" & vbNewLine)
        Dim CreateJIRAObservation As New Button.CreateJIRAObservation(taskid, cquser, cqpassword, modulename, toolpath, modulepath)
        errormsg = CreateJIRAObservation.Execute()
        If String.IsNullOrEmpty(errormsg) Then
            Set_Content(OutputTextBox, OutputTextBox.Text & vbNewLine & vbNewLine & "Finished!" & vbNewLine & vbNewLine & "Your JIRA ticket ID created: " & CreateJIRAObservation.JIRATicketID & vbNewLine & "Link: https://rb-tracker.bosch.com/tracker08/browse/" & CreateJIRAObservation.JIRATicketID)
        Else
            Set_Content(OutputTextBox, errormsg)
        End If
        Set_Status(CreateJIRAIssueBtn, True)
    End Sub

    Private Sub CreateObservationToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CreateObservationToolStripMenuItem.Click
        Dim taskinfo() As TContainer.TaskInfo = Nothing
        Dim userinfo As TContainer.UserInfo = Nothing
        Get_Info(taskinfo, userinfo)

        If taskinfo Is Nothing Or userinfo Is Nothing Then
            OutputTextBox.Text = "- No row selected."
            Exit Sub
        End If

        If taskinfo IsNot Nothing Then
            For Each t As TContainer.TaskInfo In taskinfo
                Dim _CreateJIRAObservation As New Thread(Sub() CreateObservation(t, userinfo))
                If _CreateJIRAObservation.IsAlive = False Then
                    _CreateJIRAObservation = New Thread(Sub() CreateObservation(t, userinfo))
                End If
                _CreateJIRAObservation.Start()
            Next
        End If
    End Sub

    Private Sub CreateOPL(taskinfo As TContainer.TaskInfo, userinfo As TContainer.UserInfo)
        Dim errormsg As String = Nothing

        Dim taskid As New Information.TaskID(taskinfo.Task_ID)
        Dim modulename As New Information.CheckNull("Module name", taskinfo.ModuleName)
        Dim modulepath As New Information.ModulePath(taskinfo.M_Path)
        Dim opllink As New Information.CheckNull("OPL Link", taskinfo.OPL_Link)
        Dim toolpath As New Information.CheckNull("Tool folder path", userinfo.Tfd)
        Dim cquser As New Information.CheckNull("CQ User", userinfo.CQUser)
        Dim cqpassword As New Information.CheckNull("CQ Password", userinfo.CQPassword)

        Set_Status(CreateJIRAIssueBtn, False)
        Set_Content(OutputTextBox, "The tool is running, please be patient!" & vbNewLine)
        Dim CreateJIRAOPL As New Button.CreateJIRAOPL(taskid, cquser, cqpassword, modulename, toolpath, modulepath, opllink)
        errormsg = CreateJIRAOPL.Execute()
        If String.IsNullOrEmpty(errormsg) Then
            Set_Content(OutputTextBox, OutputTextBox.Text & vbNewLine & vbNewLine & "Finished!" & vbNewLine & vbNewLine & "Your JIRA ticket ID created: " & CreateJIRAOPL.JIRATicketID & vbNewLine & "Link: https://rb-tracker.bosch.com/tracker08/browse/" & CreateJIRAOPL.JIRATicketID)
        Else
            Set_Content(OutputTextBox, errormsg)
        End If
        Set_Status(CreateJIRAIssueBtn, True)
    End Sub

    Private Sub CreateOPLToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CreateOPLToolStripMenuItem.Click
        Dim taskinfo() As TContainer.TaskInfo = Nothing
        Dim userinfo As TContainer.UserInfo = Nothing
        Get_Info(taskinfo, userinfo)

        If taskinfo Is Nothing Or userinfo Is Nothing Then
            OutputTextBox.Text = "- No row selected."
            Exit Sub
        End If

        If taskinfo IsNot Nothing Then
            For Each t As TContainer.TaskInfo In taskinfo
                Dim _CreateJIRAOPL As New Thread(Sub() CreateOPL(t, userinfo))
                If _CreateJIRAOPL.IsAlive = False Then
                    _CreateJIRAOPL = New Thread(Sub() CreateOPL(t, userinfo))
                End If
                _CreateJIRAOPL.Start()
            Next
        End If
    End Sub

    Private Sub CreateJIRADefect(taskinfo As TContainer.TaskInfo, userinfo As TContainer.UserInfo)
        Dim errormsg As String = Nothing

        Dim taskid As New Information.TaskID(taskinfo.Task_ID)
        Dim modulename As New Information.CheckNull("Module name", taskinfo.ModuleName)
        Dim toolpath As New Information.CheckNull("Tool folder path", userinfo.Tfd)
        Dim modulepath As New Information.ModulePath(taskinfo.M_Path)
        Dim cquser As New Information.CheckNull("CQ User", userinfo.CQUser)
        Dim cqpassword As New Information.CheckNull("CQ Password", userinfo.CQPassword)
        Dim hash As New Information.CheckNull("Hash", taskinfo.Revision)

        Set_Status(CreateJIRAIssueBtn, False)
        Set_Content(OutputTextBox, "The tool is running, please be patient!" & vbNewLine)
        Dim CreateJIRAObservation As New Button.CreateJIRADefect(taskid, cquser, cqpassword, modulename, toolpath, modulepath, hash)
        errormsg = CreateJIRAObservation.Execute()
        If String.IsNullOrEmpty(errormsg) Then
            Set_Content(OutputTextBox, OutputTextBox.Text & vbNewLine & vbNewLine & "Finished!" & vbNewLine & vbNewLine & "Your JIRA ticket ID created: " & CreateJIRAObservation.JIRATicketID & vbNewLine & "Link: https://rb-tracker.bosch.com/tracker08/browse/" & CreateJIRAObservation.JIRATicketID)
        Else
            Set_Content(OutputTextBox, errormsg)
        End If
        Set_Status(CreateJIRAIssueBtn, True)
    End Sub

    Private Sub CreateDefectToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CreateDefectToolStripMenuItem.Click
        Dim taskinfo() As TContainer.TaskInfo = Nothing
        Dim userinfo As TContainer.UserInfo = Nothing
        Get_Info(taskinfo, userinfo)

        If taskinfo Is Nothing Or userinfo Is Nothing Then
            OutputTextBox.Text = "- No row selected."
            Exit Sub
        End If

        If taskinfo IsNot Nothing Then
            For Each t As TContainer.TaskInfo In taskinfo
                Dim _CreateJIRADefect As New Thread(Sub() CreateJIRADefect(t, userinfo))
                If _CreateJIRADefect.IsAlive = False Then
                    _CreateJIRADefect = New Thread(Sub() CreateJIRADefect(t, userinfo))
                End If
                _CreateJIRADefect.Start()
            Next
        End If
    End Sub

    Private Sub CreateJIRADOU(taskinfo As TContainer.TaskInfo, userinfo As TContainer.UserInfo)
        Dim errormsg As String = Nothing

        Dim taskid As New Information.TaskID(taskinfo.Task_ID)
        Dim modulename As New Information.CheckNull("Module name", taskinfo.ModuleName)
        Dim ilm_link As New Information.CheckNull("ILM Link", taskinfo.ILM_Link)
        Dim modulepath As New Information.ModulePath(taskinfo.M_Path)
        Dim toolpath As New Information.CheckNull("Tool folder path", userinfo.Tfd)
        Dim cquser As New Information.CheckNull("CQ User", userinfo.CQUser)
        Dim cqpassword As New Information.CheckNull("CQ Password", userinfo.CQPassword)

        Set_Status(ReviewAssignBtn, False)
        Set_Content(OutputTextBox, "The tool is running, please be patient!" & vbNewLine)
        Dim CreateJIRADOU As New Button.CreateJIRADOU(taskid, cquser, cqpassword, modulename, toolpath, modulepath, ilm_link)
        errormsg = CreateJIRADOU.Execute()
        If String.IsNullOrEmpty(errormsg) Then
            Set_Content(OutputTextBox, OutputTextBox.Text & vbNewLine & vbNewLine & "Finished!" & vbNewLine & vbNewLine & "Your JIRA ticket ID created: " & CreateJIRADOU.JIRATicketID & vbNewLine & "Link: https://rb-tracker.bosch.com/tracker08/browse/" & CreateJIRADOU.JIRATicketID)
        Else
            Set_Content(OutputTextBox, errormsg)
        End If
        Set_Status(ReviewAssignBtn, True)
    End Sub


    Private Sub DOUReviewToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DOUReviewToolStripMenuItem.Click
        Dim taskinfo() As TContainer.TaskInfo = Nothing
        Dim userinfo As TContainer.UserInfo = Nothing
        Get_Info(taskinfo, userinfo)

        If taskinfo Is Nothing Or userinfo Is Nothing Then
            OutputTextBox.Text = "- No row selected."
            Exit Sub
        End If

        If taskinfo IsNot Nothing Then
            For Each t As TContainer.TaskInfo In taskinfo
                Dim _CreateJIRADOU As New Thread(Sub() CreateJIRADOU(t, userinfo))
                If _CreateJIRADOU.IsAlive = False Then
                    _CreateJIRADOU = New Thread(Sub() CreateJIRADOU(t, userinfo))
                End If
                _CreateJIRADOU.Start()
            Next
        End If
    End Sub

    Private Sub CreateJIRATC_TSReview(taskinfo As TContainer.TaskInfo, userinfo As TContainer.UserInfo)
        Dim errormsg As String = Nothing

        Dim taskid As New Information.TaskID(taskinfo.Task_ID)
        Dim modulename As New Information.CheckNull("Module name", taskinfo.ModuleName)
        Dim ilm_link As New Information.CheckNull("ILM Link", taskinfo.ILM_Link)
        Dim modulepath As New Information.ModulePath(taskinfo.M_Path)
        Dim toolpath As New Information.CheckNull("Tool folder path", userinfo.Tfd)
        Dim cquser As New Information.CheckNull("CQ User", userinfo.CQUser)
        Dim cqpassword As New Information.CheckNull("CQ Password", userinfo.CQPassword)

        Set_Status(ReviewAssignBtn, False)
        Set_Content(OutputTextBox, "The tool is running, please be patient!" & vbNewLine)
        Dim CreateJIRAReview As New Button.CreateJIRATC_TSReview(taskid, cquser, cqpassword, modulename, toolpath, modulepath, ilm_link)
        errormsg = CreateJIRAReview.Execute()
        If String.IsNullOrEmpty(errormsg) Then
            Set_Content(OutputTextBox, OutputTextBox.Text & vbNewLine & vbNewLine & "Finished!" & vbNewLine & vbNewLine & "Your JIRA ticket ID created: " & CreateJIRAReview.JIRATicketID & vbNewLine & "Link: https://rb-tracker.bosch.com/tracker08/browse/" & CreateJIRAReview.JIRATicketID)
        Else
            Set_Content(OutputTextBox, errormsg)
        End If
        Set_Status(ReviewAssignBtn, True)
    End Sub

    Private Sub TCTSReviewToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TCTSReviewToolStripMenuItem.Click
        Dim taskinfo() As TContainer.TaskInfo = Nothing
        Dim userinfo As TContainer.UserInfo = Nothing
        Get_Info(taskinfo, userinfo)

        If taskinfo Is Nothing Or userinfo Is Nothing Then
            OutputTextBox.Text = "- No row selected."
            Exit Sub
        End If

        If taskinfo IsNot Nothing Then
            For Each t As TContainer.TaskInfo In taskinfo
                Dim _CreateJIRATC_TSReview As New Thread(Sub() CreateJIRATC_TSReview(t, userinfo))
                If _CreateJIRATC_TSReview.IsAlive = False Then
                    _CreateJIRATC_TSReview = New Thread(Sub() CreateJIRATC_TSReview(t, userinfo))
                End If
                _CreateJIRATC_TSReview.Start()
            Next
        End If
    End Sub

    Private Sub PDCRequest(taskinfo As TContainer.TaskInfo, userinfo As TContainer.UserInfo)
        Dim errormsg As String = Nothing

        Dim taskid As New Information.TaskID(taskinfo.Task_ID)
        Dim modulename As New Information.CheckNull("Module name", taskinfo.ModuleName)
        Dim ilm_link As New Information.CheckNull("ILM Link", taskinfo.ILM_Link)
        Dim modulepath As New Information.ModulePath(taskinfo.M_Path)
        Dim toolpath As New Information.CheckNull("Tool folder path", userinfo.Tfd)
        Dim cquser As New Information.CheckNull("CQ User", userinfo.CQUser)
        Dim cqpassword As New Information.CheckNull("CQ Password", userinfo.CQPassword)

        Set_Status(ReviewAssignBtn, False)
        Set_Content(OutputTextBox, "The tool is running, please be patient!" & vbNewLine)
        Dim CreateJIRAPDC As New Button.CreateJIRAPDC(taskid, cquser, cqpassword, modulename, toolpath, modulepath, ilm_link)
        errormsg = CreateJIRAPDC.Execute()
        If String.IsNullOrEmpty(errormsg) Then
            Set_Content(OutputTextBox, OutputTextBox.Text & vbNewLine & vbNewLine & "Finished!" & vbNewLine & vbNewLine & "Your JIRA ticket ID created: " & CreateJIRAPDC.JIRATicketID & vbNewLine & "Link: https://rb-tracker.bosch.com/tracker08/browse/" & CreateJIRAPDC.JIRATicketID)
        Else
            Set_Content(OutputTextBox, errormsg)
        End If
        Set_Status(ReviewAssignBtn, True)
    End Sub

    Private Sub PDCRequestToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PDCRequestToolStripMenuItem.Click
        Dim taskinfo() As TContainer.TaskInfo = Nothing
        Dim userinfo As TContainer.UserInfo = Nothing
        Get_Info(taskinfo, userinfo)

        If taskinfo Is Nothing Or userinfo Is Nothing Then
            OutputTextBox.Text = "- No row selected."
            Exit Sub
        End If

        If taskinfo IsNot Nothing Then
            For Each t As TContainer.TaskInfo In taskinfo
                Dim _PDCRequest As New Thread(Sub() PDCRequest(t, userinfo))
                If _PDCRequest.IsAlive = False Then
                    _PDCRequest = New Thread(Sub() PDCRequest(t, userinfo))
                End If
                _PDCRequest.Start()
            Next
        End If
    End Sub

    Private Sub CantataExportScriptReport(taskinfo As TContainer.TaskInfo, userinfo As TContainer.UserInfo)
        Dim errormsg As String = Nothing

        Dim taskid As New Information.TaskID(taskinfo.Task_ID)
        Dim modulename As New Information.CheckNull("Module name", taskinfo.ModuleName)
        Dim ilm_link As New Information.CheckNull("ILM Link", taskinfo.ILM_Link)
        Dim modulepath As New Information.ModulePath(taskinfo.M_Path)
        Dim toolpath As New Information.CheckNull("Tool folder path", userinfo.Tfd)
        Dim cquser As New Information.CheckNull("CQ User", userinfo.CQUser)
        Dim cqpassword As New Information.CheckNull("CQ Password", userinfo.CQPassword)
        Dim wfd As New Information.CheckNull("Working Folder Path", userinfo.Wfd)
        Dim sfd As New Information.CheckNull("Sandbox Folder Path", userinfo.Sfd)
        Dim project As New Information.CheckNull("Project", taskinfo.Project)
        Dim sandbox As New Information.CheckNull("Sandbox", taskinfo.Sandbox)


        Set_Status(ExportScriptReportBtn, False)
        Set_Content(OutputTextBox, "The tool is running, please be patient!" & vbNewLine)
        Dim ExportScriptReport As New Button.ExportScriptReport(wfd, project, taskid, modulename, sfd, modulepath, sandbox)
        errormsg = ExportScriptReport.Execute()
        If String.IsNullOrEmpty(errormsg) Then
            Set_Content(OutputTextBox, OutputTextBox.Text & vbNewLine & vbNewLine & "Finished!")
        Else
            Set_Content(OutputTextBox, errormsg)
        End If
        Set_Status(ExportScriptReportBtn, True)
    End Sub

    Private Sub ExportScriptReport_Click(sender As Object, e As EventArgs) Handles ExportScriptReportBtn.Click
        Dim taskinfo() As TContainer.TaskInfo = Nothing
        Dim userinfo As TContainer.UserInfo = Nothing
        Get_Info(taskinfo, userinfo)

        If taskinfo Is Nothing Or userinfo Is Nothing Then
            OutputTextBox.Text = "- No row selected."
            Exit Sub
        End If

        If taskinfo IsNot Nothing Then
            For Each t As TContainer.TaskInfo In taskinfo
                Dim _CantataExportScriptReport As New Thread(Sub() CantataExportScriptReport(t, userinfo))
                If _CantataExportScriptReport.IsAlive = False Then
                    _CantataExportScriptReport = New Thread(Sub() CantataExportScriptReport(t, userinfo))
                End If
                _CantataExportScriptReport.Start()
            Next
        End If
    End Sub

    '***********************************-Upload OPL to local server-
    Public IsProject_ModelToConfig As Boolean
    Public Project_ModelToConfig As String

    Private Delegate Sub _Set_OPL_Link(ByRef taskinfo As TContainer.TaskInfo, ByRef GetOPLLink As Button.UploadOPLToServer)
    Public Sub Set_OPL_Link(ByRef taskinfo As TContainer.TaskInfo, ByRef GetOPLLink As Button.UploadOPLToServer)

        If ToolStrip1.InvokeRequired Then
            Dim d As New _Set_OPL_Link(AddressOf Set_OPL_Link)
            Invoke(d, New Object() {taskinfo, GetOPLLink})
        Else
            bsEnd()
            taskinfo.OPL_Link = GetOPLLink.OPLFullPath & "\" & UCase(taskinfo.ModuleName) & "_OPL.xls"
            bsEnd()
        End If
    End Sub

    Private Sub UpOPLToServer(taskinfo As TContainer.TaskInfo, userinfo As TContainer.UserInfo)
        Dim errormsg As String = Nothing

        Dim ExplorerPath As New Information.CheckNull("Explorer Dir", userinfo.ExplorerPath)
        Dim project As New Information.CheckNull("Project", taskinfo.Project)
        Dim taskid As New Information.TaskID(taskinfo.Task_ID)
        Dim modulename As New Information.CheckNull("Module Name", taskinfo.ModuleName)
        Dim modulepath As New Information.ModulePath(taskinfo.M_Path)

        Dim Wfd As New Information.CheckNull("Working Folder Dir", userinfo.Wfd)
        Dim jsonpath As New Information.CheckNull("Shared Path", userinfo.JsonPath)

        Set_Status(UpFilesToLocalServerBtn, False)
        Set_Content(OutputTextBox, "The tool is running, please be patient!" & vbNewLine)

        Dim ReadJsonLocalPath As New ReadJson.ReadJsonResultPath(project, jsonpath)
        Dim ResultPath As String = ReadJsonLocalPath.Execute()
        If ReadJsonLocalPath.IsValid = False Then
            OutputTextBox.Text = ReadJsonLocalPath.ErrorMsg
            Exit Sub
        End If

        If String.IsNullOrEmpty(ResultPath) Then
            Set_Content(OutputTextBox, "Can't find 'ResultPath' for " & project.GetValue & vbNewLine &
                       "Please click upload Button and update project infomation: " & vbNewLine)
            Set_Visibled(UpJsonFile, True)
            IsProject_ModelToConfig = True
            Project_ModelToConfig = project.GetValue
            Set_Status(UpFilesToLocalServerBtn, True)
            Exit Sub
        End If


        Dim UploadOPLToServer As New Button.UploadOPLToServer(ExplorerPath, Wfd, taskid, project, modulename, modulepath, ResultPath)
        errormsg = UploadOPLToServer.Execute()
        If String.IsNullOrEmpty(errormsg) Then
            Set_OPL_Link(taskinfo, UploadOPLToServer)
            Set_Content(OutputTextBox, OutputTextBox.Text & vbNewLine & vbNewLine & "Finished!")
        Else
            Set_Content(OutputTextBox, errormsg)
        End If
        Set_Status(UpFilesToLocalServerBtn, True)
    End Sub

    Private Sub UpOPLToServerToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UpOPLToServerToolStripMenuItem.Click
        Dim taskinfo() As TContainer.TaskInfo = Nothing
        Dim userinfo As TContainer.UserInfo = Nothing
        Get_Info(taskinfo, userinfo)

        If taskinfo Is Nothing Or userinfo Is Nothing Then
            OutputTextBox.Text = "- No row selected."
            Exit Sub
        End If

        If taskinfo IsNot Nothing Then
            For Each t As TContainer.TaskInfo In taskinfo
                Dim _UpOPLToServer As New Thread(Sub() UpOPLToServer(t, userinfo))
                If _UpOPLToServer.IsAlive = False Then
                    _UpOPLToServer = New Thread(Sub() UpOPLToServer(t, userinfo))
                End If
                _UpOPLToServer.Start()
            Next
        End If
    End Sub

    '***********************************-Upload Result to local server-
    Private Delegate Sub _Set_Result_Link(ByRef taskinfo As TContainer.TaskInfo, ByRef GetResultLink As Button.UploadResultToServer)
    Public Sub Set_Result_Link(ByRef taskinfo As TContainer.TaskInfo, ByRef GetResultLink As Button.UploadResultToServer)

        If ToolStrip1.InvokeRequired Then
            Dim d As New _Set_Result_Link(AddressOf Set_Result_Link)
            Invoke(d, New Object() {taskinfo, GetResultLink})
        Else
            bsEnd()
            taskinfo.Result_Path = GetResultLink.ResultFullPath
            bsEnd()
        End If
    End Sub

    Private Sub UpResultToServer(taskinfo As TContainer.TaskInfo, userinfo As TContainer.UserInfo)
        Dim errormsg As String = Nothing

        Dim ExplorerPath As New Information.CheckNull("Explorer Dir", userinfo.ExplorerPath)
        Dim project As New Information.CheckNull("Project", taskinfo.Project)
        Dim taskid As New Information.TaskID(taskinfo.Task_ID)
        Dim modulename As New Information.CheckNull("Module Name", taskinfo.ModuleName)
        Dim modulepath As New Information.ModulePath(taskinfo.M_Path)

        Dim Wfd As New Information.CheckNull("Working Folder Dir", userinfo.Wfd)
        Dim jsonpath As New Information.CheckNull("Shared Path", userinfo.JsonPath)

        Set_Status(UpFilesToLocalServerBtn, False)
        Set_Content(OutputTextBox, "The tool is running, please be patient!" & vbNewLine)

        Dim ReadJsonLocalPath As New ReadJson.ReadJsonResultPath(project, jsonpath)
        Dim ResultPath As String = ReadJsonLocalPath.Execute()
        If ReadJsonLocalPath.IsValid = False Then
            OutputTextBox.Text = ReadJsonLocalPath.ErrorMsg
            Exit Sub
        End If

        If String.IsNullOrEmpty(ResultPath) Then
            Set_Content(OutputTextBox, "Can't find 'ResultPath' for " & project.GetValue & vbNewLine &
                       "Please click upload Button and update project infomation: " & vbNewLine)
            Set_Visibled(UpJsonFile, True)
            IsProject_ModelToConfig = True
            Project_ModelToConfig = project.GetValue
            Set_Status(UpFilesToLocalServerBtn, True)
            Exit Sub
        End If


        Dim UploadResultToServer As New Button.UploadResultToServer(ExplorerPath, Wfd, taskid, project, modulename, modulepath, ResultPath)
        errormsg = UploadResultToServer.Execute()
        If String.IsNullOrEmpty(errormsg) Then
            Set_Result_Link(taskinfo, UploadResultToServer)
            Set_Content(OutputTextBox, OutputTextBox.Text & vbNewLine & vbNewLine & "Finished!")
        Else
            Set_Content(OutputTextBox, errormsg)
        End If
        Set_Status(UpFilesToLocalServerBtn, True)
    End Sub

    Private Sub UpResultToServerToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UpResultToServerToolStripMenuItem.Click
        Dim taskinfo() As TContainer.TaskInfo = Nothing
        Dim userinfo As TContainer.UserInfo = Nothing
        Get_Info(taskinfo, userinfo)

        If taskinfo Is Nothing Or userinfo Is Nothing Then
            OutputTextBox.Text = "- No row selected."
            Exit Sub
        End If

        If taskinfo IsNot Nothing Then
            For Each t As TContainer.TaskInfo In taskinfo
                Dim _UpResultToServer As New Thread(Sub() UpResultToServer(t, userinfo))
                If _UpResultToServer.IsAlive = False Then
                    _UpResultToServer = New Thread(Sub() UpResultToServer(t, userinfo))
                End If
                _UpResultToServer.Start()
            Next
        End If
    End Sub


    Private Sub UploadJIRATaskInfoBtn_ButtonClick(sender As Object, e As EventArgs) Handles UploadJIRATaskInfoBtn.ButtonClick
        UploadJIRATaskInfoBtn.ShowDropDown()
    End Sub

    Private Sub CreateJIRAIssueBtn_ButtonClick(sender As Object, e As EventArgs) Handles CreateJIRAIssueBtn.ButtonClick
        CreateJIRAIssueBtn.ShowDropDown()
    End Sub

    Private Sub ReviewAssignBtn_ButtonClick(sender As Object, e As EventArgs) Handles ReviewAssignBtn.ButtonClick
        ReviewAssignBtn.ShowDropDown()
    End Sub

    Private Sub FillDocumentBtn_ButtonClick(sender As Object, e As EventArgs) Handles FillDocumentBtn.ButtonClick
        FillDocumentBtn.ShowDropDown()
    End Sub

    Private Sub ToolStripSplitButton1_ButtonClick(sender As Object, e As EventArgs) Handles UpFilesToLocalServerBtn.ButtonClick
        UpFilesToLocalServerBtn.ShowDropDown()
    End Sub

    Private Sub UpJsonFile_Click(sender As Object, e As EventArgs) Handles UpJsonFile.Click
        Dim configf As New ConfigF
        configf.ShowDialog()
    End Sub

    Private Sub TaskReviewBtn_ButtonClick(sender As Object, e As EventArgs) Handles TaskReviewBtn.ButtonClick
        TaskReviewBtn.ShowDropDown()
    End Sub

    Private Sub ClearOutputBtn_Click(sender As Object, e As EventArgs) Handles ClearOutputBtn.Click
        OutputTextBox.Text = Nothing
    End Sub

    Private Sub TabControl1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TabControl1.SelectedIndexChanged
        If TabControl1.SelectedTab.Name = "TabPage1" Then
            ReviewerLabel.Text = "Reviewer"
            tex.init(Panel1, bs)
        ElseIf TabControl1.SelectedTab.Name = "TabPage2" Then
            ReviewerLabel.Text = "Author"
            tex.init(Panel1, bs2)
        ElseIf TabControl1.SelectedTab.Name = "TabPage3" Then
            tex.init(Panel1, bs3)
        End If

    End Sub
End Class
