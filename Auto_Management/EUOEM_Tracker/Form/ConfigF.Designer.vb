<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class ConfigF
    Inherits MetroFramework.Forms.MetroForm

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ConfigF))
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.TreeView1 = New System.Windows.Forms.TreeView()
        Me.savebtn = New MetroFramework.Controls.MetroButton()
        Me.cancelbtn = New MetroFramework.Controls.MetroButton()
        Me.MetroStyleManager1 = New MetroFramework.Components.MetroStyleManager(Me.components)
        Me.TabPage2 = New MetroFramework.Controls.MetroTabPage()
        Me.FindProBtn = New MetroFramework.Controls.MetroButton()
        Me.RemoveProBtn = New MetroFramework.Controls.MetroButton()
        Me.AddProBtn = New MetroFramework.Controls.MetroButton()
        Me.MetroLabel18 = New MetroFramework.Controls.MetroLabel()
        Me.ProjectCoorLb = New MetroFramework.Controls.MetroLabel()
        Me.ProjectLb = New MetroFramework.Controls.MetroLabel()
        Me.ProResultPathTextBox = New MetroFramework.Controls.MetroTextBox()
        Me.ProCoorTextBox = New MetroFramework.Controls.MetroTextBox()
        Me.ProModTextBox = New MetroFramework.Controls.MetroTextBox()
        Me.Grid1 = New MetroFramework.Controls.MetroGrid()
        Me.TabPage1 = New MetroFramework.Controls.MetroTabPage()
        Me.Stylecmb = New MetroFramework.Controls.MetroComboBox()
        Me.MetroLabel15 = New MetroFramework.Controls.MetroLabel()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.CQPasswordTextBox = New MetroFramework.Controls.MetroTextBox()
        Me.MetroLabel13 = New MetroFramework.Controls.MetroLabel()
        Me.CQUserTextBox = New MetroFramework.Controls.MetroTextBox()
        Me.MetroLabel14 = New MetroFramework.Controls.MetroLabel()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.JsonPathTextBox = New MetroFramework.Controls.MetroTextBox()
        Me.MetroLabel17 = New MetroFramework.Controls.MetroLabel()
        Me.FilesTemplateDirTextBox = New MetroFramework.Controls.MetroTextBox()
        Me.MetroLabel12 = New MetroFramework.Controls.MetroLabel()
        Me.MetroLabel8 = New MetroFramework.Controls.MetroLabel()
        Me.myPMTextBox = New MetroFramework.Controls.MetroTextBox()
        Me.teamleadTextBox = New MetroFramework.Controls.MetroTextBox()
        Me.MetroLabel9 = New MetroFramework.Controls.MetroLabel()
        Me.explorerTextBox = New MetroFramework.Controls.MetroTextBox()
        Me.MetroLabel10 = New MetroFramework.Controls.MetroLabel()
        Me.tfdTextBox = New MetroFramework.Controls.MetroTextBox()
        Me.MetroLabel5 = New MetroFramework.Controls.MetroLabel()
        Me.rfdTextBox = New MetroFramework.Controls.MetroTextBox()
        Me.MetroLabel4 = New MetroFramework.Controls.MetroLabel()
        Me.sfdTextBox = New MetroFramework.Controls.MetroTextBox()
        Me.MetroLabel3 = New MetroFramework.Controls.MetroLabel()
        Me.wfdTextBox = New MetroFramework.Controls.MetroTextBox()
        Me.MetroLabel2 = New MetroFramework.Controls.MetroLabel()
        Me.MyNameTextBox = New MetroFramework.Controls.MetroTextBox()
        Me.MetroLabel1 = New MetroFramework.Controls.MetroLabel()
        Me.TabControl1 = New MetroFramework.Controls.MetroTabControl()
        CType(Me.MetroStyleManager1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage2.SuspendLayout()
        CType(Me.Grid1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.SuspendLayout()
        '
        'TreeView1
        '
        Me.TreeView1.Location = New System.Drawing.Point(23, 63)
        Me.TreeView1.Name = "TreeView1"
        Me.TreeView1.Size = New System.Drawing.Size(134, 409)
        Me.TreeView1.TabIndex = 0
        '
        'savebtn
        '
        Me.savebtn.BackgroundImage = CType(resources.GetObject("savebtn.BackgroundImage"), System.Drawing.Image)
        Me.savebtn.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center
        Me.savebtn.Location = New System.Drawing.Point(580, 22)
        Me.savebtn.Name = "savebtn"
        Me.savebtn.Size = New System.Drawing.Size(50, 35)
        Me.savebtn.TabIndex = 2
        Me.savebtn.UseSelectable = True
        '
        'cancelbtn
        '
        Me.cancelbtn.BackgroundImage = CType(resources.GetObject("cancelbtn.BackgroundImage"), System.Drawing.Image)
        Me.cancelbtn.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center
        Me.cancelbtn.Location = New System.Drawing.Point(636, 22)
        Me.cancelbtn.Name = "cancelbtn"
        Me.cancelbtn.Size = New System.Drawing.Size(55, 35)
        Me.cancelbtn.TabIndex = 3
        Me.cancelbtn.UseSelectable = True
        '
        'MetroStyleManager1
        '
        Me.MetroStyleManager1.Owner = Nothing
        Me.MetroStyleManager1.Style = MetroFramework.MetroColorStyle.Teal
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.FindProBtn)
        Me.TabPage2.Controls.Add(Me.RemoveProBtn)
        Me.TabPage2.Controls.Add(Me.AddProBtn)
        Me.TabPage2.Controls.Add(Me.MetroLabel18)
        Me.TabPage2.Controls.Add(Me.ProjectCoorLb)
        Me.TabPage2.Controls.Add(Me.ProjectLb)
        Me.TabPage2.Controls.Add(Me.ProResultPathTextBox)
        Me.TabPage2.Controls.Add(Me.ProCoorTextBox)
        Me.TabPage2.Controls.Add(Me.ProModTextBox)
        Me.TabPage2.Controls.Add(Me.Grid1)
        Me.TabPage2.HorizontalScrollbarBarColor = True
        Me.TabPage2.HorizontalScrollbarHighlightOnWheel = False
        Me.TabPage2.HorizontalScrollbarSize = 10
        Me.TabPage2.Location = New System.Drawing.Point(4, 38)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Size = New System.Drawing.Size(580, 379)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "TabPage2"
        Me.TabPage2.VerticalScrollbarBarColor = True
        Me.TabPage2.VerticalScrollbarHighlightOnWheel = False
        Me.TabPage2.VerticalScrollbarSize = 10
        '
        'FindProBtn
        '
        Me.FindProBtn.BackgroundImage = CType(resources.GetObject("FindProBtn.BackgroundImage"), System.Drawing.Image)
        Me.FindProBtn.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center
        Me.FindProBtn.Location = New System.Drawing.Point(491, 335)
        Me.FindProBtn.Name = "FindProBtn"
        Me.FindProBtn.Size = New System.Drawing.Size(50, 35)
        Me.FindProBtn.TabIndex = 14
        Me.FindProBtn.UseSelectable = True
        '
        'RemoveProBtn
        '
        Me.RemoveProBtn.BackgroundImage = CType(resources.GetObject("RemoveProBtn.BackgroundImage"), System.Drawing.Image)
        Me.RemoveProBtn.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center
        Me.RemoveProBtn.Location = New System.Drawing.Point(491, 294)
        Me.RemoveProBtn.Name = "RemoveProBtn"
        Me.RemoveProBtn.Size = New System.Drawing.Size(50, 35)
        Me.RemoveProBtn.TabIndex = 13
        Me.RemoveProBtn.UseSelectable = True
        '
        'AddProBtn
        '
        Me.AddProBtn.BackgroundImage = CType(resources.GetObject("AddProBtn.BackgroundImage"), System.Drawing.Image)
        Me.AddProBtn.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center
        Me.AddProBtn.Location = New System.Drawing.Point(491, 253)
        Me.AddProBtn.Name = "AddProBtn"
        Me.AddProBtn.Size = New System.Drawing.Size(50, 35)
        Me.AddProBtn.TabIndex = 12
        Me.AddProBtn.UseSelectable = True
        '
        'MetroLabel18
        '
        Me.MetroLabel18.AutoSize = True
        Me.MetroLabel18.Location = New System.Drawing.Point(9, 312)
        Me.MetroLabel18.Name = "MetroLabel18"
        Me.MetroLabel18.Size = New System.Drawing.Size(118, 19)
        Me.MetroLabel18.TabIndex = 10
        Me.MetroLabel18.Text = "Project Result Path"
        '
        'ProjectCoorLb
        '
        Me.ProjectCoorLb.AutoSize = True
        Me.ProjectCoorLb.Location = New System.Drawing.Point(9, 283)
        Me.ProjectCoorLb.Name = "ProjectCoorLb"
        Me.ProjectCoorLb.Size = New System.Drawing.Size(126, 19)
        Me.ProjectCoorLb.TabIndex = 9
        Me.ProjectCoorLb.Text = "Project Coordinator"
        '
        'ProjectLb
        '
        Me.ProjectLb.AutoSize = True
        Me.ProjectLb.Location = New System.Drawing.Point(9, 253)
        Me.ProjectLb.Name = "ProjectLb"
        Me.ProjectLb.Size = New System.Drawing.Size(50, 19)
        Me.ProjectLb.TabIndex = 8
        Me.ProjectLb.Text = "Project"
        '
        'ProResultPathTextBox
        '
        '
        '
        '
        Me.ProResultPathTextBox.CustomButton.Image = Nothing
        Me.ProResultPathTextBox.CustomButton.Location = New System.Drawing.Point(287, 1)
        Me.ProResultPathTextBox.CustomButton.Name = ""
        Me.ProResultPathTextBox.CustomButton.Size = New System.Drawing.Size(21, 21)
        Me.ProResultPathTextBox.CustomButton.Style = MetroFramework.MetroColorStyle.Blue
        Me.ProResultPathTextBox.CustomButton.TabIndex = 1
        Me.ProResultPathTextBox.CustomButton.Theme = MetroFramework.MetroThemeStyle.Light
        Me.ProResultPathTextBox.CustomButton.UseSelectable = True
        Me.ProResultPathTextBox.CustomButton.Visible = False
        Me.ProResultPathTextBox.Lines = New String(-1) {}
        Me.ProResultPathTextBox.Location = New System.Drawing.Point(141, 312)
        Me.ProResultPathTextBox.MaxLength = 32767
        Me.ProResultPathTextBox.Name = "ProResultPathTextBox"
        Me.ProResultPathTextBox.PasswordChar = Global.Microsoft.VisualBasic.ChrW(0)
        Me.ProResultPathTextBox.ScrollBars = System.Windows.Forms.ScrollBars.None
        Me.ProResultPathTextBox.SelectedText = ""
        Me.ProResultPathTextBox.SelectionLength = 0
        Me.ProResultPathTextBox.SelectionStart = 0
        Me.ProResultPathTextBox.ShortcutsEnabled = True
        Me.ProResultPathTextBox.Size = New System.Drawing.Size(309, 23)
        Me.ProResultPathTextBox.Style = MetroFramework.MetroColorStyle.Teal
        Me.ProResultPathTextBox.TabIndex = 6
        Me.ProResultPathTextBox.UseSelectable = True
        Me.ProResultPathTextBox.WaterMarkColor = System.Drawing.Color.FromArgb(CType(CType(109, Byte), Integer), CType(CType(109, Byte), Integer), CType(CType(109, Byte), Integer))
        Me.ProResultPathTextBox.WaterMarkFont = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Pixel)
        '
        'ProCoorTextBox
        '
        '
        '
        '
        Me.ProCoorTextBox.CustomButton.Image = Nothing
        Me.ProCoorTextBox.CustomButton.Location = New System.Drawing.Point(287, 1)
        Me.ProCoorTextBox.CustomButton.Name = ""
        Me.ProCoorTextBox.CustomButton.Size = New System.Drawing.Size(21, 21)
        Me.ProCoorTextBox.CustomButton.Style = MetroFramework.MetroColorStyle.Blue
        Me.ProCoorTextBox.CustomButton.TabIndex = 1
        Me.ProCoorTextBox.CustomButton.Theme = MetroFramework.MetroThemeStyle.Light
        Me.ProCoorTextBox.CustomButton.UseSelectable = True
        Me.ProCoorTextBox.CustomButton.Visible = False
        Me.ProCoorTextBox.Lines = New String(-1) {}
        Me.ProCoorTextBox.Location = New System.Drawing.Point(141, 283)
        Me.ProCoorTextBox.MaxLength = 32767
        Me.ProCoorTextBox.Name = "ProCoorTextBox"
        Me.ProCoorTextBox.PasswordChar = Global.Microsoft.VisualBasic.ChrW(0)
        Me.ProCoorTextBox.ScrollBars = System.Windows.Forms.ScrollBars.None
        Me.ProCoorTextBox.SelectedText = ""
        Me.ProCoorTextBox.SelectionLength = 0
        Me.ProCoorTextBox.SelectionStart = 0
        Me.ProCoorTextBox.ShortcutsEnabled = True
        Me.ProCoorTextBox.Size = New System.Drawing.Size(309, 23)
        Me.ProCoorTextBox.Style = MetroFramework.MetroColorStyle.Teal
        Me.ProCoorTextBox.TabIndex = 5
        Me.ProCoorTextBox.UseSelectable = True
        Me.ProCoorTextBox.WaterMarkColor = System.Drawing.Color.FromArgb(CType(CType(109, Byte), Integer), CType(CType(109, Byte), Integer), CType(CType(109, Byte), Integer))
        Me.ProCoorTextBox.WaterMarkFont = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Pixel)
        '
        'ProModTextBox
        '
        '
        '
        '
        Me.ProModTextBox.CustomButton.Image = Nothing
        Me.ProModTextBox.CustomButton.Location = New System.Drawing.Point(287, 2)
        Me.ProModTextBox.CustomButton.Name = ""
        Me.ProModTextBox.CustomButton.Size = New System.Drawing.Size(19, 19)
        Me.ProModTextBox.CustomButton.Style = MetroFramework.MetroColorStyle.Blue
        Me.ProModTextBox.CustomButton.TabIndex = 1
        Me.ProModTextBox.CustomButton.Theme = MetroFramework.MetroThemeStyle.Light
        Me.ProModTextBox.CustomButton.UseSelectable = True
        Me.ProModTextBox.CustomButton.Visible = False
        Me.ProModTextBox.Lines = New String(-1) {}
        Me.ProModTextBox.Location = New System.Drawing.Point(141, 253)
        Me.ProModTextBox.MaxLength = 32767
        Me.ProModTextBox.Name = "ProModTextBox"
        Me.ProModTextBox.PasswordChar = Global.Microsoft.VisualBasic.ChrW(0)
        Me.ProModTextBox.ScrollBars = System.Windows.Forms.ScrollBars.None
        Me.ProModTextBox.SelectedText = ""
        Me.ProModTextBox.SelectionLength = 0
        Me.ProModTextBox.SelectionStart = 0
        Me.ProModTextBox.ShortcutsEnabled = True
        Me.ProModTextBox.Size = New System.Drawing.Size(309, 24)
        Me.ProModTextBox.Style = MetroFramework.MetroColorStyle.Teal
        Me.ProModTextBox.TabIndex = 4
        Me.ProModTextBox.UseSelectable = True
        Me.ProModTextBox.WaterMarkColor = System.Drawing.Color.FromArgb(CType(CType(109, Byte), Integer), CType(CType(109, Byte), Integer), CType(CType(109, Byte), Integer))
        Me.ProModTextBox.WaterMarkFont = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Pixel)
        '
        'Grid1
        '
        Me.Grid1.AllowUserToResizeRows = False
        Me.Grid1.BackgroundColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.Grid1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.Grid1.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.None
        Me.Grid1.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(170, Byte), Integer), CType(CType(173, Byte), Integer))
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold)
        DataGridViewCellStyle1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(201, Byte), Integer), CType(CType(206, Byte), Integer))
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.Color.FromArgb(CType(CType(17, Byte), Integer), CType(CType(17, Byte), Integer), CType(CType(17, Byte), Integer))
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Grid1.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.Grid1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Segoe UI", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel)
        DataGridViewCellStyle2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(201, Byte), Integer), CType(CType(206, Byte), Integer))
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.Color.FromArgb(CType(CType(17, Byte), Integer), CType(CType(17, Byte), Integer), CType(CType(17, Byte), Integer))
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.Grid1.DefaultCellStyle = DataGridViewCellStyle2
        Me.Grid1.EnableHeadersVisualStyles = False
        Me.Grid1.Font = New System.Drawing.Font("Segoe UI", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel)
        Me.Grid1.GridColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.Grid1.Location = New System.Drawing.Point(1, 0)
        Me.Grid1.Name = "Grid1"
        Me.Grid1.ReadOnly = True
        Me.Grid1.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle3.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(170, Byte), Integer), CType(CType(173, Byte), Integer))
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Segoe UI", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel)
        DataGridViewCellStyle3.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(201, Byte), Integer), CType(CType(206, Byte), Integer))
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.Color.FromArgb(CType(CType(17, Byte), Integer), CType(CType(17, Byte), Integer), CType(CType(17, Byte), Integer))
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Grid1.RowHeadersDefaultCellStyle = DataGridViewCellStyle3
        Me.Grid1.RowHeadersVisible = False
        Me.Grid1.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.Grid1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.Grid1.Size = New System.Drawing.Size(583, 247)
        Me.Grid1.Style = MetroFramework.MetroColorStyle.Teal
        Me.Grid1.TabIndex = 3
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.Stylecmb)
        Me.TabPage1.Controls.Add(Me.MetroLabel15)
        Me.TabPage1.Controls.Add(Me.GroupBox2)
        Me.TabPage1.Controls.Add(Me.GroupBox1)
        Me.TabPage1.HorizontalScrollbarBarColor = True
        Me.TabPage1.HorizontalScrollbarHighlightOnWheel = False
        Me.TabPage1.HorizontalScrollbarSize = 10
        Me.TabPage1.Location = New System.Drawing.Point(4, 38)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Size = New System.Drawing.Size(580, 379)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "TabPage1"
        Me.TabPage1.VerticalScrollbarBarColor = True
        Me.TabPage1.VerticalScrollbarHighlightOnWheel = False
        Me.TabPage1.VerticalScrollbarSize = 10
        '
        'Stylecmb
        '
        Me.Stylecmb.FormattingEnabled = True
        Me.Stylecmb.ItemHeight = 23
        Me.Stylecmb.Location = New System.Drawing.Point(382, 295)
        Me.Stylecmb.Name = "Stylecmb"
        Me.Stylecmb.Size = New System.Drawing.Size(125, 29)
        Me.Stylecmb.TabIndex = 5
        Me.Stylecmb.UseSelectable = True
        '
        'MetroLabel15
        '
        Me.MetroLabel15.AutoSize = True
        Me.MetroLabel15.Location = New System.Drawing.Point(326, 295)
        Me.MetroLabel15.Name = "MetroLabel15"
        Me.MetroLabel15.Size = New System.Drawing.Size(36, 19)
        Me.MetroLabel15.TabIndex = 4
        Me.MetroLabel15.Text = "Style"
        '
        'GroupBox2
        '
        Me.GroupBox2.BackColor = System.Drawing.Color.White
        Me.GroupBox2.Controls.Add(Me.CQPasswordTextBox)
        Me.GroupBox2.Controls.Add(Me.MetroLabel13)
        Me.GroupBox2.Controls.Add(Me.CQUserTextBox)
        Me.GroupBox2.Controls.Add(Me.MetroLabel14)
        Me.GroupBox2.Location = New System.Drawing.Point(3, 261)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(250, 110)
        Me.GroupBox2.TabIndex = 3
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Account"
        '
        'CQPasswordTextBox
        '
        '
        '
        '
        Me.CQPasswordTextBox.CustomButton.Image = Nothing
        Me.CQPasswordTextBox.CustomButton.Location = New System.Drawing.Point(125, 2)
        Me.CQPasswordTextBox.CustomButton.Name = ""
        Me.CQPasswordTextBox.CustomButton.Size = New System.Drawing.Size(23, 23)
        Me.CQPasswordTextBox.CustomButton.Style = MetroFramework.MetroColorStyle.Blue
        Me.CQPasswordTextBox.CustomButton.TabIndex = 1
        Me.CQPasswordTextBox.CustomButton.Theme = MetroFramework.MetroThemeStyle.Light
        Me.CQPasswordTextBox.CustomButton.UseSelectable = True
        Me.CQPasswordTextBox.CustomButton.Visible = False
        Me.CQPasswordTextBox.Lines = New String(-1) {}
        Me.CQPasswordTextBox.Location = New System.Drawing.Point(88, 63)
        Me.CQPasswordTextBox.MaxLength = 32767
        Me.CQPasswordTextBox.Name = "CQPasswordTextBox"
        Me.CQPasswordTextBox.PasswordChar = Global.Microsoft.VisualBasic.ChrW(9679)
        Me.CQPasswordTextBox.ScrollBars = System.Windows.Forms.ScrollBars.None
        Me.CQPasswordTextBox.SelectedText = ""
        Me.CQPasswordTextBox.SelectionLength = 0
        Me.CQPasswordTextBox.SelectionStart = 0
        Me.CQPasswordTextBox.ShortcutsEnabled = True
        Me.CQPasswordTextBox.Size = New System.Drawing.Size(151, 28)
        Me.CQPasswordTextBox.TabIndex = 7
        Me.CQPasswordTextBox.UseSelectable = True
        Me.CQPasswordTextBox.UseSystemPasswordChar = True
        Me.CQPasswordTextBox.WaterMarkColor = System.Drawing.Color.FromArgb(CType(CType(109, Byte), Integer), CType(CType(109, Byte), Integer), CType(CType(109, Byte), Integer))
        Me.CQPasswordTextBox.WaterMarkFont = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel)
        '
        'MetroLabel13
        '
        Me.MetroLabel13.AutoSize = True
        Me.MetroLabel13.Location = New System.Drawing.Point(15, 63)
        Me.MetroLabel13.Name = "MetroLabel13"
        Me.MetroLabel13.Size = New System.Drawing.Size(64, 19)
        Me.MetroLabel13.TabIndex = 6
        Me.MetroLabel13.Text = "Password"
        '
        'CQUserTextBox
        '
        '
        '
        '
        Me.CQUserTextBox.CustomButton.Image = Nothing
        Me.CQUserTextBox.CustomButton.Location = New System.Drawing.Point(125, 2)
        Me.CQUserTextBox.CustomButton.Name = ""
        Me.CQUserTextBox.CustomButton.Size = New System.Drawing.Size(23, 23)
        Me.CQUserTextBox.CustomButton.Style = MetroFramework.MetroColorStyle.Blue
        Me.CQUserTextBox.CustomButton.TabIndex = 1
        Me.CQUserTextBox.CustomButton.Theme = MetroFramework.MetroThemeStyle.Light
        Me.CQUserTextBox.CustomButton.UseSelectable = True
        Me.CQUserTextBox.CustomButton.Visible = False
        Me.CQUserTextBox.Lines = New String(-1) {}
        Me.CQUserTextBox.Location = New System.Drawing.Point(88, 34)
        Me.CQUserTextBox.MaxLength = 32767
        Me.CQUserTextBox.Name = "CQUserTextBox"
        Me.CQUserTextBox.PasswordChar = Global.Microsoft.VisualBasic.ChrW(0)
        Me.CQUserTextBox.ScrollBars = System.Windows.Forms.ScrollBars.None
        Me.CQUserTextBox.SelectedText = ""
        Me.CQUserTextBox.SelectionLength = 0
        Me.CQUserTextBox.SelectionStart = 0
        Me.CQUserTextBox.ShortcutsEnabled = True
        Me.CQUserTextBox.Size = New System.Drawing.Size(151, 28)
        Me.CQUserTextBox.TabIndex = 5
        Me.CQUserTextBox.UseSelectable = True
        Me.CQUserTextBox.WaterMarkColor = System.Drawing.Color.FromArgb(CType(CType(109, Byte), Integer), CType(CType(109, Byte), Integer), CType(CType(109, Byte), Integer))
        Me.CQUserTextBox.WaterMarkFont = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel)
        '
        'MetroLabel14
        '
        Me.MetroLabel14.AutoSize = True
        Me.MetroLabel14.Location = New System.Drawing.Point(15, 34)
        Me.MetroLabel14.Name = "MetroLabel14"
        Me.MetroLabel14.Size = New System.Drawing.Size(35, 19)
        Me.MetroLabel14.TabIndex = 4
        Me.MetroLabel14.Text = "User"
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.White
        Me.GroupBox1.Controls.Add(Me.JsonPathTextBox)
        Me.GroupBox1.Controls.Add(Me.MetroLabel17)
        Me.GroupBox1.Controls.Add(Me.FilesTemplateDirTextBox)
        Me.GroupBox1.Controls.Add(Me.MetroLabel12)
        Me.GroupBox1.Controls.Add(Me.MetroLabel8)
        Me.GroupBox1.Controls.Add(Me.myPMTextBox)
        Me.GroupBox1.Controls.Add(Me.teamleadTextBox)
        Me.GroupBox1.Controls.Add(Me.MetroLabel9)
        Me.GroupBox1.Controls.Add(Me.explorerTextBox)
        Me.GroupBox1.Controls.Add(Me.MetroLabel10)
        Me.GroupBox1.Controls.Add(Me.tfdTextBox)
        Me.GroupBox1.Controls.Add(Me.MetroLabel5)
        Me.GroupBox1.Controls.Add(Me.rfdTextBox)
        Me.GroupBox1.Controls.Add(Me.MetroLabel4)
        Me.GroupBox1.Controls.Add(Me.sfdTextBox)
        Me.GroupBox1.Controls.Add(Me.MetroLabel3)
        Me.GroupBox1.Controls.Add(Me.wfdTextBox)
        Me.GroupBox1.Controls.Add(Me.MetroLabel2)
        Me.GroupBox1.Controls.Add(Me.MyNameTextBox)
        Me.GroupBox1.Controls.Add(Me.MetroLabel1)
        Me.GroupBox1.Location = New System.Drawing.Point(3, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(525, 243)
        Me.GroupBox1.TabIndex = 2
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "User"
        '
        'JsonPathTextBox
        '
        '
        '
        '
        Me.JsonPathTextBox.CustomButton.Image = Nothing
        Me.JsonPathTextBox.CustomButton.Location = New System.Drawing.Point(99, 2)
        Me.JsonPathTextBox.CustomButton.Name = ""
        Me.JsonPathTextBox.CustomButton.Size = New System.Drawing.Size(23, 23)
        Me.JsonPathTextBox.CustomButton.Style = MetroFramework.MetroColorStyle.Blue
        Me.JsonPathTextBox.CustomButton.TabIndex = 1
        Me.JsonPathTextBox.CustomButton.Theme = MetroFramework.MetroThemeStyle.Light
        Me.JsonPathTextBox.CustomButton.UseSelectable = True
        Me.JsonPathTextBox.CustomButton.Visible = False
        Me.JsonPathTextBox.Lines = New String(-1) {}
        Me.JsonPathTextBox.Location = New System.Drawing.Point(379, 112)
        Me.JsonPathTextBox.MaxLength = 32767
        Me.JsonPathTextBox.Name = "JsonPathTextBox"
        Me.JsonPathTextBox.PasswordChar = Global.Microsoft.VisualBasic.ChrW(0)
        Me.JsonPathTextBox.ScrollBars = System.Windows.Forms.ScrollBars.None
        Me.JsonPathTextBox.SelectedText = ""
        Me.JsonPathTextBox.SelectionLength = 0
        Me.JsonPathTextBox.SelectionStart = 0
        Me.JsonPathTextBox.ShortcutsEnabled = True
        Me.JsonPathTextBox.Size = New System.Drawing.Size(125, 28)
        Me.JsonPathTextBox.TabIndex = 35
        Me.JsonPathTextBox.UseSelectable = True
        Me.JsonPathTextBox.WaterMarkColor = System.Drawing.Color.FromArgb(CType(CType(109, Byte), Integer), CType(CType(109, Byte), Integer), CType(CType(109, Byte), Integer))
        Me.JsonPathTextBox.WaterMarkFont = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel)
        '
        'MetroLabel17
        '
        Me.MetroLabel17.AutoSize = True
        Me.MetroLabel17.Location = New System.Drawing.Point(299, 121)
        Me.MetroLabel17.Name = "MetroLabel17"
        Me.MetroLabel17.Size = New System.Drawing.Size(71, 19)
        Me.MetroLabel17.TabIndex = 34
        Me.MetroLabel17.Text = "Shared Dir"
        '
        'FilesTemplateDirTextBox
        '
        '
        '
        '
        Me.FilesTemplateDirTextBox.CustomButton.Image = Nothing
        Me.FilesTemplateDirTextBox.CustomButton.Location = New System.Drawing.Point(99, 2)
        Me.FilesTemplateDirTextBox.CustomButton.Name = ""
        Me.FilesTemplateDirTextBox.CustomButton.Size = New System.Drawing.Size(23, 23)
        Me.FilesTemplateDirTextBox.CustomButton.Style = MetroFramework.MetroColorStyle.Blue
        Me.FilesTemplateDirTextBox.CustomButton.TabIndex = 1
        Me.FilesTemplateDirTextBox.CustomButton.Theme = MetroFramework.MetroThemeStyle.Light
        Me.FilesTemplateDirTextBox.CustomButton.UseSelectable = True
        Me.FilesTemplateDirTextBox.CustomButton.Visible = False
        Me.FilesTemplateDirTextBox.Lines = New String(-1) {}
        Me.FilesTemplateDirTextBox.Location = New System.Drawing.Point(379, 83)
        Me.FilesTemplateDirTextBox.MaxLength = 32767
        Me.FilesTemplateDirTextBox.Name = "FilesTemplateDirTextBox"
        Me.FilesTemplateDirTextBox.PasswordChar = Global.Microsoft.VisualBasic.ChrW(0)
        Me.FilesTemplateDirTextBox.ScrollBars = System.Windows.Forms.ScrollBars.None
        Me.FilesTemplateDirTextBox.SelectedText = ""
        Me.FilesTemplateDirTextBox.SelectionLength = 0
        Me.FilesTemplateDirTextBox.SelectionStart = 0
        Me.FilesTemplateDirTextBox.ShortcutsEnabled = True
        Me.FilesTemplateDirTextBox.Size = New System.Drawing.Size(125, 28)
        Me.FilesTemplateDirTextBox.TabIndex = 33
        Me.FilesTemplateDirTextBox.UseSelectable = True
        Me.FilesTemplateDirTextBox.WaterMarkColor = System.Drawing.Color.FromArgb(CType(CType(109, Byte), Integer), CType(CType(109, Byte), Integer), CType(CType(109, Byte), Integer))
        Me.FilesTemplateDirTextBox.WaterMarkFont = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel)
        '
        'MetroLabel12
        '
        Me.MetroLabel12.AutoSize = True
        Me.MetroLabel12.Location = New System.Drawing.Point(299, 92)
        Me.MetroLabel12.Name = "MetroLabel12"
        Me.MetroLabel12.Size = New System.Drawing.Size(64, 19)
        Me.MetroLabel12.TabIndex = 32
        Me.MetroLabel12.Text = "Temp Dir"
        '
        'MetroLabel8
        '
        Me.MetroLabel8.AutoSize = True
        Me.MetroLabel8.Location = New System.Drawing.Point(299, 63)
        Me.MetroLabel8.Name = "MetroLabel8"
        Me.MetroLabel8.Size = New System.Drawing.Size(51, 19)
        Me.MetroLabel8.TabIndex = 29
        Me.MetroLabel8.Text = "My PM"
        '
        'myPMTextBox
        '
        '
        '
        '
        Me.myPMTextBox.CustomButton.Image = Nothing
        Me.myPMTextBox.CustomButton.Location = New System.Drawing.Point(99, 2)
        Me.myPMTextBox.CustomButton.Name = ""
        Me.myPMTextBox.CustomButton.Size = New System.Drawing.Size(23, 23)
        Me.myPMTextBox.CustomButton.Style = MetroFramework.MetroColorStyle.Blue
        Me.myPMTextBox.CustomButton.TabIndex = 1
        Me.myPMTextBox.CustomButton.Theme = MetroFramework.MetroThemeStyle.Light
        Me.myPMTextBox.CustomButton.UseSelectable = True
        Me.myPMTextBox.CustomButton.Visible = False
        Me.myPMTextBox.Lines = New String(-1) {}
        Me.myPMTextBox.Location = New System.Drawing.Point(379, 54)
        Me.myPMTextBox.MaxLength = 32767
        Me.myPMTextBox.Name = "myPMTextBox"
        Me.myPMTextBox.PasswordChar = Global.Microsoft.VisualBasic.ChrW(0)
        Me.myPMTextBox.ScrollBars = System.Windows.Forms.ScrollBars.None
        Me.myPMTextBox.SelectedText = ""
        Me.myPMTextBox.SelectionLength = 0
        Me.myPMTextBox.SelectionStart = 0
        Me.myPMTextBox.ShortcutsEnabled = True
        Me.myPMTextBox.Size = New System.Drawing.Size(125, 28)
        Me.myPMTextBox.TabIndex = 28
        Me.myPMTextBox.UseSelectable = True
        Me.myPMTextBox.WaterMarkColor = System.Drawing.Color.FromArgb(CType(CType(109, Byte), Integer), CType(CType(109, Byte), Integer), CType(CType(109, Byte), Integer))
        Me.myPMTextBox.WaterMarkFont = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel)
        '
        'teamleadTextBox
        '
        '
        '
        '
        Me.teamleadTextBox.CustomButton.Image = Nothing
        Me.teamleadTextBox.CustomButton.Location = New System.Drawing.Point(99, 2)
        Me.teamleadTextBox.CustomButton.Name = ""
        Me.teamleadTextBox.CustomButton.Size = New System.Drawing.Size(23, 23)
        Me.teamleadTextBox.CustomButton.Style = MetroFramework.MetroColorStyle.Blue
        Me.teamleadTextBox.CustomButton.TabIndex = 1
        Me.teamleadTextBox.CustomButton.Theme = MetroFramework.MetroThemeStyle.Light
        Me.teamleadTextBox.CustomButton.UseSelectable = True
        Me.teamleadTextBox.CustomButton.Visible = False
        Me.teamleadTextBox.Lines = New String(-1) {}
        Me.teamleadTextBox.Location = New System.Drawing.Point(379, 25)
        Me.teamleadTextBox.MaxLength = 32767
        Me.teamleadTextBox.Name = "teamleadTextBox"
        Me.teamleadTextBox.PasswordChar = Global.Microsoft.VisualBasic.ChrW(0)
        Me.teamleadTextBox.ScrollBars = System.Windows.Forms.ScrollBars.None
        Me.teamleadTextBox.SelectedText = ""
        Me.teamleadTextBox.SelectionLength = 0
        Me.teamleadTextBox.SelectionStart = 0
        Me.teamleadTextBox.ShortcutsEnabled = True
        Me.teamleadTextBox.Size = New System.Drawing.Size(125, 28)
        Me.teamleadTextBox.TabIndex = 27
        Me.teamleadTextBox.UseSelectable = True
        Me.teamleadTextBox.WaterMarkColor = System.Drawing.Color.FromArgb(CType(CType(109, Byte), Integer), CType(CType(109, Byte), Integer), CType(CType(109, Byte), Integer))
        Me.teamleadTextBox.WaterMarkFont = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel)
        '
        'MetroLabel9
        '
        Me.MetroLabel9.AutoSize = True
        Me.MetroLabel9.Location = New System.Drawing.Point(299, 34)
        Me.MetroLabel9.Name = "MetroLabel9"
        Me.MetroLabel9.Size = New System.Drawing.Size(74, 19)
        Me.MetroLabel9.TabIndex = 26
        Me.MetroLabel9.Text = "Team Lead"
        '
        'explorerTextBox
        '
        '
        '
        '
        Me.explorerTextBox.CustomButton.Image = Nothing
        Me.explorerTextBox.CustomButton.Location = New System.Drawing.Point(119, 2)
        Me.explorerTextBox.CustomButton.Name = ""
        Me.explorerTextBox.CustomButton.Size = New System.Drawing.Size(23, 23)
        Me.explorerTextBox.CustomButton.Style = MetroFramework.MetroColorStyle.Blue
        Me.explorerTextBox.CustomButton.TabIndex = 1
        Me.explorerTextBox.CustomButton.Theme = MetroFramework.MetroThemeStyle.Light
        Me.explorerTextBox.CustomButton.UseSelectable = True
        Me.explorerTextBox.CustomButton.Visible = False
        Me.explorerTextBox.Lines = New String(-1) {}
        Me.explorerTextBox.Location = New System.Drawing.Point(143, 170)
        Me.explorerTextBox.MaxLength = 32767
        Me.explorerTextBox.Name = "explorerTextBox"
        Me.explorerTextBox.PasswordChar = Global.Microsoft.VisualBasic.ChrW(0)
        Me.explorerTextBox.ScrollBars = System.Windows.Forms.ScrollBars.None
        Me.explorerTextBox.SelectedText = ""
        Me.explorerTextBox.SelectionLength = 0
        Me.explorerTextBox.SelectionStart = 0
        Me.explorerTextBox.ShortcutsEnabled = True
        Me.explorerTextBox.Size = New System.Drawing.Size(145, 28)
        Me.explorerTextBox.TabIndex = 25
        Me.explorerTextBox.UseSelectable = True
        Me.explorerTextBox.WaterMarkColor = System.Drawing.Color.FromArgb(CType(CType(109, Byte), Integer), CType(CType(109, Byte), Integer), CType(CType(109, Byte), Integer))
        Me.explorerTextBox.WaterMarkFont = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel)
        '
        'MetroLabel10
        '
        Me.MetroLabel10.AutoSize = True
        Me.MetroLabel10.Location = New System.Drawing.Point(15, 170)
        Me.MetroLabel10.Name = "MetroLabel10"
        Me.MetroLabel10.Size = New System.Drawing.Size(58, 19)
        Me.MetroLabel10.TabIndex = 16
        Me.MetroLabel10.Text = "Explorer"
        '
        'tfdTextBox
        '
        '
        '
        '
        Me.tfdTextBox.CustomButton.Image = Nothing
        Me.tfdTextBox.CustomButton.Location = New System.Drawing.Point(119, 2)
        Me.tfdTextBox.CustomButton.Name = ""
        Me.tfdTextBox.CustomButton.Size = New System.Drawing.Size(23, 23)
        Me.tfdTextBox.CustomButton.Style = MetroFramework.MetroColorStyle.Blue
        Me.tfdTextBox.CustomButton.TabIndex = 1
        Me.tfdTextBox.CustomButton.Theme = MetroFramework.MetroThemeStyle.Light
        Me.tfdTextBox.CustomButton.UseSelectable = True
        Me.tfdTextBox.CustomButton.Visible = False
        Me.tfdTextBox.Lines = New String(-1) {}
        Me.tfdTextBox.Location = New System.Drawing.Point(143, 141)
        Me.tfdTextBox.MaxLength = 32767
        Me.tfdTextBox.Name = "tfdTextBox"
        Me.tfdTextBox.PasswordChar = Global.Microsoft.VisualBasic.ChrW(0)
        Me.tfdTextBox.ScrollBars = System.Windows.Forms.ScrollBars.None
        Me.tfdTextBox.SelectedText = ""
        Me.tfdTextBox.SelectionLength = 0
        Me.tfdTextBox.SelectionStart = 0
        Me.tfdTextBox.ShortcutsEnabled = True
        Me.tfdTextBox.Size = New System.Drawing.Size(145, 28)
        Me.tfdTextBox.TabIndex = 9
        Me.tfdTextBox.UseSelectable = True
        Me.tfdTextBox.WaterMarkColor = System.Drawing.Color.FromArgb(CType(CType(109, Byte), Integer), CType(CType(109, Byte), Integer), CType(CType(109, Byte), Integer))
        Me.tfdTextBox.WaterMarkFont = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel)
        '
        'MetroLabel5
        '
        Me.MetroLabel5.AutoSize = True
        Me.MetroLabel5.Location = New System.Drawing.Point(15, 141)
        Me.MetroLabel5.Name = "MetroLabel5"
        Me.MetroLabel5.Size = New System.Drawing.Size(98, 19)
        Me.MetroLabel5.TabIndex = 8
        Me.MetroLabel5.Text = "Tool Folder Dir"
        '
        'rfdTextBox
        '
        '
        '
        '
        Me.rfdTextBox.CustomButton.Image = Nothing
        Me.rfdTextBox.CustomButton.Location = New System.Drawing.Point(119, 2)
        Me.rfdTextBox.CustomButton.Name = ""
        Me.rfdTextBox.CustomButton.Size = New System.Drawing.Size(23, 23)
        Me.rfdTextBox.CustomButton.Style = MetroFramework.MetroColorStyle.Blue
        Me.rfdTextBox.CustomButton.TabIndex = 1
        Me.rfdTextBox.CustomButton.Theme = MetroFramework.MetroThemeStyle.Light
        Me.rfdTextBox.CustomButton.UseSelectable = True
        Me.rfdTextBox.CustomButton.Visible = False
        Me.rfdTextBox.Lines = New String(-1) {}
        Me.rfdTextBox.Location = New System.Drawing.Point(143, 112)
        Me.rfdTextBox.MaxLength = 32767
        Me.rfdTextBox.Name = "rfdTextBox"
        Me.rfdTextBox.PasswordChar = Global.Microsoft.VisualBasic.ChrW(0)
        Me.rfdTextBox.ScrollBars = System.Windows.Forms.ScrollBars.None
        Me.rfdTextBox.SelectedText = ""
        Me.rfdTextBox.SelectionLength = 0
        Me.rfdTextBox.SelectionStart = 0
        Me.rfdTextBox.ShortcutsEnabled = True
        Me.rfdTextBox.Size = New System.Drawing.Size(145, 28)
        Me.rfdTextBox.TabIndex = 7
        Me.rfdTextBox.UseSelectable = True
        Me.rfdTextBox.WaterMarkColor = System.Drawing.Color.FromArgb(CType(CType(109, Byte), Integer), CType(CType(109, Byte), Integer), CType(CType(109, Byte), Integer))
        Me.rfdTextBox.WaterMarkFont = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel)
        '
        'MetroLabel4
        '
        Me.MetroLabel4.AutoSize = True
        Me.MetroLabel4.Location = New System.Drawing.Point(15, 112)
        Me.MetroLabel4.Name = "MetroLabel4"
        Me.MetroLabel4.Size = New System.Drawing.Size(112, 19)
        Me.MetroLabel4.TabIndex = 6
        Me.MetroLabel4.Text = "Review Folder Dir"
        '
        'sfdTextBox
        '
        '
        '
        '
        Me.sfdTextBox.CustomButton.Image = Nothing
        Me.sfdTextBox.CustomButton.Location = New System.Drawing.Point(119, 2)
        Me.sfdTextBox.CustomButton.Name = ""
        Me.sfdTextBox.CustomButton.Size = New System.Drawing.Size(23, 23)
        Me.sfdTextBox.CustomButton.Style = MetroFramework.MetroColorStyle.Blue
        Me.sfdTextBox.CustomButton.TabIndex = 1
        Me.sfdTextBox.CustomButton.Theme = MetroFramework.MetroThemeStyle.Light
        Me.sfdTextBox.CustomButton.UseSelectable = True
        Me.sfdTextBox.CustomButton.Visible = False
        Me.sfdTextBox.Lines = New String(-1) {}
        Me.sfdTextBox.Location = New System.Drawing.Point(143, 83)
        Me.sfdTextBox.MaxLength = 32767
        Me.sfdTextBox.Name = "sfdTextBox"
        Me.sfdTextBox.PasswordChar = Global.Microsoft.VisualBasic.ChrW(0)
        Me.sfdTextBox.ScrollBars = System.Windows.Forms.ScrollBars.None
        Me.sfdTextBox.SelectedText = ""
        Me.sfdTextBox.SelectionLength = 0
        Me.sfdTextBox.SelectionStart = 0
        Me.sfdTextBox.ShortcutsEnabled = True
        Me.sfdTextBox.Size = New System.Drawing.Size(145, 28)
        Me.sfdTextBox.TabIndex = 5
        Me.sfdTextBox.UseSelectable = True
        Me.sfdTextBox.WaterMarkColor = System.Drawing.Color.FromArgb(CType(CType(109, Byte), Integer), CType(CType(109, Byte), Integer), CType(CType(109, Byte), Integer))
        Me.sfdTextBox.WaterMarkFont = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel)
        '
        'MetroLabel3
        '
        Me.MetroLabel3.AutoSize = True
        Me.MetroLabel3.Location = New System.Drawing.Point(15, 83)
        Me.MetroLabel3.Name = "MetroLabel3"
        Me.MetroLabel3.Size = New System.Drawing.Size(123, 19)
        Me.MetroLabel3.TabIndex = 4
        Me.MetroLabel3.Text = "Sandbox Folder Dir"
        '
        'wfdTextBox
        '
        '
        '
        '
        Me.wfdTextBox.CustomButton.Image = Nothing
        Me.wfdTextBox.CustomButton.Location = New System.Drawing.Point(119, 2)
        Me.wfdTextBox.CustomButton.Name = ""
        Me.wfdTextBox.CustomButton.Size = New System.Drawing.Size(23, 23)
        Me.wfdTextBox.CustomButton.Style = MetroFramework.MetroColorStyle.Blue
        Me.wfdTextBox.CustomButton.TabIndex = 1
        Me.wfdTextBox.CustomButton.Theme = MetroFramework.MetroThemeStyle.Light
        Me.wfdTextBox.CustomButton.UseSelectable = True
        Me.wfdTextBox.CustomButton.Visible = False
        Me.wfdTextBox.Lines = New String(-1) {}
        Me.wfdTextBox.Location = New System.Drawing.Point(143, 54)
        Me.wfdTextBox.MaxLength = 32767
        Me.wfdTextBox.Name = "wfdTextBox"
        Me.wfdTextBox.PasswordChar = Global.Microsoft.VisualBasic.ChrW(0)
        Me.wfdTextBox.ScrollBars = System.Windows.Forms.ScrollBars.None
        Me.wfdTextBox.SelectedText = ""
        Me.wfdTextBox.SelectionLength = 0
        Me.wfdTextBox.SelectionStart = 0
        Me.wfdTextBox.ShortcutsEnabled = True
        Me.wfdTextBox.Size = New System.Drawing.Size(145, 28)
        Me.wfdTextBox.TabIndex = 3
        Me.wfdTextBox.UseSelectable = True
        Me.wfdTextBox.WaterMarkColor = System.Drawing.Color.FromArgb(CType(CType(109, Byte), Integer), CType(CType(109, Byte), Integer), CType(CType(109, Byte), Integer))
        Me.wfdTextBox.WaterMarkFont = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel)
        '
        'MetroLabel2
        '
        Me.MetroLabel2.AutoSize = True
        Me.MetroLabel2.Location = New System.Drawing.Point(15, 54)
        Me.MetroLabel2.Name = "MetroLabel2"
        Me.MetroLabel2.Size = New System.Drawing.Size(122, 19)
        Me.MetroLabel2.TabIndex = 2
        Me.MetroLabel2.Text = "Working Folder Dir"
        '
        'MyNameTextBox
        '
        '
        '
        '
        Me.MyNameTextBox.CustomButton.Image = Nothing
        Me.MyNameTextBox.CustomButton.Location = New System.Drawing.Point(119, 2)
        Me.MyNameTextBox.CustomButton.Name = ""
        Me.MyNameTextBox.CustomButton.Size = New System.Drawing.Size(23, 23)
        Me.MyNameTextBox.CustomButton.Style = MetroFramework.MetroColorStyle.Blue
        Me.MyNameTextBox.CustomButton.TabIndex = 1
        Me.MyNameTextBox.CustomButton.Theme = MetroFramework.MetroThemeStyle.Light
        Me.MyNameTextBox.CustomButton.UseSelectable = True
        Me.MyNameTextBox.CustomButton.Visible = False
        Me.MyNameTextBox.Lines = New String(-1) {}
        Me.MyNameTextBox.Location = New System.Drawing.Point(143, 25)
        Me.MyNameTextBox.MaxLength = 32767
        Me.MyNameTextBox.Name = "MyNameTextBox"
        Me.MyNameTextBox.PasswordChar = Global.Microsoft.VisualBasic.ChrW(0)
        Me.MyNameTextBox.ScrollBars = System.Windows.Forms.ScrollBars.None
        Me.MyNameTextBox.SelectedText = ""
        Me.MyNameTextBox.SelectionLength = 0
        Me.MyNameTextBox.SelectionStart = 0
        Me.MyNameTextBox.ShortcutsEnabled = True
        Me.MyNameTextBox.Size = New System.Drawing.Size(145, 28)
        Me.MyNameTextBox.TabIndex = 1
        Me.MyNameTextBox.UseSelectable = True
        Me.MyNameTextBox.WaterMarkColor = System.Drawing.Color.FromArgb(CType(CType(109, Byte), Integer), CType(CType(109, Byte), Integer), CType(CType(109, Byte), Integer))
        Me.MyNameTextBox.WaterMarkFont = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel)
        '
        'MetroLabel1
        '
        Me.MetroLabel1.AutoSize = True
        Me.MetroLabel1.Location = New System.Drawing.Point(15, 25)
        Me.MetroLabel1.Name = "MetroLabel1"
        Me.MetroLabel1.Size = New System.Drawing.Size(67, 19)
        Me.MetroLabel1.TabIndex = 0
        Me.MetroLabel1.Text = "My Name"
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(158, 63)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(588, 421)
        Me.TabControl1.Style = MetroFramework.MetroColorStyle.Teal
        Me.TabControl1.TabIndex = 1
        Me.TabControl1.UseSelectable = True
        '
        'ConfigF
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(769, 493)
        Me.Controls.Add(Me.cancelbtn)
        Me.Controls.Add(Me.savebtn)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.TreeView1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.Name = "ConfigF"
        Me.Style = MetroFramework.MetroColorStyle.Teal
        Me.Text = "Configuration"
        CType(Me.MetroStyleManager1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        CType(Me.Grid1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.TabControl1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents TreeView1 As TreeView
    Friend WithEvents savebtn As MetroFramework.Controls.MetroButton
    Friend WithEvents cancelbtn As MetroFramework.Controls.MetroButton
    Friend WithEvents MetroStyleManager1 As MetroFramework.Components.MetroStyleManager
    Friend WithEvents TabPage2 As MetroFramework.Controls.MetroTabPage
    Friend WithEvents FindProBtn As MetroFramework.Controls.MetroButton
    Friend WithEvents RemoveProBtn As MetroFramework.Controls.MetroButton
    Friend WithEvents AddProBtn As MetroFramework.Controls.MetroButton
    Friend WithEvents MetroLabel18 As MetroFramework.Controls.MetroLabel
    Friend WithEvents ProjectCoorLb As MetroFramework.Controls.MetroLabel
    Friend WithEvents ProjectLb As MetroFramework.Controls.MetroLabel
    Friend WithEvents ProResultPathTextBox As MetroFramework.Controls.MetroTextBox
    Friend WithEvents ProCoorTextBox As MetroFramework.Controls.MetroTextBox
    Friend WithEvents ProModTextBox As MetroFramework.Controls.MetroTextBox
    Friend WithEvents Grid1 As MetroFramework.Controls.MetroGrid
    Friend WithEvents TabPage1 As MetroFramework.Controls.MetroTabPage
    Friend WithEvents Stylecmb As MetroFramework.Controls.MetroComboBox
    Friend WithEvents MetroLabel15 As MetroFramework.Controls.MetroLabel
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents CQPasswordTextBox As MetroFramework.Controls.MetroTextBox
    Friend WithEvents MetroLabel13 As MetroFramework.Controls.MetroLabel
    Friend WithEvents CQUserTextBox As MetroFramework.Controls.MetroTextBox
    Friend WithEvents MetroLabel14 As MetroFramework.Controls.MetroLabel
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents JsonPathTextBox As MetroFramework.Controls.MetroTextBox
    Friend WithEvents MetroLabel17 As MetroFramework.Controls.MetroLabel
    Friend WithEvents FilesTemplateDirTextBox As MetroFramework.Controls.MetroTextBox
    Friend WithEvents MetroLabel12 As MetroFramework.Controls.MetroLabel
    Friend WithEvents MetroLabel8 As MetroFramework.Controls.MetroLabel
    Friend WithEvents myPMTextBox As MetroFramework.Controls.MetroTextBox
    Friend WithEvents teamleadTextBox As MetroFramework.Controls.MetroTextBox
    Friend WithEvents MetroLabel9 As MetroFramework.Controls.MetroLabel
    Friend WithEvents explorerTextBox As MetroFramework.Controls.MetroTextBox
    Friend WithEvents MetroLabel10 As MetroFramework.Controls.MetroLabel
    Friend WithEvents tfdTextBox As MetroFramework.Controls.MetroTextBox
    Friend WithEvents MetroLabel5 As MetroFramework.Controls.MetroLabel
    Friend WithEvents rfdTextBox As MetroFramework.Controls.MetroTextBox
    Friend WithEvents MetroLabel4 As MetroFramework.Controls.MetroLabel
    Friend WithEvents sfdTextBox As MetroFramework.Controls.MetroTextBox
    Friend WithEvents MetroLabel3 As MetroFramework.Controls.MetroLabel
    Friend WithEvents wfdTextBox As MetroFramework.Controls.MetroTextBox
    Friend WithEvents MetroLabel2 As MetroFramework.Controls.MetroLabel
    Friend WithEvents MyNameTextBox As MetroFramework.Controls.MetroTextBox
    Friend WithEvents MetroLabel1 As MetroFramework.Controls.MetroLabel
    Friend WithEvents TabControl1 As MetroFramework.Controls.MetroTabControl
End Class
