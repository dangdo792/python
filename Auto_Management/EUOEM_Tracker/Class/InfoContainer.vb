Namespace Container
    Public Class TaskInfo
        Public ID As Information.NoCheck
        Public Project As Information.CheckNull
        Public Model As Information.CheckNull
        Public Release As Information.CheckNull
        Public Task_ID As Information.TaskID
        Public Eloc As Information.CheckNull
        Public ModuleName As Information.CheckNull
        Public Revision As Information.CheckNull
        Public Sandbox As Information.CheckNull
        Public Feature_Branch As Information.NoCheck
        Public M_Path As Information.ModulePath
        Public M_Owner As Information.CheckNull
        Public Old_Task As Information.NoCheck
        Public Reviewer As Information.CheckNull
        Public Defect_ID As Information.NoCheck
        Public Sub_Rev As Information.NoCheck
        Public ILM_Link As Information.CheckNull
        Public OPL_Link As Information.NoCheck
        Public RS As Information.CheckNull
        Public RS_BL As Information.DoorBaseline
        Public TS As Information.CheckNull
        Public TS_BL As Information.DoorBaseline
        Public SD As Information.NoCheck
        Public SD_BL As Information.NoCheck
        Public Statement As Information.CheckNull
        Public Decisions As Information.NoCheck
        Public Cover_Path As Information.NoCheck
        Public Result_Path As Information.CheckNull
        Public Package_ID As Information.NoCheck


        Public Sub New(ID_in As String,
                         Project_in As String,
                         Model_in As String,
                         Release_in As String,
                         Task_ID_in As String,
                         Eloc_in As String,
                         ModuleName_in As String,
                         Revision_in As String,
                         Sandbox_in As String,
                         Feature_Branch_in As String,
                         M_Path_in As String,
                         M_Owner_in As String,
                         Old_Task_in As String,
                         Reviewer_in As String,
                         Defect_ID_in As String,
                         Sub_Rev_in As String,
                         ILM_Link_in As String,
                         OPL_Link_in As String,
                         RS_in As String,
                         RS_BL_in As String,
                         TS_in As String,
                         TS_BL_in As String,
                         SD_in As String,
                         SD_BL_in As String,
                         Statement_in As String,
                         Decisions_in As String,
                         Cover_Path_in As String,
                         Result_Path_in As String,
                         Package_ID_in As String)

            ID = New Information.NoCheck(ID_in)
            Project = New Information.CheckNull("Project", Project_in)
            Model = New Information.CheckNull("Model", Model_in)
            Release = New Information.CheckNull("Release", Release_in)
            Task_ID = New Information.TaskID(Task_ID_in)
            Eloc = New Information.CheckNull("Eloc", Eloc_in)
            ModuleName = New Information.CheckNull("Module Name", ModuleName_in)
            Revision = New Information.CheckNull("Revision", Revision_in)
            Sandbox = New Information.CheckNull("Sandbox", Sandbox_in)
            Feature_Branch = New Information.NoCheck(Feature_Branch_in)
            M_Path = New Information.ModulePath(M_Path_in)
            M_Owner = New Information.CheckNull("Module Owner", M_Owner_in)
            Old_Task = New Information.NoCheck(Old_Task_in)
            Reviewer = New Information.CheckNull("Reviewer", Reviewer_in)
            Defect_ID = New Information.NoCheck(Defect_ID_in)
            Sub_Rev = New Information.NoCheck(Sub_Rev_in)
            ILM_Link = New Information.CheckNull("ILM Link", ILM_Link_in)
            OPL_Link = New Information.NoCheck(OPL_Link_in)
            RS = New Information.CheckNull("RS", RS_in)
            RS_BL = New Information.DoorBaseline("RS Baseline", RS_BL_in)
            TS = New Information.CheckNull("TS", TS_in)
            TS_BL = New Information.DoorBaseline("TS Baseline", TS_BL_in)
            SD = New Information.NoCheck(SD_in)
            SD_BL = New Information.NoCheck(SD_BL_in)
            Statement = New Information.CheckNull("Statement", Statement_in)
            Decisions = New Information.NoCheck(Decisions_in)
            Cover_Path = New Information.NoCheck(Cover_Path_in)
            Result_Path = New Information.CheckNull("Result Path", Result_Path_in)
            Package_ID = New Information.NoCheck(Package_ID_in)
        End Sub

    End Class

    Class UserInfo
        Public MyName As Information.CheckNull
        Public wfd As Information.CheckNull
        Public sfd As Information.CheckNull
        Public rfd As Information.CheckNull
        Public tfd As Information.CheckNull
        Public ExplorerPath As Information.CheckNull
        Public TeamLead As Information.CheckNull
        Public PM As Information.CheckNull
        Public CMacro As Information.CheckNull
        Public CppMacro As Information.CheckNull
        Public SysBPlusMacro As Information.CheckNull
        Public CQUser As Information.CheckNull
        Public CQPassword As Information.CheckNull
        Public MKSChecker As Information.CheckNull
        Public FilesTempplateDir As Information.CheckNull

        Public Sub New(MyName_in As String,
                          wfd_in As String,
                          sfd_in As String,
                          rfd_in As String,
                          tfd_in As String,
                          ExplorerPath_in As String,
                          TeamLead_in As String,
                          PM_in As String,
                          CMacro_in As String,
                          CppMacro_in As String,
                          SysBPlusMacro_in As String,
                          CQUser_in As String,
                          CQPassword_in As String,
                          MKSChecker_in As String,
                       FilesTempplateDir_in As String)

            MyName = New Information.CheckNull("My Name", MyName_in)
            wfd = New Information.CheckNull("Working Folder Dir", wfd_in)
            sfd = New Information.CheckNull("Sandbox Folder Dir", sfd_in)
            rfd = New Information.CheckNull("Review Folder Dir", rfd_in)
            tfd = New Information.CheckNull("Tool Folder Dir", tfd_in)
            ExplorerPath = New Information.CheckNull("Explorer Path", ExplorerPath_in)
            TeamLead = New Information.CheckNull("Team Leader", TeamLead_in)
            PM = New Information.CheckNull("PM", PM_in)
            CMacro = New Information.CheckNull("C Macro", CMacro_in)
            CppMacro = New Information.CheckNull("Cpp Macro", CppMacro_in)
            SysBPlusMacro = New Information.CheckNull("SysBPlus Macro", SysBPlusMacro_in)
            CQUser = New Information.CheckNull("CQ User", CQUser_in)
            CQPassword = New Information.CheckNull("CQ Password", CQPassword_in)
            MKSChecker = New Information.CheckNull("MKS Checker", MKSChecker_in)
            FilesTempplateDir = New Information.CheckNull("Files template Dir", FilesTempplateDir_in)
        End Sub
    End Class

    Class GetInfo
        Public errorFlag As Boolean
        Public errorMsg As String

        Public Function GetTaskInfo(ByRef selrow() As DataRow) As TaskInfo()
            Dim index As Integer = 0
            Dim temp_gettaskinfo() As TaskInfo = Nothing
            If Not selrow Is Nothing Then
                If selrow.Length <> 0 Then
                    ReDim Preserve temp_gettaskinfo(selrow.Count - 1)
                    For Each r As DataRow In selrow
                        Dim temp As New TaskInfo(r.Item("ID").ToString.Trim,
                                                    r.Item("Project").ToString.Trim,
                                                    r.Item("Model").ToString.Trim,
                                                    r.Item("Release").ToString.Trim,
                                                    r.Item("Task_ID").ToString.Trim,
                                                    r.Item("Eloc").ToString.Trim,
                                                    r.Item("Module").ToString.Trim,
                                                    r.Item("Revision").ToString.Trim,
                                                    r.Item("Sandbox").ToString.Trim,
                                                    r.Item("Feature_Branch").ToString.Trim,
                                                    r.Item("M_Path").ToString.Trim,
                                                    r.Item("M_Owner").ToString.Trim,
                                                    r.Item("Old_Task").ToString.Trim,
                                                    r.Item("Reviewer").ToString.Trim,
                                                    r.Item("Defect_ID").ToString.Trim,
                                                    r.Item("Sub_Rev").ToString.Trim,
                                                    r.Item("ILM_Link").ToString.Trim,
                                                    r.Item("OPL_Link").ToString.Trim,
                                                    r.Item("RS").ToString.Trim,
                                                    r.Item("RS_BL").ToString.Trim,
                                                    r.Item("TS").ToString.Trim,
                                                    r.Item("TS_BL").ToString.Trim,
                                                    r.Item("SD").ToString.Trim,
                                                    r.Item("SD_BL").ToString.Trim,
                                                    r.Item("Statement").ToString.Trim,
                                                    r.Item("Decisions").ToString.Trim,
                                                    r.Item("Cover_Path").ToString.Trim,
                                                    r.Item("Result_Path").ToString.Trim,
                                                    r.Item("Package_ID").ToString.Trim)
                        temp_gettaskinfo(index) = temp
                        index = index + 1
                    Next
                Else
                End If
                ReDim Preserve GetTaskInfo(temp_gettaskinfo.Count - 1)
            Else
            End If


            GetTaskInfo = temp_gettaskinfo
        End Function

        Public Function GetUserInfo() As UserInfo
            Dim temp As New UserInfo(MainF.ds.Tables("User_Config").Rows(0).Item("MyName").ToString.Trim,
                                    MainF.ds.Tables("User_Config").Rows(0).Item("wfd").ToString.Trim,
                                    MainF.ds.Tables("User_Config").Rows(0).Item("sfd").ToString.Trim,
                                    MainF.ds.Tables("User_Config").Rows(0).Item("rfd").ToString.Trim,
                                    MainF.ds.Tables("User_Config").Rows(0).Item("tfd").ToString.Trim,
                                    MainF.ds.Tables("User_Config").Rows(0).Item("explorer").ToString.Trim,
                                    MainF.ds.Tables("User_Config").Rows(0).Item("teamlead").ToString.Trim,
                                    MainF.ds.Tables("User_Config").Rows(0).Item("myPM").ToString.Trim,
                                    MainF.ds.Tables("User_Config").Rows(0).Item("CMacro").ToString.Trim,
                                    MainF.ds.Tables("User_Config").Rows(0).Item("CppMacro").ToString.Trim,
                                    MainF.ds.Tables("User_Config").Rows(0).Item("SysBPlusMacro").ToString.Trim,
                                    MainF.ds.Tables("User_Config").Rows(0).Item("User").ToString.Trim,
                                    MainF.ds.Tables("User_Config").Rows(0).Item("Password").ToString.Trim,
                                    MainF.ds.Tables("User_Config").Rows(0).Item("mks_checker").ToString.Trim(),
                                    MainF.ds.Tables("User_Config").Rows(0).Item("FilesTemplateDir").ToString.Trim())
            GetUserInfo = temp
        End Function
    End Class


End Namespace


