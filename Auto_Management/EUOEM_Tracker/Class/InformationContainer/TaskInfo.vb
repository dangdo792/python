Imports EUOEM_Tracker.Information

Namespace TContainer
    Public Class TaskInfo
        Private selrow As DataRow

        Public Sub New(ByRef selrow_in As DataRow)
            selrow = selrow_in
        End Sub

        Public Property ID As String
            Get
                Return selrow.Item("ID").ToString
            End Get
            Set(value As String)
                selrow.Item("ID") = value
            End Set
        End Property

        Public Property Project As String
            Get
                Return selrow.Item("Project").ToString
            End Get
            Set(value As String)
                selrow.Item("Project") = value
            End Set
        End Property

        Public Property Model As String
            Get
                Return selrow.Item("Model").ToString
            End Get
            Set(value As String)
                selrow.Item("Model") = value
            End Set
        End Property

        Public Property Release As String
            Get
                Return selrow.Item("Release").ToString
            End Get
            Set(value As String)
                selrow.Item("Release") = value
            End Set
        End Property

        Public Property Task_ID As String
            Get
                Return selrow.Item("Task_ID").ToString
            End Get
            Set(value As String)
                selrow.Item("Task_ID") = value
            End Set
        End Property

        Public Property Eloc As String
            Get
                Return selrow.Item("Eloc").ToString
            End Get
            Set(value As String)
                selrow.Item("Eloc") = value
            End Set
        End Property

        Public Property ModuleName As String
            Get
                Return selrow.Item("Module").ToString
            End Get
            Set(value As String)
                selrow.Item("Module") = value
            End Set
        End Property

        Public Property Revision As String
            Get
                Return selrow.Item("Revision").ToString
            End Get
            Set(value As String)
                selrow.Item("Revision") = value
            End Set
        End Property

        Public Property Sandbox As String
            Get
                Return selrow.Item("Sandbox").ToString
            End Get
            Set(value As String)
                selrow.Item("Sandbox") = value
            End Set
        End Property

        Public Property Feature_Branch As String
            Get
                Return selrow.Item("Feature_Branch").ToString
            End Get
            Set(value As String)
                selrow.Item("Feature_Branch") = value
            End Set
        End Property

        Public Property M_Path As String
            Get
                Return selrow.Item("M_Path").ToString
            End Get
            Set(value As String)
                selrow.Item("M_Path") = value
            End Set
        End Property

        Public Property M_Owner As String
            Get
                Return selrow.Item("M_Owner").ToString
            End Get
            Set(value As String)
                selrow.Item("M_Owner") = value
            End Set
        End Property

        Public Property Old_Task As String
            Get
                Return selrow.Item("Old_Task").ToString
            End Get
            Set(value As String)
                selrow.Item("Old_Task") = value
            End Set
        End Property

        Public Property Reviewer As String
            Get
                Return selrow.Item("Reviewer").ToString
            End Get
            Set(value As String)
                selrow.Item("Reviewer") = value
            End Set
        End Property

        Public Property Defect_ID As String
            Get
                Return selrow.Item("Defect_ID").ToString
            End Get
            Set(value As String)
                selrow.Item("Defect_ID") = value
            End Set
        End Property

        Public Property Sub_Rev As String
            Get
                Return selrow.Item("Sub_Rev").ToString
            End Get
            Set(value As String)
                selrow.Item("Sub_Rev") = value
            End Set
        End Property

        Public Property ILM_Link As String
            Get
                Return selrow.Item("ILM_Link").ToString
            End Get
            Set(value As String)
                selrow.Item("ILM_Link") = value
            End Set
        End Property

        Public Property OPL_Link As String
            Get
                Return selrow.Item("OPL_Link").ToString
            End Get
            Set(value As String)
                selrow.Item("OPL_Link") = value
            End Set
        End Property

        Public Property RS As String
            Get
                Return selrow.Item("RS").ToString
            End Get
            Set(value As String)
                selrow.Item("RS") = value
            End Set
        End Property

        Public Property RS_BL As String
            Get
                Return selrow.Item("RS_BL").ToString
            End Get
            Set(value As String)
                selrow.Item("RS_BL") = value
            End Set
        End Property

        Public Property TS As String
            Get
                Return selrow.Item("TS").ToString
            End Get
            Set(value As String)
                selrow.Item("TS") = value
            End Set
        End Property

        Public Property TS_BL As String
            Get
                Return selrow.Item("TS_BL").ToString
            End Get
            Set(value As String)
                selrow.Item("TS_BL") = value
            End Set
        End Property

        Public Property SD As String
            Get
                Return selrow.Item("SD").ToString
            End Get
            Set(value As String)
                selrow.Item("SD") = value
            End Set
        End Property

        Public Property SD_BL As String
            Get
                Return selrow.Item("SD_BL").ToString
            End Get
            Set(value As String)
                selrow.Item("SD_BL") = value
            End Set
        End Property

        Public Property Statement As String
            Get
                Return selrow.Item("Statement").ToString
            End Get
            Set(value As String)
                selrow.Item("Statement") = value
            End Set
        End Property

        Public Property Decisions As String
            Get
                Return selrow.Item("Decisions").ToString
            End Get
            Set(value As String)
                selrow.Item("Decisions") = value
            End Set
        End Property

        Public Property Cover_Path As String
            Get
                Return selrow.Item("Cover_Path").ToString
            End Get
            Set(value As String)
                selrow.Item("Cover_Path") = value
            End Set
        End Property

        Public Property Result_Path As String
            Get
                Return selrow.Item("Result_Path").ToString
            End Get
            Set(value As String)
                selrow.Item("Result_Path") = value
            End Set
        End Property

        Public Property Package_ID As String
            Get
                Return selrow.Item("Package_ID").ToString
            End Get
            Set(value As String)
                selrow.Item("Package_ID") = value
            End Set
        End Property

        Public Property Status As String
            Get
                Return selrow.Item("Status").ToString
            End Get
            Set(value As String)
                selrow.Item("Status") = value
            End Set
        End Property

    End Class

    Class UserInfo
        Private row As DataRow

        Public Sub New(row_in As DataRow)
            row = row_in
        End Sub

        Public Property MyName As String
            Get
                Return row.Item("MyName").ToString
            End Get
            Set(value As String)
                row.Item("MyName") = value
            End Set
        End Property

        Public Property Wfd As String
            Get
                Return row.Item("wfd").ToString
            End Get
            Set(value As String)
                row.Item("wfd") = value
            End Set
        End Property

        Public Property Sfd As String
            Get
                Return row.Item("sfd").ToString
            End Get
            Set(value As String)
                row.Item("sfd") = value
            End Set
        End Property

        Public Property Rfd As String
            Get
                Return row.Item("rfd").ToString
            End Get
            Set(value As String)
                row.Item("rfd") = value
            End Set
        End Property

        Public Property Tfd As String
            Get
                Return row.Item("tfd").ToString
            End Get
            Set(value As String)
                row.Item("tfd") = value
            End Set
        End Property

        Public Property ExplorerPath As String
            Get
                Return row.Item("explorer").ToString
            End Get
            Set(value As String)
                row.Item("explorer") = value
            End Set
        End Property

        Public Property TeamLead As String
            Get
                Return row.Item("teamlead").ToString
            End Get
            Set(value As String)
                row.Item("teamlead") = value
            End Set
        End Property

        Public Property PM As String
            Get
                Return row.Item("myPM").ToString
            End Get
            Set(value As String)
                row.Item("myPM") = value
            End Set
        End Property

        Public Property CMacro As String
            Get
                Return row.Item("CMacro").ToString
            End Get
            Set(value As String)
                row.Item("CMacro") = value
            End Set
        End Property

        Public Property CppMacro As String
            Get
                Return row.Item("CppMacro").ToString
            End Get
            Set(value As String)
                row.Item("CppMacro") = value
            End Set
        End Property

        Public Property SysBPlusMacro As String
            Get
                Return row.Item("SysBPlusMacro").ToString
            End Get
            Set(value As String)
                row.Item("SysBPlusMacro") = value
            End Set
        End Property

        Public Property CQUser As String
            Get
                Return row.Item("User").ToString
            End Get
            Set(value As String)
                row.Item("User") = value
            End Set
        End Property

        Public Property CQPassword As String
            Get
                Return row.Item("Password").ToString
            End Get
            Set(value As String)
                row.Item("Password") = value
            End Set
        End Property

        Public Property MKSChecker As String
            Get
                Return row.Item("mks_checker").ToString
            End Get
            Set(value As String)
                row.Item("mks_checker") = value
            End Set
        End Property

        Public Property FilesTemplateDir As String
            Get
                Return row.Item("FilesTemplateDir").ToString
            End Get
            Set(value As String)
                row.Item("FilesTemplateDir") = value
            End Set
        End Property

        Public Property JsonPath As String
            Get
                Return row.Item("JsonPath").ToString
            End Get
            Set(value As String)
                row.Item("JsonPath") = value
            End Set
        End Property

    End Class
End Namespace