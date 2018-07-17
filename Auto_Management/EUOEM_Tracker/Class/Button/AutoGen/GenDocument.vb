Namespace Button
    Class AutoGenDoc
        Inherits AutoGenBase

        ''' <summary>
        ''' The position of first 3 field is a must, do not adapt it
        ''' </summary>
        Private Enum Position
            TOOL_FOLDER
            MODULE_PATH
            TASK_ID
            EXPLORER
            WORKING_FOLDER
            PROJECT
            MODEL
            MODULE_NAME
        End Enum

        Public Sub New(explorer_in As Information.CheckNull,
                       wfd_in As Information.CheckNull,
                       tfd_in As Information.CheckNull,
                       project_in As Information.CheckNull,
                       model_in As Information.CheckNull,
                       taskid_in As Information.TaskID,
                       modulename_in As Information.CheckNull,
                       modulepath_in As Information.ModulePath)
            MyBase.New({tfd_in,
                        modulepath_in,
                        taskid_in,
                        explorer_in,
                        wfd_in,
                        project_in,
                        model_in,
                        modulename_in})
        End Sub

        Public Overrides Function GetFullPath() As String
            Dim getdocpath As New Button.GotoDocument(listinfo(Position.EXPLORER), listinfo(Position.WORKING_FOLDER), listinfo(Position.PROJECT), listinfo(Position.TASK_ID), listinfo(Position.MODULE_NAME))
            Return getdocpath.GetFullPath
        End Function
    End Class
End Namespace

