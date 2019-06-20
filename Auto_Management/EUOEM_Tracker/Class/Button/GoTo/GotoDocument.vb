Namespace Button
    Class GotoDocument
        Inherits GotoBase

        Private Enum Position
            EXPLORER
            WORKING_FOLDER
            PROJECT
            TASK_ID
            MODULE_NAME
        End Enum

        Public Sub New(explorerpath_in As Information.CheckNull,
                       wfd As Information.CheckNull,
                       project As Information.CheckNull,
                       taskid As Information.TaskID,
                       modulename As Information.CheckNull)

            MyBase.New({explorerpath_in, wfd, project, taskid, modulename})
        End Sub

        ''' <summary>
        ''' Get full path. Example D:\Working\Task\GWM_CHB131_MRRevo14\SDR_ESR2_613431_RTRT_Mod_Test_FMHADAPT_X
        ''' </summary>
        ''' <returns>string of full path</returns>
        Public Overrides Function GetFullPath() As String
            Return listinfo(Position.WORKING_FOLDER).GetValue & "\" & UCase(listinfo(Position.PROJECT).GetValue) & "\" & listinfo(Position.TASK_ID).GetValue & "_" & listinfo(Position.MODULE_NAME).GetValue
        End Function
    End Class

End Namespace
