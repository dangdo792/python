Namespace Button
    Class GotoReview
        Inherits GotoBase

        Private Enum Position
            EXPLORER
            REVIEW_FOLDER
            TASK_ID
            MODULE_NAME
        End Enum


        Public Sub New(explorerpath_in As Information.CheckNull,
                       rfd As Information.CheckNull,
                       taskid As Information.TaskID,
                       modulename As Information.CheckNull)

            MyBase.New({explorerpath_in, rfd, taskid, modulename})
        End Sub

        ''' <summary>
        ''' Get full path. Example D:\Working\Review\SDR_ESR2_613431_RTRT_Mod_Test_FMHADAPT_X
        ''' </summary>
        ''' <returns>string of full path</returns>
        Public Overrides Function GetFullPath() As String
            Return listinfo(Position.REVIEW_FOLDER).GetValue & "\" & listinfo(Position.TASK_ID).GetValue & "_" & UCase(listinfo(Position.MODULE_NAME).GetValue)
        End Function
    End Class

End Namespace
