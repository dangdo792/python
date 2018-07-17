Namespace Button
    Class AutoGenReview
        Inherits AutoGenBase

        ''' <summary>
        ''' The position of first 3 field is a must, do not adapt it
        ''' </summary>
        Private Enum Position
            TOOL_FOLDER
            MODULE_PATH
            TASK_ID
            EXPLORER
            REVIEW_FOLDER
            MODULE_NAME
        End Enum

        Public Sub New(explorer_in As Information.CheckNull,
                       rfd_in As Information.CheckNull,
                       tfd_in As Information.CheckNull,
                       taskid_in As Information.TaskID,
                       modulename_in As Information.CheckNull,
                       modulepath_in As Information.ModulePath)
            MyBase.New({tfd_in,
                        modulepath_in,
                        taskid_in,
                        explorer_in,
                        rfd_in,
                        modulename_in})
        End Sub

        Public Overrides Function GetFullPath() As String
            Dim getrevpath As New Button.GotoReview(listinfo(Position.EXPLORER), listinfo(Position.REVIEW_FOLDER), listinfo(Position.TASK_ID), listinfo(Position.MODULE_NAME))
            Return getrevpath.GetFullPath
        End Function

    End Class
End Namespace
