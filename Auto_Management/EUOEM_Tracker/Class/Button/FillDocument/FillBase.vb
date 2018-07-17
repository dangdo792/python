Namespace Button
    Public MustInherit Class FillBase
        Inherits ButtonBase

        Private Enum Position
            EXPLORER
            WORKING_FOLDER
            PROJECT
            TASK_ID
            MODULE_NAME
        End Enum

        Public Sub New(listinfo_in() As Information.InfoBase)
            MyBase.New(listinfo_in)
        End Sub

        ''' <summary>
        ''' Check if the file is exist or not
        ''' </summary>
        ''' <returns>True if file exist, False otherwise</returns>
        Overrides Function AdditionCondition() As Boolean
            Dim PathObject As Information.CheckPathExist = New Information.CheckPathExist(Me.GetFullPath())
            Dim validFlag As Boolean = PathObject.IsValid()
            Me.additional_errorMsg = PathObject.GetErrorMsg()

            Return validFlag
        End Function

        ''' <summary>
        ''' The full path of file
        ''' </summary>
        ''' <returns>String full path of file</returns>
        Public MustOverride Function GetFullPath() As String

        ''' <summary>
        ''' Get document folder path
        ''' </summary>
        ''' <returns>String document folder path</returns>
        Protected Function GetDocumentFolderPath()
            Dim getdocpath As New Button.GotoDocument(listinfo(Position.EXPLORER), listinfo(Position.WORKING_FOLDER), listinfo(Position.PROJECT),
                                                      listinfo(Position.TASK_ID), listinfo(Position.MODULE_NAME))
            Return getdocpath.GetFullPath()
        End Function

    End Class
End Namespace
