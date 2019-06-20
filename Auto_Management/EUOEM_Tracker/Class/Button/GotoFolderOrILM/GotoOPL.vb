Namespace Button
    ''' <summary>
    ''' Goto OPL folder button, need to adapt the folder path for FOLDER (not ILM) situation
    ''' </summary>
    Class GotoOPL
        Inherits GotoFolderOrILM

        Private Enum Position
            EXPLORER
            OPL_LINK
        End Enum

        Public Sub New(explorerpath_in As Information.CheckNull, opllink_in As Information.CheckNull)
            MyBase.New(explorerpath_in, opllink_in)
        End Sub

        Public Overrides Function GetCorrectPath_Folder() As String
            Return reg.Search_f(listinfo(Position.OPL_LINK).GetValue, "^.*\\[O][P][L]")
        End Function
    End Class

End Namespace
