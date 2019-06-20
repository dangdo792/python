Namespace Button
    ''' <summary>
    ''' Goto result folder button, use default behaviour of its base class
    ''' </summary>
    Class GotoResult
        Inherits GotoFolderOrILM

        Private Enum Position
            EXPLORER
            RESULT_LINK
        End Enum

        Public Sub New(explorerpath_in As Information.CheckNull, resultlink_in As Information.CheckNull)
            MyBase.New(explorerpath_in, resultlink_in)
        End Sub

    End Class
End Namespace
