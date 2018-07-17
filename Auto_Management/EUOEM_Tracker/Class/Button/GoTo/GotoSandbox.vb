Namespace Button
    Class GotoSandBox
        Inherits GotoBase

        Private Enum Position
            EXPLORER
            SANDBOX_FOLDER
            SANDBOX
            MODULE_PATH
        End Enum


        Public Sub New(explorerpath_in As Information.CheckNull,
                       sfd As Information.CheckNull,
                       sandbox As Information.CheckNull,
                       modulepath As Information.ModulePath)

            MyBase.New({explorerpath_in, sfd, sandbox, modulepath})
        End Sub

        Public Overrides Function GetFullPath() As String
            Return listinfo(Position.SANDBOX_FOLDER).GetValue & "\" & listinfo(Position.SANDBOX).GetValue _
                & Left(listinfo(Position.MODULE_PATH).GetValue, InStrRev(listinfo(Position.MODULE_PATH).GetValue, "\") - 1)
        End Function
    End Class

End Namespace
