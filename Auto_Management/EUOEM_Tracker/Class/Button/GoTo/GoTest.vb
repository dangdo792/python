Namespace Button
    Class GotoTest
        Inherits GotoBase

        Private Enum Position
            EXPLORER
            SANDBOX_FOLDER
            SANDBOX
            MODULE_PATH
            MODULE_NAME
        End Enum


        Public Sub New(explorerpath_in As Information.CheckNull,
                       sfd As Information.CheckNull,
                       sandbox As Information.CheckNull,
                       modulepath As Information.ModulePath,
                       modulename As Information.CheckNull)

            MyBase.New({explorerpath_in, sfd, sandbox, modulepath, modulename})
        End Sub

        Public Overrides Function GetFullPath() As String
            Return listinfo(Position.SANDBOX_FOLDER).GetValue & "\" & listinfo(Position.SANDBOX).GetValue & "\_test\" & listinfo(Position.MODULE_NAME).GetValue
        End Function
    End Class

End Namespace
