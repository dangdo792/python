Imports System.Reflection

Namespace My
    ' The following events are available for MyApplication:
    ' Startup: Raised when the application starts, before the startup form is created.
    ' Shutdown: Raised after all application forms are closed.  This event is not raised if the application terminates abnormally.
    ' UnhandledException: Raised if the application encounters an unhandled exception.
    ' StartupNextInstance: Raised when launching a single-instance application and the application is already active. 
    ' NetworkAvailabilityChanged: Raised when the network connection is connected or disconnected.
    Partial Friend Class MyApplication
        Private WithEvents Domain As AppDomain = AppDomain.CurrentDomain

        Private Function Domain_AssemblyResolve(sender As Object, e As ResolveEventArgs) As Assembly Handles Domain.AssemblyResolve
            If e.Name.Contains("MetroFramework") Then
                Return System.Reflection.Assembly.Load(My.Resources.MetroFramework)
            ElseIf e.Name.Contains("MetroFramework.Design") Then
                Return System.Reflection.Assembly.Load(My.Resources.MetroFramework_Design)
            ElseIf e.Name.Contains("MetroFramework.Fonts") Then
                Return System.Reflection.Assembly.Load(My.Resources.MetroFramework_Fonts)
            ElseIf e.Name.Contains("Microsoft.Office.Interop.Excel") Then
                Return System.Reflection.Assembly.Load(My.Resources.Microsoft_Office_Interop_Excel)
            ElseIf e.Name.Contains("Microsoft.Office.Interop.Word") Then
                Return System.Reflection.Assembly.Load(My.Resources.Microsoft_Office_Interop_Word)
            ElseIf e.Name.Contains("Microsoft.Office.Interop.Outlook") Then
                Return System.Reflection.Assembly.Load(My.Resources.Microsoft_Office_Interop_Outlook)
            ElseIf e.Name.Contains("Newtonsoft.Json") Then
                Return System.Reflection.Assembly.Load(My.Resources.Newtonsoft_Json)
            ElseIf e.Name.Contains("interop.adox") Then
                Return System.Reflection.Assembly.Load(My.Resources.interop_adox)
            ElseIf e.Name.Contains("System.IO.Compression.FileSystem") Then
                Return System.Reflection.Assembly.Load(My.Resources.System_IO_Compression_FileSystem)
            ElseIf e.Name.Contains("Interop.MSXML2") Then
                Return System.Reflection.Assembly.Load(My.Resources.Interop_MSXML2)
            ElseIf e.Name.Contains("System.IO.Compression") Then
                Return System.Reflection.Assembly.Load(My.Resources.System_IO_Compression)
            Else
                Return Nothing
            End If
        End Function
    End Class
End Namespace
