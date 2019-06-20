Imports Scripting
Imports VBScript_RegExp_55

Public Module ilm
    Private Const INTERNET_FLAG_NO_CACHE_WRITE = &H4000000
    Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal lpszAgent As String, ByVal dwAccessType As Int32, ByVal lpszProxyName As String, ByVal lpszProxyBypass As String, ByVal dwFlags As Int32) As Int32
    Private Declare Function InternetReadBinaryFile Lib "wininet.dll" Alias "InternetReadFile" (ByVal hfile As Int32, ByRef bytearray_firstelement As Byte, ByVal lNumBytesToRead As Int32, ByRef lNumberOfBytesRead As Int32) As Integer
    Private Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" (ByVal hInternetSession As Int32, ByVal sUrl As String, ByVal sHeaders As String, ByVal lHeadersLength As Int32, ByVal lFlags As Int32, ByVal lContext As Int32) As Int32
    Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Int32) As Integer


    Private Const RETRIEVE_INFO = "https://inside-ilm.bosch.com/irj/go/nui/cmd/ls/?id="
    Private Const DOWNLOAD_FILE = "https://inside-ilm.bosch.com/irj/go/km/docs/?action=download&URI=/irj/go/km/docs"
    'Private Const ROOT_DIR = "/room_extensions_ilm_2/cm_stores/documents/workspaces/b1066f26-0e03-3410-2db6-f61ff40ede23/EI-300010_CC_DA_SWEngineering/Content/10_Development/DA_RADAR_MT/Project_Engineering_VSS/"
    Private Const ROOT_DIR = "/room_extensions_ilm_2/cm_stores/documents/workspaces/b1066f26-0e03-3410-2db6-f61ff40ede23/EI-300010_CC_DA_SWEngineering/Content/10_Development/DA_RADAR_MT/Project_Engineer_RADAR_Gen5"

    Public Function GetResult(ByVal url As String) As String
        Dim XMLHTTP = CreateObject("MSXML2.XMLHTTP.6.0")
        Dim ret As String
        XMLHTTP.Open("GET", url, False)
        XMLHTTP.setRequestHeader("Host", "inside-ilm.bosch.com")
        XMLHTTP.send()
        ret = XMLHTTP.ResponseText
        GetResult = ret
    End Function

    Private Function GetListItems(ByVal str As String) As Collection
        Dim resCol As Collection
        resCol = New Collection
        Dim Regex As New RegExp

        'Preprocessing()
        With Regex
            .Pattern = """found""[\s]*:[\s]*\[([^\]]+)\]"
            .Global = True
            .IgnoreCase = True
        End With
        If Regex.Test(str) Then
            str = Regex.Execute(str)(0).Value
        Else
            GoTo endFunc
        End If

        'Process list item
        Regex.Pattern = """path""[^\}]+?\{"
        If Regex.Test(str) Then
            'Continue
        Else
            GoTo endFunc
        End If

        Dim item As String
        Dim Match As Match
        For Each Match In Regex.Execute(str)
            item = Match.Value
            Dim tPath As String
            tPath = Search_f(item, """name.+?:.+?""(.+?)""")
            tPath = "path=" + Replace_f(tPath, """name.+?:.+?""(.+?)""", "$1")
            Dim tCollect As String
            tCollect = Search_f(item, """iscollection"".+?:.+?(\w+)\W")
            tCollect = ";iscollection=" + Replace_f(tCollect, """iscollection"".+?:.+?(\w+)\W", "$1")
            If tPath <> "" Then resCol.Add(tPath + tCollect)
        Next Match

endFunc:
        GetListItems = resCol
    End Function

    Public Function GetListOfSubProject(ByVal Path As String) As Collection
        Dim requestResult As String
        If InStr(Path, "https://") = 0 Then
            Path = RETRIEVE_INFO + Path
        End If
        requestResult = GetResult(Path)
        GetListOfSubProject = GetListItems(requestResult)
    End Function

    Function Search_VSS_Link(ByVal Task_ID As String, ByVal Project As String, ByRef LinkILM As String, ByRef ErrorMsg As String) As Boolean
        Dim ErrorFlag As Boolean = False
        If Task_ID = "" Then ErrorMsg = "'Task_ID' is empty. Please check" : ErrorFlag = True
        Dim taskid_double As Double
        Double.TryParse(Search_f(Task_ID, "\d+"), taskid_double)
        If ErrorFlag = False Then
            Dim listCustomer1 As Collection
            Dim Year_No As Integer
            Year_No = Year(Date.Today)
            Do While (Year_No > 2015) And (taskid_double <> 0)
                listCustomer1 = GetListOfSubProject(ROOT_DIR & "/" & Year_No)

                If listCustomer1.Count > 0 Then
                    Dim i_cus As String
                    For Each i_cus In listCustomer1
                        i_cus = Replace_f(i_cus, "path=(.+);iscollection=(.+)", "$1")
                        If InStr(LCase(i_cus), LCase(Project)) <> 0 Or Project = "" Then
                            Dim listModule As Collection
                            listModule = GetListOfSubProject(ROOT_DIR & "/" & Year_No & "/" & i_cus)
                            If listModule.Count > 0 Then
                                Dim i_mod As String
                                For Each i_mod In listModule
                                    i_mod = Replace_f(i_mod, "path=(.+);iscollection=(.+)", "$1")
                                    If InStr(LCase(i_mod), taskid_double) <> 0 And i_mod <> "" Then
                                        LinkILM = ROOT_DIR & "/" & Year_No & "/" & i_cus & "/" & i_mod
                                        Search_VSS_Link = ErrorFlag
                                        Exit Function
                                    End If
                                Next
                            End If
                        End If
                    Next
                End If
                Year_No = Year_No - 1
            Loop
            ErrorMsg = "This task doesn't exist in ILM server" & vbNewLine &
                        "Please check whether your input is right" & vbNewLine &
                        "'Task ID': " & Task_ID & vbNewLine &
                        "'Project': " & Project : ErrorFlag = True
        End If
        Search_VSS_Link = ErrorFlag
    End Function

    Public Sub DownloadFile(ByVal sUrl As String, ByVal filePath As String, ByVal consoletextbox As MetroFramework.Controls.MetroTextBox, Optional ByVal overWriteFile As Boolean = False)
        Dim hInternet As Long, hSession As Long, lngDataReturned As Long, sBuffer() As Byte, totalRead As Long
        Dim FileName As String
        'SendMsg.Start()
        FileName = Replace_f(filePath, ".+[\/\\](.+)", "$1")
        Const bufSize = 128
        ReDim sBuffer(bufSize)
        hSession = InternetOpen("browser", 0, vbNullString, vbNullString, 0)
        If hSession Then hInternet = InternetOpenUrl(hSession, sUrl, vbNullString, 0, INTERNET_FLAG_NO_CACHE_WRITE, 0)
        Dim oStream As Object
        oStream = CreateObject("ADODB.Stream")
        oStream.Open()
        oStream.Type = 1

        Dim iReadFileResult As Integer
        If hInternet Then
            iReadFileResult = InternetReadBinaryFile(hInternet, sBuffer(0), UBound(sBuffer) - LBound(sBuffer), lngDataReturned)
            ReDim Preserve sBuffer(lngDataReturned - 1)
            oStream.Write(sBuffer)
            ReDim sBuffer(bufSize)
            totalRead = totalRead + lngDataReturned

            Do While lngDataReturned <> 0
                iReadFileResult = InternetReadBinaryFile(hInternet, sBuffer(0), UBound(sBuffer) - LBound(sBuffer), lngDataReturned)
                If lngDataReturned = 0 Then Exit Do

                ReDim Preserve sBuffer(lngDataReturned - 1)
                oStream.Write(sBuffer)
                ReDim sBuffer(bufSize)
                totalRead = totalRead + lngDataReturned
            Loop


SAVE:       Try
                oStream.SaveToFile(filePath, IIf(overWriteFile, 2, 1))
                oStream.Close()
            Catch
ERR:            Dim tConfirm As Integer
                tConfirm = MsgBox("Can't save file: " + FileName + ". Is file opening? Please close." + vbNewLine + "-> Yes: Continue" + vbNewLine + "-> No: Save new file", vbYesNoCancel)
                If tConfirm = vbYes Then
                ElseIf tConfirm = vbNo Then
                    filePath = Replace_f(filePath, "(.+[\/\\])(.+)", "$1")
                    Dim fso As New FileSystemObject
                    Dim No As Integer
                    No = 1
                    Dim newFileName As String
                    newFileName = FileName
                    Do While fso.FileExists(filePath + newFileName)
                        newFileName = Replace_f(FileName, "(.+)(\.)", "$1(" + CStr(No) + ")$2")
                        No = No + 1
                    Loop
                    filePath = filePath + newFileName
                Else
                    Exit Sub
                End If
                GoTo SAVE
            End Try
        End If
        Call InternetCloseHandle(hInternet)

    End Sub

    Sub ILM_DownloadItemInFolder(ByVal Path_ILM As String, ByVal DesFolder As String, ByVal consoletextbox As MetroFramework.Controls.MetroTextBox)
        Dim Path As String
        Dim savePlace As String
        Path = Path_ILM
        savePlace = DesFolder
        Dim listItems As Collection
        listItems = GetListOfSubProject(Path)
        Dim fso As New FileSystemObject
        If fso.FolderExists(savePlace) = False Then
            fso.CreateFolder(savePlace)
        End If
        If listItems.Count > 0 Then
            For Each item In listItems
                If item <> "" Then
                    Dim sUrl As String
                    Dim sSaveFile As String
                    Dim collectioncheck As String
                    collectioncheck = Replace_f(item, "path=(.+);iscollection=(.+)", "$2")
                    item = Replace_f(item, "path=(.+);iscollection=(.+)", "$1")
                    If collectioncheck <> "true" Then
                        sUrl = DOWNLOAD_FILE + Path + "/" + item
                        sSaveFile = savePlace + "\" + item
                        DownloadFile(sUrl, sSaveFile, consoletextbox, True)
                    Else
                        Call ILM_DownloadItemInFolder(Path + "/" + item, savePlace + "\" + item, consoletextbox)
                    End If
                End If
            Next item
        End If
    End Sub
End Module
