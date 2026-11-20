Imports System.IO
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Runtime.InteropServices
Imports Microsoft.Win32
Imports System.Web.Script.Serialization

Module Functions

    Function RunSetup() As Integer
        'Run setup
        Dim SetupPath As String = Path.Combine(Main.TempPath, Main.SetupFilename)
        Dim x As New Process
        x.StartInfo.FileName = SetupPath
        x.StartInfo.Arguments = "/configure configuration.xml"
        x.StartInfo.WorkingDirectory = Main.TempPath
        x.StartInfo.WindowStyle = ProcessWindowStyle.Hidden
        x.StartInfo.CreateNoWindow = True
        x.Start()

        'Wait until installation is finished
        Do Until x.HasExited
            Application.DoEvents()
            Threading.Thread.Sleep(30)
        Loop
        Return x.ExitCode
    End Function

    Sub Cleanup()
        DeleteFileIfExist(Main.SetupPath)
        DeleteFileIfExist(Main.ConfigPath)
    End Sub

    Sub DeleteFileIfExist(path As String)
        Try
            If File.Exists(path) Then
                File.Delete(path)
            End If
        Catch ex As Exception

        End Try
    End Sub

    Sub WaitToClose()
        Dim c As New Stopwatch
        c.Start()
        Do Until c.Elapsed.Seconds >= 30
            Application.DoEvents()
        Loop
        c.Stop()
    End Sub

    Sub OutputError(text As String)
        Main.LogTextbox.AppendText("Error: " & text & vbNewLine)
    End Sub

    Sub OutputAppend(text As String)
        Main.LogTextbox.AppendText(text & vbNewLine)
        WriteLog(text)
    End Sub

    Sub TLSchannelFix()
        'Use TLS 1.2 for this program's own downloads, without changing any system settings
        Net.ServicePointManager.SecurityProtocol = Net.ServicePointManager.SecurityProtocol Or Net.SecurityProtocolType.Tls12
    End Sub

    Private LogWriteFailed As Boolean

    Private LogEnabled As Boolean
    Private PendingLog As String = ""

    ''' <summary>Turns on the log file and writes anything that happened before the decision was made.</summary>
    Sub EnableLog()
        LogEnabled = True
        WriteLogText(PendingLog)
        PendingLog = ""
    End Sub

    Sub WriteLog(message As String)
        Dim line As String = String.Format("{0:yyyy-MM-dd HH:mm:ss} - {1}", DateTime.Now, message) & vbNewLine
        ' The log file is opt-in. Lines are held in memory until then so early settings errors are not lost.
        If LogEnabled Then WriteLogText(line) Else PendingLog &= line
    End Sub

    Private Sub WriteLogText(text As String)
        ' Once the log file proves unwritable (read-only share, etc.), say so once and stop trying
        If LogWriteFailed OrElse text = "" Then Return
        Try
            File.AppendAllText(Path.Combine(Application.StartupPath, Main.ErrorLog), text)
        Catch ex As Exception
            LogWriteFailed = True
            OutputError("Error writing to log file: " & ex.Message)
        End Try
    End Sub

    Public ErrorCount As Integer

    Sub LogError(message As String)
        ErrorCount += 1

        ' Replace newlines in the message to ensure it logs as one line
        Dim singleLineMessage As String = message.Replace(vbCrLf, " ").Replace(vbCr, " ").Replace(vbLf, " ")

        ' Output to log textbox
        OutputError(singleLineMessage)

        ' Log to file
        WriteLog("Error: " & singleLineMessage)
    End Sub

    Function ReadIniFile(path As String) As Dictionary(Of String, String)
        Dim data As New Dictionary(Of String, String)

        Try
            Dim lines() As String = File.ReadAllLines(path)
            For Each line As String In lines
                If Not String.IsNullOrWhiteSpace(line) AndAlso Not line.StartsWith(";") Then
                    Dim parts() As String = line.Split("="c)
                    If parts.Length >= 2 Then
                        Dim key As String = parts(0).Trim().ToLower()
                        Dim value As String = String.Join("=", parts.Skip(1)).Trim()
                        data(key) = value

                        'Debug log
                        Debug.WriteLine(key & " = " & value)
                    Else
                        ' Log lines without valid key-value pairs
                        LogError("Invalid line in INI file: " & line)
                    End If
                End If
            Next
        Catch ex As Exception
            ' Log error if there's an issue reading the INI file
            LogError("Error reading INI file: " & ex.Message)
        End Try

        Return data
    End Function

    ''' <summary>
    ''' Turns command-line parameters like year=2024, /silent=true or --edition=Word into the same
    ''' name/value pairs the INI file produces. Anything that is not name=value is reported and skipped.
    ''' </summary>
    Function ParseArguments(args As IEnumerable(Of String)) As Dictionary(Of String, String)
        Dim data As New Dictionary(Of String, String)
        For Each arg As String In args
            Dim trimmed As String = arg.TrimStart("/"c, "-"c)
            Dim split As Integer = trimmed.IndexOf("="c)
            If split > 0 Then
                data(trimmed.Substring(0, split).Trim().ToLower()) = trimmed.Substring(split + 1).Trim()
            Else
                LogError("Parameter ignored, expected name=value: " & arg)
            End If
        Next
        Return data
    End Function

    Public Function TryStartProcess(command As String) As Boolean
        Try
            Process.Start(command)
            Return True
        Catch ex As Exception
            LogError("Error starting process: " & ex.Message)
            Return False
        End Try
    End Function

    Public Class OfficePathHelper
        Public Shared Function GetOfficePath(appName As String) As String
            Dim officeKeyPath As String = "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\" & appName
            Dim registryKey As RegistryKey = Registry.LocalMachine.OpenSubKey(officeKeyPath)

            Debug.WriteLine("Checking registry key: " & officeKeyPath)

            If registryKey IsNot Nothing Then
                Dim path As Object = registryKey.GetValue("")
                If path IsNot Nothing AndAlso TypeOf path Is String Then
                    Return path.ToString()
                End If
            End If

            Return String.Empty ' Return empty string if path not found or registry key is null
        End Function
    End Class

    Public Class ShortcutHelper
        Public Shared Sub CreateDesktopShortcuts(allUsers As Boolean)
            Dim desktop As String = Environment.GetFolderPath(If(allUsers, Environment.SpecialFolder.CommonDesktopDirectory, Environment.SpecialFolder.Desktop))
            Dim OfficeApps As New Dictionary(Of String, String) From {
                {"WINPROJ.EXE", "Project"},
                {"WINWORD.EXE", "Word"},
                {"EXCEL.EXE", "Excel"},
                {"POWERPNT.EXE", "PowerPoint"},
                {"OUTLOOK.EXE", "Outlook"},
                {"ONENOTE.EXE", "OneNote"},
                {"MSPUB.EXE", "Publisher"},
                {"MSACCESS.EXE", "Access"},
                {"VISIO.EXE", "Visio"}
            }

            For Each exe As KeyValuePair(Of String, String) In OfficeApps
                Dim officePath As String = OfficePathHelper.GetOfficePath(exe.Key)
                If Not String.IsNullOrEmpty(officePath) AndAlso File.Exists(officePath) Then
                    Debug.WriteLine("Creating shortcut for " & exe.Value & " at " & officePath)
                    CreateShortcut(officePath, desktop, exe.Value)
                End If
            Next
        End Sub

        Public Shared Sub CreateShortcut(ByVal TargetName As String, ByVal ShortcutPath As String, ByVal ShortcutName As String)
            Dim oShell As Object = Nothing
            Dim oLink As Object = Nothing

            Try
                Dim ShortcutFullPath As String = Path.Combine(ShortcutPath, ShortcutName & ".lnk")

                oShell = CreateObject("WScript.Shell")
                oLink = oShell.CreateShortcut(ShortcutFullPath)

                oLink.TargetPath = TargetName
                oLink.WindowStyle = 1
                oLink.Save()

                Application.DoEvents()
            Catch ex As Exception
                ' Log or handle the exception as needed
                Debug.WriteLine("Error creating shortcut: " & ex.Message)
            Finally
                ' Cleanup resources
                If oLink IsNot Nothing Then Marshal.FinalReleaseComObject(oLink)
                If oShell IsNot Nothing Then Marshal.FinalReleaseComObject(oShell)
            End Try
        End Sub
    End Class

    Public Class GitHubRepoUpdater
        Private ReadOnly _repoOwner As String
        Private ReadOnly _repoName As String
        Private ReadOnly _httpClient As HttpClient

        Public Sub New(repoOwner As String, repoName As String)
            _repoOwner = repoOwner
            _repoName = repoName
            _httpClient = New HttpClient()
            _httpClient.DefaultRequestHeaders.UserAgent.Add(New ProductInfoHeaderValue("Deploy-Office", "1.0"))
            _httpClient.DefaultRequestHeaders.Accept.Add(New MediaTypeWithQualityHeaderValue("application/vnd.github.v3+json"))
        End Sub

        Public Async Function CheckForLatestReleaseAsync() As Task(Of String)
            Dim url As String = $"https://api.github.com/repos/{_repoOwner}/{_repoName}/releases/latest"

            Try
                Dim response As HttpResponseMessage = Await _httpClient.GetAsync(url)
                response.EnsureSuccessStatusCode()

                Dim jsonResponse As String = Await response.Content.ReadAsStringAsync()
                Dim latestRelease As GitHubRelease = DeserializeJson(Of GitHubRelease)(jsonResponse)

                If latestRelease IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(latestRelease.tag_name) Then
                    Return latestRelease.tag_name
                Else
                    Return ""
                End If
            Catch ex As HttpRequestException
                Return $"HTTP Request Error: {ex.Message}"
            Catch ex As TaskCanceledException
                Return "Request timed out."
            Catch ex As Exception
                Return $"An error occurred: {ex.Message}"
            End Try
        End Function

        Protected Overrides Sub Finalize()
            _httpClient.Dispose()
        End Sub

        Private Shared Function DeserializeJson(Of T)(json As String) As T
            Dim serializer As New JavaScriptSerializer()
            Return serializer.Deserialize(Of T)(json)
        End Function

        Private Class GitHubRelease
            Public Property tag_name As String
        End Class
    End Class



End Module