Imports System.IO
Imports System.Net

Public Class Main
    Const SetupFileName As String = "setup.exe"
    Const ConfigFileName As String = "configuration.xml"
    Dim SetupURL As String = "https://github.com/asheroto/Deploy-Office-2019/releases/latest/download/setup.exe"
    Dim ConfigURL As String = "https://github.com/asheroto/Deploy-Office-2019/releases/latest/download/configuration.xml"

    Sub Download(URL As String, FileName As String)
        Dim DownloadPath As String = Path.Combine(Application.StartupPath, FileName)

        If File.Exists(DownloadPath) Then
            Try
                File.Delete(DownloadPath)
            Catch ex As Exception
                MsgBox(String.Format("Failed to delete existing {0}. Please end all setup.exe processes or restart your computer and try again.", FileName), vbExclamation)
                End
            End Try

        End If

        Try
            Dim wc As WebClient = New WebClient()
            wc.DownloadFile(SetupURL, DownloadPath)
        Catch ex As Exception
            MsgBox(String.Format("Failed to download {0}. Please ensure you have Internet connectivity or restart your computer and try again.", FileName), vbExclamation)
            End
        End Try

        If File.Exists(DownloadPath) = False Then
            MsgBox(String.Format("Problem downloading {0}. Please ensure you have Internet connectivity or restart your computer and try again.", FileName), vbExclamation)
        End If

    End Sub

    Private Sub Main_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Download setup.exe
        Download(SetupURL, SetupFileName)

        'Download configuration.xml
        Download(ConfigURL, ConfigFileName)

        End
    End Sub

End Class