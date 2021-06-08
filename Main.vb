Imports System.IO

Public Class Main

    Sub DownloadSetup()
        Dim SetupPath As String = Path.Combine(Application.StartupPath, "setup.exe")
        Dim SetupURL As String = "https://github.com/asheroto/Deploy-Office-2019/releases/latest/download/setup.exe"

        If File.Exists(SetupPath) Then
            Try
                File.Delete(SetupPath)
            Catch ex As Exception
                MsgBox("Failed to delete existing setup.exe. Please end all setup.exe processes or restart your computer and try again.", vbExclamation)
                End
            End Try

        End If

        Try
            My.Computer.Network.DownloadFile(SetupURL, SetupPath)
        Catch ex As Exception
            MsgBox("Failed to download setup.exe. Please ensure you have Internet connectivity or restart your computer and try again.", vbExclamation)
            End
        End Try

        MsgBox(Application.StartupPath)
    End Sub

    Private Sub Main_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DownloadSetup()
    End Sub

End Class