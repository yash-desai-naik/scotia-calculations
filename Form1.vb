Imports System.IO

Public Class Form1
    Public downloadPath As String

    Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

        ' Update UI with settings
        ' Retrieve download path and interval from settings
        downloadPath = My.Settings.DownloadPath

        ' Update UI with settings
        ToolStripStatusLabel1.Text = downloadPath
        ToolStripStatusLabel1.Text = downloadPath



    End Sub
    Private Sub btnCalculate_Click(sender As Object, e As EventArgs) Handles btnCalculate.Click
        Select Case ComboBox1.SelectedItem.ToString()
            Case "LATAM"
                LATAM.PopulateReportFromCalculationFile(Me.ProgressBar1)
            Case "K2"
                MsgBox("k2")


        End Select
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.SelectedIndex = 0
    End Sub

    Private Sub ProgressBar1_Click(sender As Object, e As EventArgs) Handles ProgressBar1.Click

    End Sub

    Private Sub btnSelectDownloadPath_Click(sender As Object, e As EventArgs) Handles btnSelectDownloadPath.Click
        ' Open folder browser dialog to select download location
        Dim folderBrowserDialog As New FolderBrowserDialog()
        If folderBrowserDialog.ShowDialog() = DialogResult.OK Then
            downloadPath = folderBrowserDialog.SelectedPath
            ToolStripStatusLabel1.Text = downloadPath

            ' Save download path to settings
            My.Settings.DownloadPath = downloadPath
            My.Settings.Save()
        End If
    End Sub


End Class
