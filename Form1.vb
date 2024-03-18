Imports System.IO

Public Class Form1
    Public downloadPath As String

    Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

        ' Update UI with settings


        ' Update UI with settings
        Dim documentPath As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        ToolStripStatusLabel1.Text = Path.Combine(documentPath, "scotia-automation")



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



End Class
