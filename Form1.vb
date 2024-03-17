Public Class Form1
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
End Class
