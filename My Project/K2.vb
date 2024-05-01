Imports Microsoft.VisualBasic.FileIO
Imports Microsoft.Office.Interop.Excel
Imports System.IO
Imports Microsoft.Office.Interop
Imports System.Windows.Forms
Imports System.Runtime.InteropServices

Module K2
    Dim documentPath As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
    Dim rootPath = Path.Combine(documentPath, "scotia-automation")

    Sub CFCTE(ProgressBar1 As ProgressBar)
        Dim currentDate As DateTime = DateTime.Now
        Dim currentYear As String = currentDate.ToString("yyyy")
        Dim prevMonth As String = currentDate.AddMonths(-1).ToString("MMM")
        Dim AssemblyDirectory = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location)
        Dim MasterReportFileName = "K2 and Portal Data Summary"
        Dim MasterReportFilePath = System.IO.Path.Combine(AssemblyDirectory, $"{MasterReportFileName}.xlsx")
        Dim WorkingDirectoryPath = System.IO.Path.Combine(rootPath, $"{currentYear}\{prevMonth}\Supporting Files K2 and Murex")
        EnsureCreation(WorkingDirectoryPath)
        Dim formattedDate As String = $"January 1, {currentYear} to December 31, {currentYear}"
        Dim csvFileName As String = $"CFTCExtract_{formattedDate}.csv"
        Dim csvFilePath As String = Path.Combine(WorkingDirectoryPath, "K2", csvFileName)
        Dim wbName As String = $"K2 and Portal Data Summary_{formattedDate}.xlsx"
        Dim wbNamePath As String = Path.Combine(WorkingDirectoryPath, wbName)

        Try
            File.Copy(MasterReportFilePath, wbNamePath, True)
        Catch ex As Exception
            MsgBox("Sorry, Couldn't prepare files" & ex.Message)
        End Try

        Dim CalculationFileName = $"CFTC Deminimis LatAm Extracts\MINIMIS Calculation Template (Chile) {formattedDate}.xlsx"
        Dim CalculationFilePath = ""
        Dim csvFD As New OpenFileDialog()
        Dim WBFD As New OpenFileDialog()

        ' Set the title and filter for the dialog
        csvFD.Title = "Select a csv file"
        csvFD.Filter = "CSV files (*.csv)|*.csv"

        WBFD.Title = "Select a xlsx file"
        WBFD.Filter = "Excel files (*.xlsx)|*.xlsx"

        ' Show the dialog and check if the user clicked OK
        If csvFD.ShowDialog() = DialogResult.OK Then
            ' Get the selected file path
            csvFilePath = csvFD.FileName
            MessageBox.Show("You selected: " & csvFilePath)
        End If

        If WBFD.ShowDialog() = DialogResult.OK Then
            ' Get the selected file path
            wbNamePath = WBFD.FileName
            MessageBox.Show("You selected: " & wbNamePath)
        End If


        Dim xlApp As New Excel.Application
        Dim wb As Excel.Workbook = Nothing
        Dim wsK2 As Excel.Worksheet = Nothing
        Dim parser As TextFieldParser = Nothing

        Try
            parser = New TextFieldParser(csvFilePath)
            parser.TextFieldType = FieldType.Delimited
            parser.SetDelimiters(",")

            wb = xlApp.Workbooks.Open(wbNamePath)
            wsK2 = CType(wb.Sheets("K2 Extract"), Excel.Worksheet)

            Dim totalRows As Integer = File.ReadAllLines(csvFilePath).Length
            ProgressBar1.Minimum = 0
            ProgressBar1.Maximum = totalRows
            ProgressBar1.Value = 0
            ProgressBar1.Step = 1

            Dim currentRow As String()
            Dim rowIndex As Integer = 1

            While Not parser.EndOfData
                currentRow = parser.ReadFields()
                For columnIndex As Integer = 1 To currentRow.Length
                    Select Case columnIndex
                        Case 1 To 9
                            Dim v = currentRow(columnIndex - 1)
                            wsK2.Cells(rowIndex, columnIndex).Value = v
                        Case 10
                            wsK2.Cells(rowIndex, columnIndex + 1).Value = currentRow(columnIndex - 1)
                        Case 11 To 33
                            wsK2.Cells(rowIndex, columnIndex + 2).Value = currentRow(columnIndex - 1)
                        Case Else
                            ' Handle additional columns if needed
                    End Select
                    ProgressBar1.PerformStep()
                Next
                rowIndex += 1
            End While
        Catch ex As Exception
            MsgBox("An error occurred: " & ex.Message)
        Finally
            ' Release COM objects
            ReleaseComObject(wsK2)
            ReleaseComObject(wb)
            ReleaseComObject(xlApp)
            ReleaseComObject(parser)
            ' Call garbage collector to release memory
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try
    End Sub

    Sub CCDExtractCSV(ProgressBar1 As ProgressBar)
        Dim currentDate As DateTime = DateTime.Now
        Dim currentYear As String = currentDate.ToString("yyyy")
        Dim prevMonth As String = currentDate.AddMonths(-1).ToString("MMM")
        Dim formattedDate As String = $"January 1, {currentYear} to December 31, {currentYear}"
        Dim WorkingDirectoryPath = System.IO.Path.Combine(rootPath, $"{currentYear}\{prevMonth}\Supporting Files K2 and Murex")
        EnsureCreation(WorkingDirectoryPath)
        Dim csvFileName As String = "CCD Extract.csv"
        Dim csvFilePath As String = Path.Combine(WorkingDirectoryPath, "K2", csvFileName)
        Dim wbName As String = $"K2 and Portal Data Summary_{formattedDate}.xlsx"

        Dim xlApp As New Excel.Application
        Dim wb As Excel.Workbook = Nothing
        Dim wsCCD As Excel.Worksheet = Nothing
        Dim parser As TextFieldParser = Nothing

        Try
            parser = New TextFieldParser(csvFilePath)
            parser.TextFieldType = FieldType.Delimited
            parser.SetDelimiters(",")

            wb = xlApp.Workbooks.Open(WorkingDirectoryPath & "\" & wbName)
            wsCCD = CType(wb.Sheets("CCD Extract"), Excel.Worksheet)

            Dim totalRows As Integer = File.ReadAllLines(csvFilePath).Length
            ProgressBar1.Minimum = 0
            ProgressBar1.Maximum = totalRows
            ProgressBar1.Value = 0
            ProgressBar1.Step = 1

            Dim rowIndex As Integer = 1

            While Not parser.EndOfData
                Dim csvDataRange As String() = parser.ReadFields()
                For Each field As String In csvDataRange
                    wsCCD.Cells(rowIndex, 1).Value = field
                    rowIndex += 1
                    ProgressBar1.PerformStep()
                Next
            End While
        Catch ex As Exception
            MsgBox("An error occurred: " & ex.Message)
        Finally
            ' Release COM objects
            ReleaseComObject(wsCCD)
            ReleaseComObject(wb)
            ReleaseComObject(xlApp)
            ReleaseComObject(parser)
            ' Call garbage collector to release memory
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try
    End Sub

    Sub ReleaseComObject(ByVal obj As Object)
        If obj IsNot Nothing AndAlso Marshal.IsComObject(obj) Then
            Marshal.ReleaseComObject(obj)
        End If
    End Sub




    Sub EnsureCreation(path As String)
        If Not Directory.Exists(path) Then
            Directory.CreateDirectory(path)
        End If
    End Sub
End Module
