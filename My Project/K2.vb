Imports Microsoft.VisualBasic.FileIO
Imports Microsoft.Office.Interop.Excel
Imports System.IO
Imports Microsoft.Office.Interop
Imports System.Windows.Forms
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
        ' Define file paths
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

        ' Open the CSV file
        Using parser As New TextFieldParser(csvFilePath)
            parser.TextFieldType = FieldType.Delimited
            parser.SetDelimiters(",")

            ' Reference to K2 Extract sheet
            Dim xlApp As New Excel.Application
            Dim wb As Excel.Workbook = xlApp.Workbooks.Open(wbNamePath)
            Dim wsK2 As Excel.Worksheet = CType(wb.Sheets("K2 Extract"), Excel.Worksheet)
            Dim totalRows As Integer = File.ReadAllLines(csvFilePath).Length
            ProgressBar1.Minimum = 0
            ProgressBar1.Maximum = totalRows
            ProgressBar1.Value = 0
            ProgressBar1.Step = 1
            ' Copy data from CSV to K2 Extract sheet
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

                'parser.DoEvents()
            End While

            ' Close the workbook
            wb.Close(SaveChanges:=True)
            xlApp.Quit()
        End Using
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

        ' Open the CSV file
        Using parser As New TextFieldParser(csvFilePath)
            parser.TextFieldType = FieldType.Delimited
            parser.SetDelimiters(",")

            Dim xlApp As New Excel.Application

            ' Reference to CCD Extract sheet
            'MsgBox("try" & WorkingDirectoryPath & "wbname" & wbName)
            Dim wb As Excel.Workbook = xlApp.Workbooks.Open(WorkingDirectoryPath, wbName)
            Dim wsCCD As Excel.Worksheet = CType(wb.Sheets("CCD Extract"), Excel.Worksheet)

            ' Set the data range in the CSV file
            Dim rowIndex As Integer = 1
            Dim totalRows As Integer = File.ReadAllLines(csvFilePath).Length
            ProgressBar1.Minimum = 0
            ProgressBar1.Maximum = totalRows
            ProgressBar1.Value = 0
            ProgressBar1.Step = 1

            While Not parser.EndOfData
                Dim csvDataRange As String() = parser.ReadFields()
                For Each field As String In csvDataRange
                    wsCCD.Cells(rowIndex, 1).Value = field
                    rowIndex += 1
                    ProgressBar1.PerformStep()
                Next

            End While
        End Using
    End Sub
End Module
