Imports System.Globalization
Imports System.IO
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Windows.Forms ' Import necessary namespace for ProgressBar


Module LATAM
    Dim documentPath As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)

    Dim rootPath = Path.Combine(documentPath, "scotia-automation")

    Sub PopulateReportFromCalculationFile(ProgressBar1 As ProgressBar)
        EnsureCreation(rootPath)

        Dim CalculationFile As String
        Dim ReportFile As String
        Dim fwdSheet As Excel.Worksheet
        Dim swapSheet As Excel.Worksheet
        Dim reportSheet As Excel.Worksheet
        Dim lastRowFwd As Long
        Dim lastRowSwap As Long
        Dim lastRowReport As Long
        Dim i As Long
        Dim j As Long
        Dim counterparty As String
        Dim notional As Double
        Dim foundMatch As Boolean
        Dim calcWorkbook As Excel.Workbook
        Dim reportWorkbook As Excel.Workbook
        Dim gCountry As String
        Dim currentDate As DateTime = DateTime.Now
        Dim currentYear As String = currentDate.ToString("yyyy")
        Dim prevMonth As String = currentDate.AddMonths(-1).ToString("MMM")
        Dim AssemblyDirectory = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location)

        Dim MasterReportFileName = "SupportdataforMINIMIS Report"

        Dim MasterReportFilePath = System.IO.Path.Combine(AssemblyDirectory, MasterReportFileName & ".xlsx")
        ' Define file paths
        Dim WorkingDirectoryPath = System.IO.Path.Combine(rootPath, $"{currentYear}\{prevMonth}\Latam De Minimis Calculation")
        EnsureCreation(WorkingDirectoryPath)
        Dim formattedDate As String = $"January 1, {currentYear} to December 31, {currentYear}"
        Dim ReportFileName = MasterReportFileName & " " & formattedDate & ".xlsx" ' Use same workbook for ReportFile
        Dim ReportFilePath = System.IO.Path.Combine(WorkingDirectoryPath, ReportFileName)
        'EnsureCreation(ReportFilePath, method:="file")

        Try

            File.Copy(MasterReportFilePath, ReportFilePath, True)
        Catch ex As Exception
            MsgBox("Sorry, Couldn't prepare files" & ex.Message)
        End Try
        'Dim ReportFilePath = System.IO.Path.Combine(AssemblyFile, ReportFileName)
        Dim CalculationFileName = $"CFTC Deminimis LatAm Extracts\MINIMIS Calculation Template (Chile) {formattedDate}.xlsx"
        Dim CalculationFilePath = System.IO.Path.Combine(WorkingDirectoryPath, CalculationFileName)

        Dim excelApp As New Excel.Application()
        excelApp.Visible = False ' You can set this to True to see Excel operations happening.

        ' Check if calculation file is already open
        Try
            calcWorkbook = excelApp.Workbooks(CalculationFilePath)
        Catch ex As Exception
            calcWorkbook = excelApp.Workbooks.Open(CalculationFilePath)
        End Try

        ' Check if report file is already open
        Try
            reportWorkbook = excelApp.Workbooks(ReportFilePath)
        Catch ex As Exception
            reportWorkbook = excelApp.Workbooks.Open(ReportFilePath)
        End Try


        ' Set sheet references
        fwdSheet = calcWorkbook.Sheets("FWD")
        swapSheet = calcWorkbook.Sheets("SWAP")
        reportSheet = reportWorkbook.Sheets("Clients")

        ' Find last row for each sheet
        lastRowFwd = fwdSheet.Cells(fwdSheet.Rows.Count, "H").End(Excel.XlDirection.xlUp).Row
        lastRowSwap = swapSheet.Cells(swapSheet.Rows.Count, "M").End(Excel.XlDirection.xlUp).Row
        lastRowReport = reportSheet.Cells(reportSheet.Rows.Count, "C").End(Excel.XlDirection.xlUp).Row

        ' Set up progress bar
        Dim totalRows As Long = lastRowFwd + lastRowSwap
        ProgressBar1.Minimum = 0
        ProgressBar1.Maximum = totalRows
        ProgressBar1.Value = 0
        ProgressBar1.Step = 1

        ' Loop through SWAP sheet data
        For i = 4 To lastRowSwap
            counterparty = swapSheet.Cells(i, "M").Value
            notional = swapSheet.Cells(i, "O").Value
            foundMatch = False

            '' Check if country and product already exist in report
            'For j = 2 To lastRowReport
            '    If reportSheet.Cells(j, "A").Value = swapSheet.Cells(i, "A").Value And reportSheet.Cells(j, "B").Value = "Swaps" Then
            '        foundMatch = True
            '        reportSheet.Cells(j, "A").Value = "in"
            '        reportSheet.Cells(j, "B").Value = "Swaps"
            '        reportSheet.Cells(j, "C").Value = counterparty
            '        reportSheet.Cells(j, "D").Value = notional
            '        Exit For
            '    End If

            'Next j

            ' If not found, add new row
            If Not foundMatch Then
                lastRowReport = lastRowReport + 1
                reportSheet.Cells(lastRowReport, "A").Value = swapSheet.Cells(i, "A").Value
                reportSheet.Cells(lastRowReport, "B").Value = "Swaps"
                reportSheet.Cells(lastRowReport, "C").Value = counterparty
                reportSheet.Cells(lastRowReport, "D").Value = notional
            End If
            ' Update progress bar

            ProgressBar1.PerformStep()
            Application.DoEvents() ' Allow the UI to update
        Next i

        ' Loop through FWD sheet data
        For i = 4 To lastRowFwd
            gCountry = fwdSheet.Cells(i, "M").Value
            counterparty = fwdSheet.Cells(i, "H").Value
            notional = fwdSheet.Cells(i, "J").Value
            foundMatch = False

            '' Check if country and product already exist in report
            'For j = i - 2 To lastRowReport
            '    foundMatch = True

            '    reportSheet.Cells(j, "C").Value = counterparty
            '    reportSheet.Cells(j, "D").Value = notional
            '    Exit For
            'Next j

            ' If not found, add new row
            If Not foundMatch Then
                lastRowReport = lastRowReport + 1
                reportSheet.Cells(lastRowReport, "A").Value = gCountry
                reportSheet.Cells(lastRowReport, "B").Value = "Forwards"
                reportSheet.Cells(lastRowReport, "C").Value = counterparty
                reportSheet.Cells(lastRowReport, "D").Value = notional
            End If
            ' Update progress bar
            ProgressBar1.PerformStep()
            Application.DoEvents() ' Allow the UI to update
        Next i





        ' Close calculation file
        calcWorkbook.Close(SaveChanges:=False)
        reportWorkbook.Close(SaveChanges:=True)
        excelApp.Quit()

        ' Update progress bar to finish
        ProgressBar1.Value = ProgressBar1.Maximum
        Application.DoEvents() ' Allow the UI to update
        MsgBox("DONE")

        ' Release COM objects
        ReleaseComObject(fwdSheet)
        ReleaseComObject(swapSheet)
        ReleaseComObject(reportSheet)
        ReleaseComObject(calcWorkbook)
        ReleaseComObject(excelApp)
    End Sub

    Private Sub ReleaseComObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
        Catch ex As Exception
        Finally
            obj = Nothing
        End Try
    End Sub


    Public Function EnsureCreation(path As String, Optional ByVal method As String = "dir") As Boolean


        If method = "dir" Then
            Try
                If Not Directory.Exists(path) Then
                    Dim directoryInfo As New DirectoryInfo(path)
                    directoryInfo.Create() ' Create directory with intermediate directories if needed
                End If
                Return True
            Catch ex As Exception
                Console.WriteLine($"Error creating directory: {path} ({ex.Message})")
                Return False
            End Try
        ElseIf method = "file" Then
            Try
                Dim fileStream As New FileStream(path, FileMode.Create)
                fileStream.Close() ' Create an empty file
                Return True
            Catch ex As Exception
                Console.WriteLine($"Error creating file: {path} ({ex.Message})")
                Return False
            End Try
        Else
            Console.WriteLine($"Invalid method: {method}. Supported methods are 'dir' and 'file'.")
            Return False
        End If
    End Function



End Module
