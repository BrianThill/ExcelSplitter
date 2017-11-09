Imports Microsoft.VisualBasic.FileIO
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports System.IO

Module Main
    Sub ExtractFirstWorkSheetToCSV(App As Excel.Application, full_fname As String, trgDir As String)
        Dim wbSource, wbTarget As Excel.Workbook
        Dim ws As Excel.Worksheet

        Dim srcInfo As FileInfo
        Dim dirname, fname, ext As String

        srcInfo = New FileInfo(full_fname)
        If srcInfo.Extension = ".xlsx" Or srcInfo.Extension = ".xls" Then
            fname = srcInfo.Name
            dirName = srcInfo.DirectoryName
            ext = srcInfo.Extension
            Try
                wbSource = App.Workbooks.Open(full_fname)
                If wbSource IsNot Nothing Then
                    For Each ws In wbSource.Worksheets
                        fname = Replace(fname, ext, ".csv", Count:=1)
                        wbTarget = App.Workbooks.Add()
                        ws.Copy(After:=wbTarget.Worksheets(1))
                        wbTarget.SaveAs(trgDir & "\" & fname, xlCSV)
                        Exit For   ' only process the first worksheet for this usage
                    Next
                End If
            Catch ex As Exception
            Finally
                If wbSource IsNot Nothing Then
                    wbSource.Saved = True
                    wbSource.Close(SaveChanges:=False)
                End If
                If wbTarget IsNot Nothing Then
                    wbTarget.Close()
                End If

            End Try

        End If

    End Sub
    Sub Main(args As String())

        Dim full_fname As String = ""
        Dim srcDir, trgDir As String


        Select Case args.Count
            Case 0
                srcDir = Directory.GetCurrentDirectory
                trgDir = Directory.GetCurrentDirectory
            Case 1
                srcDir = args(0)
                trgDir = srcDir
            Case 2
                srcDir = args(0)
                trgDir = args(1)
            Case Else
                Console.WriteLine("Invalid parameters")
        End Select

        Dim App As New Excel.Application

        Try
            For Each full_fname In System.IO.Directory.EnumerateFiles(srcDir)
                full_fname = System.IO.Path.GetFullPath(full_fname)
                ExtractFirstWorkSheetToCSV(App, full_fname, trgDir)
            Next

        Catch ex As Exception
            Console.WriteLine("Error processing file: " & full_fname)
        Finally
            App.Quit()
            Environment.Exit(0)
        End Try

    End Sub

End Module
