Imports Microsoft.VisualBasic.FileIO
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports System.IO

Module Main

    Sub Main(args As String())

        Dim fname, path, full_fname As String
        Dim wbSource, wbTarget As Excel.Workbook
        Dim ws As Excel.Worksheet
        Dim App As New Excel.Application
        If args.Count = 1 Then
            fname = args(0)
        Else
            Console.WriteLine("Input filename not provided")
        End If

        path = Directory.GetCurrentDirectory
        full_fname = System.IO.Path.GetFullPath(fname)

        If Not System.IO.File.Exists(full_fname) Then
            full_fname = ""
            Console.WriteLine("Input file not found: " & fname)
        End If


        If full_fname <> "" Then

            Try
                wbSource = App.Workbooks.Open(full_fname)
                If wbSource IsNot Nothing Then
                    For Each ws In wbSource.Worksheets
                        fname = path & "\" & ws.Name & ".csv"
                        wbTarget = App.Workbooks.Add()
                        ws.Copy(After:=wbTarget.Worksheets(1))
                        wbTarget.SaveAs(fname, xlCSV)
                    Next
                End If



            Catch ex As Exception
            Finally
                If wbSource IsNot Nothing Then
                    wbSource.Saved = True
                    wbSource.Close(SaveChanges:=False)
                End If
                App.Quit()
                Environment.Exit(0)
            End Try


        End If

    End Sub

End Module
