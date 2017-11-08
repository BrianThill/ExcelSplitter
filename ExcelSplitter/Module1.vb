Imports Microsoft.VisualBasic.FileIO
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel.XlFileFormat


Module Module1

    Sub Main()
        Dim cnt As Integer
        Dim fname As String
        Dim wbSource, wbTarget As Excel.Workbook
        Dim ws As Excel.Worksheet
        Dim App As New Excel.Application
        fname = "c:\temp\test.xlsx"
        If System.IO.File.Exists(fname) Then
            wbSource = App.Workbooks.Open(fname)
            If wbSource IsNot Nothing Then
                For Each ws In wbSource.Worksheets
                    fname = "c:\temp\test" & CStr(cnt) & ".csv"
                    wbTarget = App.Workbooks.Add()
                    ws.Copy(wbTarget)
                    wbTarget.SaveAs(fname, xlCSV)
                Next
            End If
        End If

    End Sub

End Module
