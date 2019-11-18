Imports System.Data.OleDb
Imports System.IO

Module Module1
    Sub Main()
        Dim dt As DataTable = Nothing
        Dim pathLoad = "C:\Users\adm.m.gosciniak\Desktop\csv.csv"
        Import(dt, pathLoad, True)
        PrintInConsoleDataTable(dt)
        Dim pathSave = "C:\Users\adm.m.gosciniak\Desktop\csv2.csv"
        Dim message As String = Nothing
        Dim success As Boolean = Nothing
        Export(dt, pathSave, message, success, ";", True)
        Console.ReadKey()
    End Sub
    Sub PrintInConsoleDataTable(ByVal dt As DataTable)
        Dim c As String = Nothing
        For Each col As DataColumn In dt.Columns
            c = c + col.ColumnName + ";"
        Next
        c = c.Remove(c.Length - 1, 1)
        Console.WriteLine(c)

        Dim r As String = Nothing
        For Each row As DataRow In dt.Rows
            Console.WriteLine(String.Join(";", row.ItemArray.Select(Function(i) If(Nothing, i.ToString()))))
        Next
    End Sub
    Sub Split_Path(ByVal File_Path As String, ByRef Folder As String, ByRef File_Name As String)
        Folder = Path.GetDirectoryName(File_Path)
        File_Name = Path.GetFileName(File_Path)
    End Sub
    Sub Import(ByRef Values As DataTable, ByVal CSV_Location As String, ByVal First_Line_Is_Header As Boolean)
        Dim HDRString As String = "No"
        If First_Line_Is_Header Then HDRString = "Yes"


        Dim Folder As String = Nothing
        Dim FileName As String = Nothing
        Split_Path(CSV_Location, Folder, FileName)

        Dim cn As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Folder & ";Extended Properties=""Text;HDR=" & HDRString & ";FMT=Delimited""")
        Dim da As New OleDbDataAdapter()
        Dim ds As New DataSet()
        Dim cd As New OleDbCommand("SELECT * FROM [" & FileName & "]", cn)

        cn.Open()
        da.SelectCommand = cd
        ds.Clear()
        da.Fill(ds, "CSV")
        Values = ds.Tables(0)
        cn.Close()
    End Sub
    Sub Export(ByVal dataCollection As DataTable, ByVal pathToSaveCSV As String, ByRef message As String, ByRef success As Boolean, Optional sourceTextQualifier As String = ",", Optional firstLineIsHeader As Boolean = True)
        Try
            Using sw As StreamWriter = File.CreateText(pathToSaveCSV)
                If (firstLineIsHeader) Then
                    Dim c As String = Nothing
                    For Each col As DataColumn In dataCollection.Columns
                        c = c + col.ColumnName + sourceTextQualifier
                    Next
                    c = c.Remove(c.Length - 1, 1)

                    sw.WriteLine(c.Trim())
                End If

                Dim r As String = Nothing
                For Each row As DataRow In dataCollection.Rows
                    sw.WriteLine(String.Join(sourceTextQualifier, row.ItemArray.Select(Function(i) If(Nothing, i.ToString()))))
                Next
            End Using

            success = True
            message = ""
        Catch ex As Exception
            success = False
            message = ex.Message
        End Try
    End Sub
End Module
