
#Region " Information "
' Class Name : Excel File Handler '
' Programmer : Vivek Purohit '
' Purpose    : Handle Excel File Operations. '
' Date       : 20-Dec-2008'
#End Region

#Region " Import Section"
Imports System
Imports System.Collections.Generic
Imports System.Text
Imports System.Data
Imports System.Data.OleDb
Imports Microsoft.Office.Interop.Excel
Imports System.Reflection
Imports System.Runtime.InteropServices
#End Region


' Excel File handler used to read and write excel file. '
Public Class ExcelHandler

    ' Return data in dataset from excel file. '   
    Public Function GetDataFromExcel(ByVal a_sFilepath As String) As DataSet
        Dim ds As New DataSet()
        Dim cn As New System.Data.OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & a_sFilepath & ";Extended Properties= Excel 8.0")

        ' For xls files try it with this Provider=Microsoft.ACE.OLEDB.12.0;Data Source=myFile.xls;Extended Properties=""Excel 8.0;IMEX=1;HDR=YES;""

        Try
            cn.Open()
        Catch ex As OleDbException
            Console.WriteLine(ex.Message)
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try

        ' It Represents Excel data table Schema.'
        Dim dt As New System.Data.DataTable()
        dt = cn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, Nothing)
        If dt IsNot Nothing OrElse dt.Rows.Count > 0 Then
            For sheet_count As Integer = 0 To dt.Rows.Count - 1
                Try
                    ' Create Query to get Data from sheet. '
                    Dim sheetname As String = dt.Rows(sheet_count)("table_name").ToString()
                    Dim da As New OleDbDataAdapter("SELECT * FROM [" & sheetname & "]", cn)
                    da.Fill(ds, sheetname)
                Catch ex As DataException
                    Console.WriteLine(ex.Message)
                Catch ex As Exception
                    Console.WriteLine(ex.Message)
                End Try
            Next
        End If
        cn.Close()
        Return ds
    End Function


    ' Write Excel file as given file name with given data.'  

    Public Function ExportToExcel(ByVal a_sFilename As String, ByVal a_sData As DataSet, ByVal a_sFileTitle As String, ByRef a_sErrorMessage As String) As Boolean
        a_sErrorMessage = String.Empty
        Dim bRetVal As Boolean = False
        Dim dsDataSet As DataSet = Nothing
        Try
            dsDataSet = a_sData

            Dim xlObject As Microsoft.Office.Interop.Excel.Application = Nothing
            Dim xlWB As Microsoft.Office.Interop.Excel.Workbook = Nothing
            Dim xlSh As Microsoft.Office.Interop.Excel.Worksheet = Nothing
            Dim rg As Range = Nothing
            Try
                xlObject = New Microsoft.Office.Interop.Excel.Application()
                xlObject.AlertBeforeOverwriting = False
                xlObject.DisplayAlerts = False

                ' This Adds a new woorkbook, you could open the workbook from file also '
                xlWB = xlObject.Workbooks.Add(Type.Missing)
                xlWB.SaveAs(a_sFilename, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, _
                Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Missing.Value, Missing.Value, Missing.Value, Missing.Value)

                xlSh = DirectCast(xlObject.ActiveWorkbook.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)

                Dim sUpperRange As String = "A1"
                Dim sLastCol As String = "E"
                Dim sLowerRange As String = sLastCol + (dsDataSet.Tables(0).Rows.Count + 1).ToString()

                rg = xlSh.Range(sUpperRange, sLowerRange)
                rg.Value2 = GetData(dsDataSet.Tables(0))

                ' formating '
                xlSh.Range("A1", sLastCol & "1").Font.Bold = True
                xlSh.Range("A1", sLastCol & "1").HorizontalAlignment = XlHAlign.xlHAlignCenter
                xlSh.Range(sUpperRange, sLowerRange).EntireColumn.AutoFit()

                If String.IsNullOrEmpty(a_sFileTitle) Then
                    xlObject.Caption = "untitled"
                Else
                    xlObject.Caption = a_sFileTitle
                End If

                xlWB.Save()
                bRetVal = True
            Catch ex As System.Runtime.InteropServices.COMException
                If ex.ErrorCode = -2147221164 Then
                    a_sErrorMessage = "Error in export: Please install Microsoft Office (Excel) to use the Export to Excel feature."
                ElseIf ex.ErrorCode = -2146827284 Then
                    a_sErrorMessage = "Error in export: Excel allows only 65,536 maximum rows in a sheet."
                Else
                    a_sErrorMessage = (("Error in export: " & ex.Message) + Environment.NewLine & " Error: ") + ex.ErrorCode
                End If
            Catch ex As Exception
                a_sErrorMessage = "Error in export: " & ex.Message
            Finally
                Try
                    If xlWB IsNot Nothing Then
                        xlWB.Close(Nothing, Nothing, Nothing)
                    End If
                    xlObject.Workbooks.Close()
                    xlObject.Quit()
                    If rg IsNot Nothing Then
                        Marshal.ReleaseComObject(rg)
                    End If
                    If xlSh IsNot Nothing Then
                        Marshal.ReleaseComObject(xlSh)
                    End If
                    If xlWB IsNot Nothing Then
                        Marshal.ReleaseComObject(xlWB)
                    End If
                    If xlObject IsNot Nothing Then
                        Marshal.ReleaseComObject(xlObject)
                    End If

                Catch
                End Try
                xlSh = Nothing
                xlWB = Nothing
                xlObject = Nothing
                ' force final cleanup! '
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try
        Catch ex As Exception
            a_sErrorMessage = "Error in export: " & ex.Message
        End Try

        Return bRetVal
    End Function


    ' returns data as two dimentional object array.   ' 
    Private Function GetData(ByVal a_dtData As System.Data.DataTable) As Object(,)
        Dim obj As Object(,) = New Object((a_dtData.Rows.Count + 1) - 1, a_dtData.Columns.Count - 1) {}

        Try
            For j As Integer = 0 To a_dtData.Columns.Count - 1
                obj(0, j) = a_dtData.Columns(j).Caption
            Next

            Dim dt As New DateTime()
            Dim sTmpStr As String = String.Empty

            For i As Integer = 1 To a_dtData.Rows.Count
                For j As Integer = 0 To a_dtData.Columns.Count - 1
                    If a_dtData.Columns(j).DataType Is dt.[GetType]() Then
                        If a_dtData.Rows(i - 1)(j) IsNot DBNull.Value Then
                            DateTime.TryParse(a_dtData.Rows(i - 1)(j).ToString(), dt)
                            obj(i, j) = dt.ToString("MM/dd/yy hh:mm tt")
                        Else
                            obj(i, j) = a_dtData.Rows(i - 1)(j)
                        End If
                    ElseIf a_dtData.Columns(j).DataType Is sTmpStr.[GetType]() Then
                        If a_dtData.Rows(i - 1)(j) IsNot DBNull.Value Then
                            sTmpStr = a_dtData.Rows(i - 1)(j).ToString().Replace(vbCr, "")
                            obj(i, j) = sTmpStr
                        Else
                            obj(i, j) = a_dtData.Rows(i - 1)(j)
                        End If
                    Else
                        obj(i, j) = a_dtData.Rows(i - 1)(j)
                    End If

                Next
            Next
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try
        Return obj
    End Function

End Class


