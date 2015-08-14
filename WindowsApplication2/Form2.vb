Imports Microsoft.Office.Interop
Public Class Form2
    Private dt As DataTable
    Private MaxLevel As Integer
    Public TextPlanSQL As String
    Dim ArrayString(10) As PlanRow

    Public Class PlanRow
        Public s1 As String
        Public s2 As String
        Public s3 As String
        Public s4 As String
        Public s5 As String
        Public s6 As String
        Public s7 As String
        Public s8 As String
        Public Level As Integer
        Public Text As String
    End Class

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Close()
    End Sub

    Private Sub InitForm()
        Dim gdParent As Guid

        dt = New DataTable()

        FillTableDefinition()

        TextBox1.Text = TextPlanSQL

        Dim Delimeter As Integer = 0
        Dim ArrayGUID(100) As System.Guid

        Dim Level As Integer = 0

        ReDim ArrayString(TextBox1.Lines.Length)
        Dim f As PlanRow
        Dim Count As Integer = 0

        For Each ss In TextBox1.Lines

            Delimeter = ss.IndexOf(",")
            Dim s1 As String = ss.Substring(0, Delimeter).Trim
            ss = ss.Substring(Delimeter + 1)

            Delimeter = ss.IndexOf(",")
            Dim s2 As String = ss.Substring(0, Delimeter).Trim
            ss = ss.Substring(Delimeter + 1)

            Delimeter = ss.IndexOf(",")
            Dim s3 As String = ss.Substring(0, Delimeter).Trim
            ss = ss.Substring(Delimeter + 1)

            Delimeter = ss.IndexOf(",")
            Dim s4 As String = ss.Substring(0, Delimeter).Trim
            ss = ss.Substring(Delimeter + 1)

            Delimeter = ss.IndexOf(",")
            Dim s5 As String = ss.Substring(0, Delimeter).Trim
            ss = ss.Substring(Delimeter + 1)

            Delimeter = ss.IndexOf(",")
            Dim s6 As String = ss.Substring(0, Delimeter).Trim
            ss = ss.Substring(Delimeter + 1)

            Delimeter = ss.IndexOf(",")
            Dim s7 As String = ss.Substring(0, Delimeter).Trim
            ss = ss.Substring(Delimeter + 1)

            Delimeter = ss.IndexOf(",")
            Dim s8 As String = ss.Substring(0, Delimeter).Trim
            ss = ss.Substring(Delimeter + 4) 'Обрежем также 3 лишних пробела


            Delimeter = ss.IndexOf("|--")
            ss = ss.Substring(Delimeter + 3)

            Level = Delimeter / 5
            If Level < 1 Then
                Level = 0
                gdParent = FillTable("", s1, s2, s3, s4, s5, s6, s7, s8, ss)
                ArrayGUID(Level) = gdParent
            Else
                gdParent = FillTable(ArrayGUID(Level - 1), s1, s2, s3, s4, s5, s6, s7, s8, ss)
                ArrayGUID(Level) = gdParent
            End If

            If MaxLevel < Level Then
                MaxLevel = Level
            End If

            f = New PlanRow
            f.s1 = s1
            f.s2 = s2
            f.s3 = s3
            f.s4 = s4
            f.s5 = s5
            f.s6 = s6
            f.s7 = s7
            f.s8 = s8
            f.Level = Level
            f.Text = ss

            ArrayString(Count) = f

            Count += 1

        Next

        ArrayGUID = Nothing

        dt.DefaultView.Sort = "EstimateCPU"

        InitHierarchy()

    End Sub
    Private Sub InitHierarchy()
        hierarchicalDataGridView1.PrimaryKeyColumnName = "id"
        hierarchicalDataGridView1.HierarchicalColumnName = "id_parent"
        hierarchicalDataGridView1.RootValue = ""
        hierarchicalDataGridView1.DataSource = dt
    End Sub

    Private Function FillTable(ByVal gd As Object,
                               ByVal Rows As String,
                               ByVal Executes As String,
                               ByVal EstimateRows As String,
                               ByVal EstimateIO As String,
                               ByVal EstimateCPU As String,
                               ByVal AvgRowSize As String,
                               ByVal TotalSubtreeCost As String,
                               ByVal EstimateExecutions As String,
                               ByVal StmtText As String) As Guid
        Dim dr As DataRow
        Dim gdc As Guid

        dr = dt.NewRow()
        gdc = Guid.NewGuid
        dr("id") = gdc
        dr("Rows") = Rows
        dr("Executes") = Executes
        dr("EstimateRows") = EstimateRows
        dr("EstimateIO") = EstimateIO
        dr("EstimateCPU") = EstimateCPU
        dr("AvgRowSize") = AvgRowSize
        dr("TotalSubtreeCost") = TotalSubtreeCost
        dr("EstimateExecutions") = EstimateExecutions
        dr("StmtText") = StmtText
        dr("id_parent") = gd
        dt.Rows.Add(dr)

        Return gdc
    End Function

    Private Sub FillTableDefinition()
        Dim dc As DataColumn

        dc = New DataColumn()
        dc.DataType = System.Type.GetType("System.String")
        dc.ColumnName = "id"
        dt.Columns.Add(dc)

        dc = New DataColumn()
        dc.DataType = System.Type.GetType("System.String")
        dc.ColumnName = "Rows"
        dt.Columns.Add(dc)

        dc = New DataColumn()
        dc.DataType = System.Type.GetType("System.String")
        dc.ColumnName = "Executes"
        dt.Columns.Add(dc)

        dc = New DataColumn()
        dc.DataType = System.Type.GetType("System.String")
        dc.ColumnName = "EstimateRows"
        dt.Columns.Add(dc)

        dc = New DataColumn()
        dc.DataType = System.Type.GetType("System.String")
        dc.ColumnName = "EstimateIO"
        dt.Columns.Add(dc)

        dc = New DataColumn()
        dc.DataType = System.Type.GetType("System.String")
        dc.ColumnName = "EstimateCPU"
        dt.Columns.Add(dc)

        dc = New DataColumn()
        dc.DataType = System.Type.GetType("System.String")
        dc.ColumnName = "AvgRowSize"
        dt.Columns.Add(dc)

        dc = New DataColumn
        dc.DataType = System.Type.GetType("System.String")
        dc.ColumnName = "TotalSubtreeCost"
        dt.Columns.Add(dc)

        dc = New DataColumn
        dc.DataType = System.Type.GetType("System.String")
        dc.ColumnName = "EstimateExecutions"
        dt.Columns.Add(dc)

        dc = New DataColumn
        dc.DataType = System.Type.GetType("System.String")
        dc.ColumnName = "StmtText"
        dt.Columns.Add(dc)

        dc = New DataColumn
        dc.DataType = System.Type.GetType("System.String")
        dc.ColumnName = "id_parent"
        dt.Columns.Add(dc)

    End Sub

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        InitForm()
    End Sub



    Private Sub hierarchicalDataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles hierarchicalDataGridView1.CellContentClick

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value


        SaveFileDialog.ShowDialog()

        xlApp = New Excel.ApplicationClass
        xlWorkBook = xlApp.Workbooks.Add(misValue)
        xlWorkSheet = xlWorkBook.Sheets(1)

        xlWorkSheet.Columns(1).NumberFormat = "0,00E+00"
        xlWorkSheet.Columns(2).NumberFormat = "0,00E+00"
        xlWorkSheet.Columns(3).NumberFormat = "0,00E+00"
        xlWorkSheet.Columns(4).NumberFormat = "0,00E+00"
        xlWorkSheet.Columns(5).NumberFormat = "0,00E+00"
        xlWorkSheet.Columns(6).NumberFormat = "0,00E+00"
        xlWorkSheet.Columns(7).NumberFormat = "0,00E+00"
        xlWorkSheet.Columns(8).NumberFormat = "0,00E+00"

        xlWorkSheet.Cells(1, 1) = "Rows"
        xlWorkSheet.Cells(1, 2) = "Executes"
        xlWorkSheet.Cells(1, 3) = "EstimateRows"
        xlWorkSheet.Cells(1, 4) = "EstimateIO"
        xlWorkSheet.Cells(1, 5) = "EstimateCPU"
        xlWorkSheet.Cells(1, 6) = "AvgRowSize"
        xlWorkSheet.Cells(1, 7) = "TotalSubtreeCost"
        xlWorkSheet.Cells(1, 8) = "EstimateExecutions"
        'xlWorkSheet.Cells(1, 9) = "Level"
        For i As Integer = 9 To MaxLevel + 9
            xlWorkSheet.Columns(i).ColumnWidth = 0.83
        Next

        xlWorkSheet.Cells(1, 9) = "StmtText"

        Dim Count As Integer = 2
        For Each ss In ArrayString
            If Not ss Is Nothing Then
                xlWorkSheet.Cells(Count, 1) = ss.s1.Replace(".", ",")
                xlWorkSheet.Cells(Count, 2) = ss.s2.Replace(".", ",")
                xlWorkSheet.Cells(Count, 3) = ss.s3.Replace(".", ",")
                xlWorkSheet.Cells(Count, 4) = ss.s4.Replace(".", ",")
                xlWorkSheet.Cells(Count, 5) = ss.s5.Replace(".", ",")
                xlWorkSheet.Cells(Count, 6) = ss.s6.Replace(".", ",")
                xlWorkSheet.Cells(Count, 7) = ss.s7.Replace(".", ",")
                xlWorkSheet.Cells(Count, 8) = ss.s8.Replace(".", ",")
                'xlWorkSheet.Cells(Count, 9) = ss.Level
                xlWorkSheet.Cells(Count, ss.Level + 9) = ss.Text
                Count += 1
            End If
        Next

        xlWorkSheet.Columns(1).NumberFormat = "0,00"
        xlWorkSheet.Columns(2).NumberFormat = "0,00"
        xlWorkSheet.Columns(3).NumberFormat = "0,00"
        xlWorkSheet.Columns(4).NumberFormat = "0,00"
        xlWorkSheet.Columns(5).NumberFormat = "0,00"
        xlWorkSheet.Columns(6).NumberFormat = "0,00"
        xlWorkSheet.Columns(7).NumberFormat = "0,00"
        xlWorkSheet.Columns(8).NumberFormat = "0,00"

        xlWorkSheet.Columns("A:H").EntireColumn.AutoFit()

        xlWorkSheet.SaveAs(SaveFileDialog.FileName)
        xlWorkBook.Close(True)
        xlApp.Quit()

        releaseObject(xlApp)
        releaseObject(xlWorkBook)
        releaseObject(xlWorkSheet)

        MsgBox("Выгрузка завершена!")
    End Sub
    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub
End Class