Imports System.Threading
Imports System.IO
Imports System.Data.SqlClient
Imports Microsoft.Office.Interop
Imports System.Text.RegularExpressions
Imports System.Collections
Imports System.Diagnostics
Imports Newtonsoft.Json
Imports System.Net
Imports System.Text


Public Class Form1
    Dim MaxCountRowFromEvent As Integer = 1000
    Dim ParallelsThreads As Integer = 10 ' 0 to 9
    Dim FinishedThreads As Integer
    Dim fibArray(ParallelsThreads) As TJReader.TJ.LoaderTj
    Dim Param As IniFile

    Public Class IniFile
        ' List of IniSection objects keeps track of all the sections in the INI file
        Private m_sections As Hashtable

        ' Public constructor
        Public Sub New()
            m_sections = New Hashtable(StringComparer.InvariantCultureIgnoreCase)
        End Sub

        ' Loads the Reads the data in the ini file into the IniFile object
        Public Sub Load(ByVal sFileName As String, Optional ByVal bMerge As Boolean = False)
            If Not bMerge Then
                RemoveAllSections()
            End If
            '  Clear the object... 
            Dim tempsection As IniSection = Nothing
            Dim oReader As New StreamReader(sFileName)
            Dim regexcomment As New Regex("^([\s]*#.*)", (RegexOptions.Singleline Or RegexOptions.IgnoreCase))
            ' Broken but left for history
            'Dim regexsection As New Regex("\[[\s]*([^\[\s].*[^\s\]])[\s]*\]", (RegexOptions.Singleline Or RegexOptions.IgnoreCase))
            Dim regexsection As New Regex("^[\s]*\[[\s]*([^\[\s].*[^\s\]])[\s]*\][\s]*$", (RegexOptions.Singleline Or RegexOptions.IgnoreCase))
            Dim regexkey As New Regex("^\s*([^=\s]*)[^=]*=(.*)", (RegexOptions.Singleline Or RegexOptions.IgnoreCase))
            While Not oReader.EndOfStream
                Dim line As String = oReader.ReadLine()
                If line <> String.Empty Then
                    Dim m As Match = Nothing
                    If regexcomment.Match(line).Success Then
                        m = regexcomment.Match(line)
                        Trace.WriteLine(String.Format("Skipping Comment: {0}", m.Groups(0).Value))
                    ElseIf regexsection.Match(line).Success Then
                        m = regexsection.Match(line)
                        Trace.WriteLine(String.Format("Adding section [{0}]", m.Groups(1).Value))
                        tempsection = AddSection(m.Groups(1).Value)
                    ElseIf regexkey.Match(line).Success AndAlso tempsection IsNot Nothing Then
                        m = regexkey.Match(line)
                        Trace.WriteLine(String.Format("Adding Key [{0}]=[{1}]", m.Groups(1).Value, m.Groups(2).Value))
                        tempsection.AddKey(m.Groups(1).Value).Value = m.Groups(2).Value
                    ElseIf tempsection IsNot Nothing Then
                        '  Handle Key without value
                        Trace.WriteLine(String.Format("Adding Key [{0}]", line))
                        tempsection.AddKey(line)
                    Else
                        '  This should not occur unless the tempsection is not created yet...
                        Trace.WriteLine(String.Format("Skipping unknown type of data: {0}", line))
                    End If
                End If
            End While
            oReader.Close()
        End Sub

        ' Used to save the data back to the file or your choice
        Public Sub Save(ByVal sFileName As String)
            Dim oWriter As New StreamWriter(sFileName, False)
            For Each s As IniSection In Sections
                Trace.WriteLine(String.Format("Writing Section: [{0}]", s.Name))
                oWriter.WriteLine(String.Format("[{0}]", s.Name))
                For Each k As IniSection.IniKey In s.Keys
                    If k.Value <> String.Empty Then
                        Trace.WriteLine(String.Format("Writing Key: {0}={1}", k.Name, k.Value))
                        oWriter.WriteLine(String.Format("{0}={1}", k.Name, k.Value))
                    Else
                        Trace.WriteLine(String.Format("Writing Key: {0}", k.Name))
                        oWriter.WriteLine(String.Format("{0}", k.Name))
                    End If
                Next
            Next
            oWriter.Close()
        End Sub

        ' Gets all the sections
        Public ReadOnly Property Sections() As System.Collections.ICollection
            Get
                Return m_sections.Values
            End Get
        End Property

        ' Adds a section to the IniFile object, returns a IniSection object to the new or existing object
        Public Function AddSection(ByVal sSection As String) As IniSection
            Dim s As IniSection = Nothing
            sSection = sSection.Trim()
            ' Trim spaces
            If m_sections.ContainsKey(sSection) Then
                s = DirectCast(m_sections(sSection), IniSection)
            Else
                s = New IniSection(Me, sSection)
                m_sections(sSection) = s
            End If
            Return s
        End Function

        ' Removes a section by its name sSection, returns trus on success
        Public Function RemoveSection(ByVal sSection As String) As Boolean
            sSection = sSection.Trim()
            Return RemoveSection(GetSection(sSection))
        End Function

        ' Removes section by object, returns trus on success
        Public Function RemoveSection(ByVal Section As IniSection) As Boolean
            If Section IsNot Nothing Then
                Try
                    m_sections.Remove(Section.Name)
                    Return True
                Catch ex As Exception
                    Trace.WriteLine(ex.Message)
                End Try
            End If
            Return False
        End Function

        '  Removes all existing sections, returns trus on success
        Public Function RemoveAllSections() As Boolean
            m_sections.Clear()
            Return (m_sections.Count = 0)
        End Function

        ' Returns an IniSection to the section by name, NULL if it was not found
        Public Function GetSection(ByVal sSection As String) As IniSection
            sSection = sSection.Trim()
            ' Trim spaces
            If m_sections.ContainsKey(sSection) Then
                Return DirectCast(m_sections(sSection), IniSection)
            End If
            Return Nothing
        End Function

        '  Returns a KeyValue in a certain section
        Public Function GetKeyValue(ByVal sSection As String, ByVal sKey As String) As String
            Dim s As IniSection = GetSection(sSection)
            If s IsNot Nothing Then
                Dim k As IniSection.IniKey = s.GetKey(sKey)
                If k IsNot Nothing Then
                    Return k.Value
                End If
            End If
            Return String.Empty
        End Function

        ' Sets a KeyValuePair in a certain section
        Public Function SetKeyValue(ByVal sSection As String, ByVal sKey As String, ByVal sValue As String) As Boolean
            Dim s As IniSection = AddSection(sSection)
            If s IsNot Nothing Then
                Dim k As IniSection.IniKey = s.AddKey(sKey)
                If k IsNot Nothing Then
                    k.Value = sValue
                    Return True
                End If
            End If
            Return False
        End Function

        ' Renames an existing section returns true on success, false if the section didn't exist or there was another section with the same sNewSection
        Public Function RenameSection(ByVal sSection As String, ByVal sNewSection As String) As Boolean
            '  Note string trims are done in lower calls.
            Dim bRval As Boolean = False
            Dim s As IniSection = GetSection(sSection)
            If s IsNot Nothing Then
                bRval = s.SetName(sNewSection)
            End If
            Return bRval
        End Function

        ' Renames an existing key returns true on success, false if the key didn't exist or there was another section with the same sNewKey
        Public Function RenameKey(ByVal sSection As String, ByVal sKey As String, ByVal sNewKey As String) As Boolean
            '  Note string trims are done in lower calls.
            Dim s As IniSection = GetSection(sSection)
            If s IsNot Nothing Then
                Dim k As IniSection.IniKey = s.GetKey(sKey)
                If k IsNot Nothing Then
                    Return k.SetName(sNewKey)
                End If
            End If
            Return False
        End Function

        ' Remove a key by section name and key name
        Public Function RemoveKey(ByVal sSection As String, ByVal sKey As String) As Boolean
            Dim s As IniSection = GetSection(sSection)
            If s IsNot Nothing Then
                Return s.RemoveKey(sKey)
            End If
            Return False
        End Function

        ' IniSection class 
        Public Class IniSection
            '  IniFile IniFile object instance
            Private m_pIniFile As IniFile
            '  Name of the section
            Private m_sSection As String
            '  List of IniKeys in the section
            Private m_keys As Hashtable

            ' Constuctor so objects are internally managed
            Protected Friend Sub New(ByVal parent As IniFile, ByVal sSection As String)
                m_pIniFile = parent
                m_sSection = sSection
                m_keys = New Hashtable(StringComparer.InvariantCultureIgnoreCase)
            End Sub

            ' Returns all the keys in a section
            Public ReadOnly Property Keys() As System.Collections.ICollection
                Get
                    Return m_keys.Values
                End Get
            End Property

            ' Returns the section name
            Public ReadOnly Property Name() As String
                Get
                    Return m_sSection
                End Get
            End Property

            ' Adds a key to the IniSection object, returns a IniKey object to the new or existing object
            Public Function AddKey(ByVal sKey As String) As IniKey
                sKey = sKey.Trim()
                Dim k As IniSection.IniKey = Nothing
                If sKey.Length <> 0 Then
                    If m_keys.ContainsKey(sKey) Then
                        k = DirectCast(m_keys(sKey), IniKey)
                    Else
                        k = New IniSection.IniKey(Me, sKey)
                        m_keys(sKey) = k
                    End If
                End If
                Return k
            End Function

            ' Removes a single key by string
            Public Function RemoveKey(ByVal sKey As String) As Boolean
                Return RemoveKey(GetKey(sKey))
            End Function

            ' Removes a single key by IniKey object
            Public Function RemoveKey(ByVal Key As IniKey) As Boolean
                If Key IsNot Nothing Then
                    Try
                        m_keys.Remove(Key.Name)
                        Return True
                    Catch ex As Exception
                        Trace.WriteLine(ex.Message)
                    End Try
                End If
                Return False
            End Function

            ' Removes all the keys in the section
            Public Function RemoveAllKeys() As Boolean
                m_keys.Clear()
                Return (m_keys.Count = 0)
            End Function

            ' Returns a IniKey object to the key by name, NULL if it was not found
            Public Function GetKey(ByVal sKey As String) As IniKey
                sKey = sKey.Trim()
                If m_keys.ContainsKey(sKey) Then
                    Return DirectCast(m_keys(sKey), IniKey)
                End If
                Return Nothing
            End Function

            ' Sets the section name, returns true on success, fails if the section
            ' name sSection already exists
            Public Function SetName(ByVal sSection As String) As Boolean
                sSection = sSection.Trim()
                If sSection.Length <> 0 Then
                    ' Get existing section if it even exists...
                    Dim s As IniSection = m_pIniFile.GetSection(sSection)
                    If s IsNot Me AndAlso s IsNot Nothing Then
                        Return False
                    End If
                    Try
                        ' Remove the current section
                        m_pIniFile.m_sections.Remove(m_sSection)
                        ' Set the new section name to this object
                        m_pIniFile.m_sections(sSection) = Me
                        ' Set the new section name
                        m_sSection = sSection
                        Return True
                    Catch ex As Exception
                        Trace.WriteLine(ex.Message)
                    End Try
                End If
                Return False
            End Function

            ' Returns the section name
            Public Function GetName() As String
                Return m_sSection
            End Function

            ' IniKey class
            Public Class IniKey
                '  Name of the Key
                Private m_sKey As String
                '  Value associated
                Private m_sValue As String
                '  Pointer to the parent CIniSection
                Private m_section As IniSection

                ' Constuctor so objects are internally managed
                Protected Friend Sub New(ByVal parent As IniSection, ByVal sKey As String)
                    m_section = parent
                    m_sKey = sKey
                End Sub

                ' Returns the name of the Key
                Public ReadOnly Property Name() As String
                    Get
                        Return m_sKey
                    End Get
                End Property

                ' Sets or Gets the value of the key
                Public Property Value() As String
                    Get
                        Return m_sValue
                    End Get
                    Set(ByVal value As String)
                        m_sValue = value
                    End Set
                End Property

                ' Sets the value of the key
                Public Sub SetValue(ByVal sValue As String)
                    m_sValue = sValue
                End Sub
                ' Returns the value of the Key
                Public Function GetValue() As String
                    Return m_sValue
                End Function

                ' Sets the key name
                ' Returns true on success, fails if the section name sKey already exists
                Public Function SetName(ByVal sKey As String) As Boolean
                    sKey = sKey.Trim()
                    If sKey.Length <> 0 Then
                        Dim k As IniKey = m_section.GetKey(sKey)
                        If k IsNot Me AndAlso k IsNot Nothing Then
                            Return False
                        End If
                        Try
                            ' Remove the current key
                            m_section.m_keys.Remove(m_sKey)
                            ' Set the new key name to this object
                            m_section.m_keys(sKey) = Me
                            ' Set the new key name
                            m_sKey = sKey
                            Return True
                        Catch ex As Exception
                            Trace.WriteLine(ex.Message)
                        End Try
                    End If
                    Return False
                End Function

                ' Returns the name of the Key
                Public Function GetName() As String
                    Return m_sKey
                End Function
            End Class
            ' End of IniKey class
        End Class
        ' End of IniSection class
    End Class

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        If BaseCatalog.Text.Trim = "" Then
            MsgBox("Не указан каталог, содержащий log-файлы технологического журнала!")
        ElseIf ConnectionString.Text.Trim = "" Then
            MsgBox("Не указана строка соединения с MS SQL-базой, в которую будут помещены события тех.журнала!")
        Else
            Dim List = My.Computer.FileSystem.GetFiles(BaseCatalog.Text, FileIO.SearchOption.SearchAllSubDirectories, "*.log")

            Dim a As Integer = List.Count
            For i As Integer = 0 To a - 1
                ListFiles.Items.Add(List(i))
            Next
        End If

    End Sub

    Private Sub StartThread(i As Integer, AlreadyExists As Boolean)
        Dim File As String
        Dim f As TJReader.TJ.LoaderTj

        If ListFiles.Items.Count > 0 Then
            File = ListFiles.Items(0).ToString
            If Not AlreadyExists Then
                f = New TJReader.TJ.LoaderTj
                f.UploadToMSSQL = True
                f.Number = i
                f.MaxCountRowFromEvent = MaxCountRowFromEvent
                f.ConnectionString = ConnectionString.Text
                fibArray(i) = f
            Else
                f = fibArray(i)
            End If
            f.FileName = File
            Dim Thread As New Threading.Thread(New ParameterizedThreadStart(AddressOf f.ThreadPoolCallBack))
            Thread.Start()
            ListFiles.Items.Remove(ListFiles.Items(0))

        End If

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

        Dim a As Integer = 0
        For i As Integer = 0 To ParallelsThreads - 1
            If fibArray(i).FileName = "Done" Then

                If ListFiles.Items.Count > 0 Then
                    FinishedThreads += 1
                    StartThread(i, True)
                Else
                    a += 1
                End If
            End If
        Next
        'Выполненное число потоков считается неправильно на последних этапах, когда новые уже не запускаются, а старые еще продолжают завершаться..
        Label1.Text = "Выполнено потоков всего = " + FinishedThreads.ToString + vbNewLine + _
            "Активно потоков = " + (ParallelsThreads - a).ToString

        Try
            ProgressBar.Value = FinishedThreads
        Catch ex As Exception

        End Try

        If a = ParallelsThreads Then
            ProgressBar.Visible = False
            Timer1.Enabled = False
            Label1.Text = "Обработка завершена"
        End If


    End Sub

    Private Function RestoreIniValue(SectionText As String, KeyText As String) As String
        Dim Section As IniFile.IniSection = Param.GetSection(SectionText)
        If Not Section Is Nothing Then
            Dim Key As IniFile.IniSection.IniKey = Section.GetKey(KeyText)
            If Not Key Is Nothing Then
                RestoreIniValue = Key.Value
            Else
                RestoreIniValue = ""
            End If
        Else
            RestoreIniValue = ""
        End If
    End Function

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'TjDataSet.logs' table. You can move, or remove it, as needed.
        'Me.LogsTableAdapter.Fill(Me.TjDataSet.logs)

        BaseCatalog.Text = "E:\temp\"

        ConnectionString.Text = "Data Source=MSSQL1;Server=[SERVERNAME];Database=[DATABASENAME];Integrated Security=true;"
        'ConnectionString.Text = "Data Source=MSSQL1;Server=srv1007;Database=tj;Integrated Security=true;"
        ParallelsExec.Value = ParallelsThreads
        'For Each Item As String In WindowsApplication2.My.Application.CommandLineArgs
        'CommandLineArgs.Text += ";" + Item
        'Next

        'QueryText.Text = "SELECT TOP 10 " + vbNewLine + _
        '    "DateTime, FileName, Moment, Duration, EventName, " + vbNewLine + _
        '    "Process, Level, ProcessName, text, EventNumber, " + vbNewLine + _
        '    "t_clientID, t_applicationName, t_computerName, t_connectID, SessionID, " + vbNewLine + _
        '    "Usr, AppID, dbpid, Sql, TablesList, Prm, ILev, Rows, Context, Func, Trans, " + vbNewLine + _
        '    "RowsAffected, Descr, planSQLText, Exception " + vbNewLine + _
        '    "FROM logs" + vbNewLine + _
        '    "ORDER BY Duration DESC"

        QueryText.Text = "SELECT        SUM(Duration / 100000) AS Duration, COUNT(1) AS Count, TablesList, ProcessName" + vbNewLine + _
            "FROM logs" + vbNewLine + _
            "GROUP BY TablesList, ProcessName" + vbNewLine + _
            "ORDER BY Duration DESC"

        Param = New IniFile()
        If My.Computer.FileSystem.FileExists("setting.ini") Then
            Param.Load("setting.ini")
        End If

        Dim s As String

        s = RestoreIniValue("CurrentValues", "ConnectionString")
        If Not s = "" Then
            ConnectionString.Text = s
        End If

        s = RestoreIniValue("CurrentValues", "BaseCatalog")
        If Not s = "" Then
            BaseCatalog.Text = s
        End If

        s = RestoreIniValue("CurrentValues", "QueryText")
        If Not s = "" Then
            QueryText.Text = System.Text.Encoding.Unicode.GetString(Convert.FromBase64String(s))
        End If

        s = RestoreIniValue("CurrentValues", "ParallelsExec").Trim
        If Not s = "" Then
            ParallelsExec.Value = Convert.ToInt16(s)
        End If

        s = RestoreIniValue("CurrentValues", "CheckBoxClearTableBeforeRun").Trim
        CheckBoxClearTableBeforeRun.Checked = (s = "1")

    End Sub

    Private Sub NumericUpDown1_ValueChanged(sender As Object, e As EventArgs) Handles ParallelsExec.ValueChanged
        ParallelsThreads = ParallelsExec.Value
        ReDim fibArray(ParallelsThreads)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        If ParallelsThreads > ListFiles.Items.Count Then
            ParallelsThreads = ListFiles.Items.Count
        End If

        ProgressBar.Visible = True
        ProgressBar.Maximum = ListFiles.Items.Count

        'http://msdn.microsoft.com/ru-ru/library/3dasc8as.aspx

        Dim StopLoad As Boolean = False

        If CheckBoxClearTableBeforeRun.Checked Then
            Try
                Dim objConn As New SqlConnection(ConnectionString.Text.Trim)
                objConn.Open()
                Dim command As New SqlCommand("TRUNCATE TABLE [dbo].[logs]", objConn)
                command.ExecuteNonQuery()
                command.CommandText = "SELECT TOP 1 * FROM logs"
                command.ExecuteReader()
            Catch ex As Exception
                MsgBox("Ошибка при очистке таблицы: " + ex.Message)
                StopLoad = True
            End Try
        End If

        If Not StopLoad Then

            For i As Integer = 0 To ParallelsThreads - 1
                StartThread(i, False)
            Next
            Timer1.Enabled = True

        End If


    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        If ConnectionString.Text.Trim = "" Then
            MsgBox("Не указана строка соединения с MS SQL-базой, в которую будут помещены события тех.журнала!")
        Else

            Try
                Dim objConn As New SqlConnection(ConnectionString.Text.Trim)
                objConn.Open()
                Dim command As New SqlCommand("	CREATE TABLE [dbo].[logs](		" +
                                              "		[DateTime] [datetime] NULL,	" +
                                              "		[FileName] [char](200) NULL,	" +
                                              "		[Moment] [char](12) NULL,	" +
                                              "		[Duration] [numeric](18, 5) NULL,	" +
                                              "		[EventName] [char](15) NULL,	" +
                                              "		[Process] [char](50) NULL,	" +
                                              "		[Level] [char](3) NULL,	" +
                                              "		[ProcessName] [char](50) NULL,	" +
                                              "		[text] [varchar](max) NULL,	" +
                                              "		[EventNumber] [int] NULL,	" +
                                              "		[t_clientID] [char](10) NULL,	" +
                                              "		[t_applicationName] [char](50) NULL,	" +
                                              "		[t_computerName] [char](50) NULL,	" +
                                              "		[t_connectID] [char](10) NULL,	" +
                                              "		[SessionID] [char](10) NULL,	" +
                                              "		[Usr] [char](100) NULL,	" +
                                              "		[AppID] [char](20) NULL,	" +
                                              "		[dbpid] [char](10) NULL,	" +
                                              "		[Sql] [varchar](max) NULL,	" +
                                              "		[TablesList] [varchar](max) NULL,	" +
                                              "		[Prm] [varchar](max) NULL,	" +
                                              "		[ILev] [char](20) NULL,	" +
                                              "		[Rows] [char](10) NULL,	" +
                                              "		[Context] [varchar](max) NULL,	" +
                                              "		[ContextLastRow] [varchar](max) NULL,	" +
                                              "		[Func] [char](50) NULL,	" +
                                              "		[Trans] [char](1) NULL,	" +
                                              "		[RowsAffected] [char](10) NULL,	" +
                                              "		[Descr] [varchar](max) NULL,	" +
                                              "		[planSQLText] [varchar](max) NULL,	" +
                                              "		[Exception] [char](100) NULL	" +
                                              "	) ON [PRIMARY]", objConn)
                command.ExecuteNonQuery()
                command.CommandText = "CREATE CLUSTERED INDEX cix_Logs ON [dbo].[logs] ([DateTime],[EventName],[ProcessName],[t_clientID],[SessionID])"
                command.ExecuteNonQuery()
                Try
                    'Если версия MSSQL позволяет, то сразу и сожмем данные
                    command.CommandText = "ALTER INDEX cix_Logs ON [dbo].[logs] REBUILD PARTITION = ALL WITH (DATA_COMPRESSION = PAGE)"
                    command.ExecuteNonQuery()
                Catch ex As Exception
                End Try
                command.CommandText = "SELECT TOP 1 * FROM logs"
                command.ExecuteReader()
                MsgBox("Таблица создана успешно!")
            Catch ex As Exception
                MsgBox("Ошибка при создании таблицы: " + ex.Message)
            End Try
        End If


    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        Try
            Dim objConn As New SqlConnection(ConnectionString.Text.Trim)
            objConn.Open()
            Dim command As New SqlCommand("SELECT 1", objConn)
            command.ExecuteReader()
            MsgBox("Подключение выполнено успешно!")
        Catch ex As Exception
            MsgBox("Ошибка при подключении: " + ex.Message)
        End Try

    End Sub

    'Private Sub LongQueryToolStripButton_Click(sender As Object, e As EventArgs)
    '    Try
    '        Me.LogsTableAdapter.LongQuery(Me.TjDataSet.logs)
    '    Catch ex As System.Exception
    '        System.Windows.Forms.MessageBox.Show(ex.Message)
    '    End Try

    'End Sub

    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click

        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value
        Dim i As Integer
        Dim j As Integer

        SaveFileDialog.ShowDialog()

        xlApp = New Excel.ApplicationClass
        xlWorkBook = xlApp.Workbooks.Add(misValue)
        xlWorkSheet = xlWorkBook.Sheets(1)

        For i = 0 To DataGridView.RowCount - 2
            For j = 0 To DataGridView.ColumnCount - 1
                xlWorkSheet.Cells(i + 1, j + 1) = DataGridView(j, i).Value.ToString()
            Next
        Next

        xlWorkSheet.SaveAs(SaveFileDialog.FileName)
        xlWorkBook.Close()
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


    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        FolderBrowserDialog.SelectedPath = BaseCatalog.Text
        FolderBrowserDialog.ShowDialog()
        BaseCatalog.Text = FolderBrowserDialog.SelectedPath
    End Sub

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        Dim cnn As SqlConnection

        cnn = New SqlConnection(ConnectionString.Text)
        cnn.Open()

        Dim dscmd As New SqlDataAdapter(QueryText.Text, cnn)
        Dim ds As New DataSet
        dscmd.Fill(ds)
        DataGridView.DataSource = ds.Tables(0)
        cnn.Close()
    End Sub


    Private Sub АнализПланаЗапросаToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles АнализПланаЗапросаToolStripMenuItem.Click
        If DataGridView.SelectedCells.Count = 1 Then
            Dim s = DataGridView.SelectedCells(0)
            If DataGridView.Columns(s.ColumnIndex).Name = "planSQLText" Then
                Form2.TextPlanSQL = s.Value.ToString
                Form2.Show(Me)
            Else
                MsgBox("Выделите ячейку с планом запроса!")
            End If
        Else
            MsgBox("Выделите ячейку с планом запроса!")
        End If


    End Sub

    Private Sub DataGridView_MouseClick(sender As Object, e As MouseEventArgs) Handles DataGridView.MouseClick
        Dim c As Control = CType(sender, Control)
        If e.Button = Windows.Forms.MouseButtons.Right Then
            ContextMenuStrip1.Show(c, e.Location, ToolStripDropDownDirection.Default)
        End If
    End Sub

    Private Sub Form1_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed

        Param.AddSection("CurrentValues").AddKey("ConnectionString").Value = ConnectionString.Text
        Param.AddSection("CurrentValues").AddKey("BaseCatalog").Value = BaseCatalog.Text
        Param.AddSection("CurrentValues").AddKey("ParallelsExec").Value = ParallelsExec.Value.ToString
        If CheckBoxClearTableBeforeRun.Checked Then
            Param.AddSection("CurrentValues").AddKey("CheckBoxClearTableBeforeRun").Value = 1
        Else
            Param.AddSection("CurrentValues").AddKey("CheckBoxClearTableBeforeRun").Value = 0
        End If

        Param.AddSection("CurrentValues").AddKey("QueryText").Value = Convert.ToBase64String(System.Text.Encoding.Unicode.GetBytes(QueryText.Text))

        Param.Save("setting.ini")

    End Sub

End Class
