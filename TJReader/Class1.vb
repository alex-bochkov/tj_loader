Imports System.Data.SqlClient
Imports System.IO
Imports System.Net
Imports Newtonsoft.Json
Imports System.Text
Imports System.Text.RegularExpressions


Public Class TJ

    Structure TJRecord
        Dim DateTime As String
        Dim ID As Guid
        Dim FileName
        Dim Moment
        Dim Duration
        Dim EventName
        Dim Level

        Dim Varibles As Dictionary(Of String, Object)

        Dim Process
        Dim ProcessName
        Dim text
        Dim EventNumber
        Dim t_clientID
        Dim t_applicationName
        Dim t_computerName
        Dim t_connectID
        Dim SessionID
        Dim Usr
        Dim AppID
        Dim dbpid
        Dim Sql
        Dim TablesList
        Dim Prm
        Dim ILev
        Dim Rows
        Dim Context
        Dim ContextLastRow
        Dim Func
        Dim Trans
        Dim RowsAffected
        Dim Descr
        Dim planSQLText
        Dim Exception

    End Structure

    Public Class StringBuilder

        Private m_start As String
        Private m_current As String

        Public Piece_LEN As Integer 'если надо делить на большие или меньшие куски
        'то можно изменить это свойство
        'если вся строка относительно небольшая, то его можно уменьшить
        'если очень большая - увеличить

        Sub New()
            Piece_LEN = 10000
        End Sub
        '*********************************************************
        'Назначение:добавляет часть строки к общей строке
        'am v1.0.0_030717_13:46:11
        '*********************************************************
        Public Sub Append(sString As String)
            If Len(m_current) > Piece_LEN Then
                m_start = m_start & m_current
                m_current = ""
            End If
            m_current = m_current & sString
        End Sub
        '*********************************************************
        'Назначение:очищает строку
        'am v1.0.0_030717_13:46:11
        '*********************************************************
        Public Sub Clear()
            m_current = ""
            m_start = ""
        End Sub
        '*********************************************************
        'Назначение:возвращает текущую строку
        'am v1.0.0_030717_13:49:38
        '*********************************************************
        Public Function GetString() As String
            'Dim s As String
            m_start = m_start & m_current
            GetString = m_start
            m_current = ""
        End Function
    End Class

    Public Class LoaderTj
        Public UploadToMSSQL As Boolean = False
        Public UploadToElastic As Boolean = False
        Public FileName As String
        Public MaxCountRowFromEvent As Integer
        Public Number As Integer
        Public ConnectionString As String
        Public URLElastic As String
        Public ObjectName As String
        Public IndexName As String
        Public DeleteElasticIndexBeforeUploads As Boolean = False
        Private _FileName As String

        'Private _doneEvent As ManualResetEvent

        Sub New()
            ' Sub New(ByVal FileName As String, ByVal doneEvent As ManualResetEvent)

            '_doneEvent = doneEvent 
        End Sub

        ' Wrapper method for use with the thread pool.
        Public Sub ThreadPoolCallBack(ByVal threadContext As Object)
            '_FileName = FileName
            FileName = Calculate(FileName)

        End Sub

        Function ReadFullString(ByVal FullString As String, ByVal exportedAddressFile As String, ByVal fileName As String) As TJRecord

            Dim Rec = New TJRecord
            Rec.ID = Guid.NewGuid

            Dim BufferStr As String = FullString
            Dim ParamName As String = ""
            Dim StrValue As String = ""
            Dim EventNum As Integer = 0
            Dim Delimeter As Integer = 0
            Dim ItIs83 As Boolean = False

            '8.3 - 53:30.004003-1 **** 8.2 - 00:00.1686-1

            If FullString.Substring(0, 13) Like "##:##.######-" Then
                ItIs83 = True
            End If

            Dim d As String = New Date(2000 + fileName.Substring(0, 2), fileName.Substring(2, 2), fileName.Substring(4, 2),
                                       fileName.Substring(6, 2), BufferStr.Substring(0, 2), BufferStr.Substring(3, 2)).ToString("yyyy-MM-dd HH:mm:ss")

            '2015-08-18 11:00:15.153'


            Rec.FileName = exportedAddressFile.Trim
            Rec.text = FullString

            '**Output0Buffer.Moment = BufferStr.Substring(0, 10)
            If ItIs83 Then
                Rec.Moment = BufferStr.Substring(0, 12)
                BufferStr = BufferStr.Substring(13)
            Else
                Rec.Moment = BufferStr.Substring(0, 10)
                BufferStr = BufferStr.Substring(11)
            End If

            'add milliseconds into string date
            Rec.DateTime = d + "." + Rec.Moment.Substring(6, 3)

            Delimeter = BufferStr.IndexOf(",")
            If Delimeter > 0 Then
                Rec.Duration = Convert.ToInt32(BufferStr.Substring(0, Delimeter)) / IIf(ItIs83, 100000, 10000)
            End If

            BufferStr = BufferStr.Substring(Delimeter + 1)
            Delimeter = BufferStr.IndexOf(",")
            If Delimeter > 0 Then
                Rec.EventName = BufferStr.Substring(0, Delimeter).Trim
            End If

            BufferStr = BufferStr.Substring(Delimeter + 1)
            Delimeter = BufferStr.IndexOf(",")
            If Delimeter > 0 Then
                Rec.Level = Convert.ToInt32(BufferStr.Substring(0, Delimeter))
            End If

            BufferStr = BufferStr.Substring(Delimeter + 1)
            Delimeter = BufferStr.IndexOf("=")

            BufferStr = BufferStr.Replace("''", "¦")
            BufferStr = BufferStr.Replace("""""", "÷")


            While Delimeter > 0
                ParamName = BufferStr.Substring(0, Delimeter)
                StrValue = ""

                BufferStr = BufferStr.Substring(Delimeter + 1)
                If Not BufferStr = "" Then
                    If BufferStr.Substring(0, 1) = "'" Then

                        BufferStr = BufferStr.Substring(1)
                        Delimeter = BufferStr.IndexOf("'")
                        If Delimeter > 0 Then
                            StrValue = BufferStr.Substring(0, Delimeter).Trim
                            StrValue = StrValue.Replace("¦", "'")
                        End If
                        If BufferStr.Length > Delimeter + 1 Then
                            BufferStr = BufferStr.Substring(Delimeter + 1 + 1)
                        Else
                            BufferStr = ""
                        End If

                    ElseIf BufferStr.Substring(0, 1) = """" Then

                        BufferStr = BufferStr.Substring(1)
                        Delimeter = BufferStr.IndexOf("""")
                        If Delimeter > 0 Then
                            StrValue = BufferStr.Substring(0, Delimeter).Trim
                            StrValue = StrValue.Replace("÷", """""")
                        End If

                        If BufferStr.Length > Delimeter + 1 Then
                            BufferStr = BufferStr.Substring(Delimeter + 1 + 1)
                        Else
                            BufferStr = ""
                        End If
                    Else

                        Delimeter = BufferStr.IndexOf(",")
                        If Delimeter > 0 Then
                            StrValue = BufferStr.Substring(0, Delimeter).Trim
                        ElseIf Delimeter = -1
                            StrValue = BufferStr
                        End If

                        If BufferStr.Length > Delimeter Then
                            BufferStr = BufferStr.Substring(Delimeter + 1)
                        Else
                            BufferStr = ""
                        End If

                    End If
                End If

                Delimeter = BufferStr.IndexOf("=")

                If ParamName = "process" Then
                    '        Output0Buffer.Process = StrValue
                    Rec.Process = StrValue
                ElseIf ParamName = "p:processName" Then
                    '        Output0Buffer.ProcessName = StrValue
                    Rec.ProcessName = StrValue
                ElseIf ParamName = "t:clientID" Then
                    '        Output0Buffer.tclientID = StrValue
                    Rec.t_clientID = StrValue
                ElseIf ParamName = "t:applicationName" Then
                    '        Output0Buffer.tapplicationName = StrValue
                    Rec.t_applicationName = StrValue
                ElseIf ParamName = "t:computerName" Then
                    '        Output0Buffer.tcomputerName = StrValue
                    Rec.t_computerName = StrValue
                ElseIf ParamName = "t:connectID" Then
                    '        Output0Buffer.tconnectID = StrValue
                    Rec.t_connectID = StrValue
                ElseIf ParamName = "SessionID" Then
                    '        Output0Buffer.SessionID = StrValue
                    Rec.SessionID = StrValue
                ElseIf ParamName = "Usr" Then
                    '        Output0Buffer.Usr = StrValue
                    Rec.Usr = StrValue
                ElseIf ParamName = "AppID" Then
                    '        Output0Buffer.AppID = StrValue
                    Rec.AppID = StrValue
                ElseIf ParamName = "dbpid" Then
                    '        Output0Buffer.dbpid = StrValue
                    Rec.dbpid = StrValue
                ElseIf ParamName = "Sql" Or ParamName = "Sdbl" Then
                    '        Output0Buffer.Sql.AddBlobData(System.Text.Encoding.Unicode.GetBytes(StrValue))
                    Rec.Sql = StrValue

                    If Not String.IsNullOrEmpty(StrValue) Then

                        Dim TablesList As String = ""

                        Dim match As Match = Regex.Match(StrValue, "\sFROM\s+([^ ,]+)(?:\s*,\s*([^ ,]+))*\s+")

                        While match.Success
                            Dim S As String = match.Value.Replace("FROM ", "").Trim
                            S = S.Replace("FROM" & vbCrLf, "").Trim
                            'remove subqueries and temp tables
                            If Not S.StartsWith("(") And Not S.StartsWith("#") Then
                                TablesList = TablesList + " | " + S.Replace("dbo.", "")
                            End If
                            match = match.NextMatch()
                        End While

                        match = Regex.Match(StrValue, "\sJOIN\s+([^ ,]+)(?:\s*,\s*([^ ,]+))*\s+")

                        While match.Success
                            Dim S As String = match.Value.Replace("JOIN ", "").Trim
                            S = S.Replace("JOIN" & vbCrLf, "").Trim
                            'remove subqueries and temp tables
                            If Not S.StartsWith("(") And Not S.StartsWith("#") Then
                                TablesList = TablesList + " | " + S.Replace("dbo.", "")
                            End If
                            match = match.NextMatch()
                        End While

                        match = Regex.Match(StrValue, "\sINTO\s+([^ ,]+)(?:\s*,\s*([^ ,]+))*\s+")

                        While match.Success
                            Dim S As String = match.Value.Replace("INTO ", "").Trim
                            S = S.Replace("INTO" & vbCrLf, "").Trim
                            'remove subqueries and temp tables
                            If Not S.StartsWith("(") And Not S.StartsWith("#") Then
                                TablesList = TablesList + " | " + S.Replace("dbo.", "")
                            End If
                            match = match.NextMatch()
                        End While

                        Rec.TablesList = TablesList

                    End If

                    'here we need to calculate all tables list from this query text

                ElseIf ParamName = "Prm" Then
                    '        Output0Buffer.Prm.AddBlobData(System.Text.Encoding.Unicode.GetBytes(StrValue))
                    Rec.Prm = StrValue
                ElseIf ParamName = "ILev" Then
                    '        Output0Buffer.ILev = StrValue
                    Rec.ILev = StrValue
                ElseIf ParamName = "Rows" Then
                    '        Output0Buffer.Rows = StrValue
                    Rec.Rows = StrValue
                ElseIf ParamName = "Context" Then
                    '        Output0Buffer.Context.AddBlobData(System.Text.Encoding.Unicode.GetBytes(StrValue))
                    Rec.Context = StrValue

                    Dim TextLines() As String = StrValue.Split(Environment.NewLine.ToCharArray, System.StringSplitOptions.RemoveEmptyEntries)
                    If TextLines.Length > 0 Then
                        Rec.ContextLastRow = TextLines(TextLines.Length - 1).Trim
                    End If

                ElseIf ParamName = "Trans" Then
                    '        Output0Buffer.Trans = StrValue
                    Rec.Trans = StrValue
                ElseIf ParamName = "Func" Then
                    If StrValue.Length > 50 Then
                        StrValue = StrValue.Substring(0, 50)
                    End If
                    '        Output0Buffer.Func = StrValue
                    Rec.Func = StrValue
                ElseIf ParamName = "RowsAffected" Then
                    'If Not StrValue = "" Then
                    ' Output0Buffer.RowsAffected = Convert.ToInt32(StrValue)
                    'End If
                ElseIf ParamName = "Descr" Then
                    '        Output0Buffer.Descr.AddBlobData(System.Text.Encoding.Unicode.GetBytes(StrValue))
                    Rec.Descr = StrValue
                ElseIf ParamName = "planSQLText" Then
                    '        Output0Buffer.planSQLText.AddBlobData(System.Text.UTF8Encoding.Unicode.GetBytes(StrValue))
                    Rec.planSQLText = StrValue
                ElseIf ParamName = "Exception" Then
                    '        Output0Buffer.Exception = StrValue
                    Rec.Exception = StrValue
                End If
            End While

            Return Rec

        End Function

        Function ReadFullString_Elastic(ByVal FullString As String, ByVal exportedAddressFile As String, ByVal fileName As String) As TJRecord

            Dim Rec = New TJRecord
            Rec.ID = Guid.NewGuid
            Rec.Varibles = New Dictionary(Of String, Object)
            'Rec.Varibles.Add("EventID", Guid.NewGuid)

            Dim BufferStr As String = FullString
            Dim ParamName As String = ""
            Dim StrValue As String = ""
            Dim EventNum As Integer = 0
            Dim Delimeter As Integer = 0
            Dim ItIs83 As Boolean = False

            '8.3 - 53:30.004003-1 **** 8.2 - 00:00.1686-1

            If FullString.Substring(0, 13) Like "##:##.######-" Then
                ItIs83 = True
            End If

            Dim d As DateTime = New Date(2000 + fileName.Substring(0, 2), fileName.Substring(2, 2), fileName.Substring(4, 2), fileName.Substring(6, 2), BufferStr.Substring(0, 2), BufferStr.Substring(3, 2))

            Rec.Varibles.Add("DateTime", d)
            Rec.Varibles.Add("FileName", exportedAddressFile.Trim)
            Rec.Varibles.Add("text", FullString)

            '**Output0Buffer.Moment = BufferStr.Substring(0, 10)
            If ItIs83 Then
                Rec.Varibles.Add("Moment", BufferStr.Substring(0, 12))
                BufferStr = BufferStr.Substring(13)
            Else
                Rec.Varibles.Add("Moment", BufferStr.Substring(0, 10))
                BufferStr = BufferStr.Substring(11)
            End If

            Delimeter = BufferStr.IndexOf(",")
            If Delimeter > 0 Then
                Rec.Varibles.Add("Duration", Convert.ToInt32(BufferStr.Substring(0, Delimeter)) / IIf(ItIs83, 100000, 10000))
            End If

            BufferStr = BufferStr.Substring(Delimeter + 1)
            Delimeter = BufferStr.IndexOf(",")
            If Delimeter > 0 Then
                Rec.Varibles.Add("EventName", BufferStr.Substring(0, Delimeter).Trim)
            End If

            BufferStr = BufferStr.Substring(Delimeter + 1)
            Delimeter = BufferStr.IndexOf(",")
            If Delimeter > 0 Then
                Rec.Varibles.Add("Level", Convert.ToInt32(BufferStr.Substring(0, Delimeter)))
            End If

            BufferStr = BufferStr.Substring(Delimeter + 1)
            Delimeter = BufferStr.IndexOf("=")

            BufferStr = BufferStr.Replace("''", "¦")
            BufferStr = BufferStr.Replace("""""", "÷")


            While Delimeter > 0
                ParamName = BufferStr.Substring(0, Delimeter)
                StrValue = ""

                BufferStr = BufferStr.Substring(Delimeter + 1)
                If Not BufferStr = "" Then
                    If BufferStr.Substring(0, 1) = "'" Then

                        BufferStr = BufferStr.Substring(1)
                        Delimeter = BufferStr.IndexOf("'")

                        If Delimeter > 0 Then
                            StrValue = BufferStr.Substring(0, Delimeter).Trim
                            StrValue = StrValue.Replace("¦", "'")
                        End If
                        If BufferStr.Length > Delimeter + 1 Then
                            BufferStr = BufferStr.Substring(Delimeter + 1 + 1)
                        Else
                            BufferStr = ""
                        End If

                    ElseIf BufferStr.Substring(0, 1) = """" Then

                        BufferStr = BufferStr.Substring(1)
                        Delimeter = BufferStr.IndexOf("""")

                        If Delimeter > 0 Then
                            StrValue = BufferStr.Substring(0, Delimeter).Trim
                            StrValue = StrValue.Replace("÷", """""")
                        End If

                        If BufferStr.Length > Delimeter + 1 Then
                            BufferStr = BufferStr.Substring(Delimeter + 1 + 1)
                        Else
                            BufferStr = ""
                        End If
                    Else

                        Delimeter = BufferStr.IndexOf(",")
                        If Delimeter > 0 Then
                            StrValue = BufferStr.Substring(0, Delimeter).Trim
                        Else
                            StrValue = BufferStr
                        End If

                        If BufferStr.Length > Delimeter Then
                            BufferStr = BufferStr.Substring(Delimeter + 1)
                        Else
                            BufferStr = ""
                        End If

                    End If
                End If

                Delimeter = BufferStr.IndexOf("=")

                If ParamName = "Func" Then
                    If StrValue.Length > 50 Then
                        StrValue = StrValue.Substring(0, 50)
                    End If
                End If

                Rec.Varibles.Add(ParamName.Replace(":", "_"), StrValue)

            End While

            Return Rec

        End Function

        Private Function SendRequest(uri As String, jsonDataBytes As Byte(), contentType As String, method As String) As String

            Dim req As WebRequest = WebRequest.Create(uri)
            req.ContentType = contentType
            req.Method = method
            req.ContentLength = jsonDataBytes.Length


            Dim stream = req.GetRequestStream()
            stream.Write(jsonDataBytes, 0, jsonDataBytes.Length)
            stream.Close()

            Dim response = req.GetResponse().GetResponseStream()

            Dim reader As New StreamReader(response)
            Dim res = reader.ReadToEnd()
            reader.Close()
            response.Close()

            Return res
        End Function

        Public Sub UpdateElasticTypes()

            If DeleteElasticIndexBeforeUploads Then
                Try
                    Dim data1 = Encoding.UTF8.GetBytes("")
                    Dim result_post1 = SendRequest(URLElastic + "/" + IndexName, data1, "application/json", "DELETE")
                Catch ex As Exception
                End Try
            End If

            Try
                Dim data2 = Encoding.UTF8.GetBytes("")
                Dim result_post2 = SendRequest(URLElastic + "/" + IndexName, data2, "application/json", "PUT")
            Catch ex As Exception
            End Try

            Try

                Dim jsonSring = "{""tj"" : {""properties"" : {""Descr"" : {""type"" : ""string"", ""index"" : ""not_analyzed""}, " + _
                    """Context"" : {""type"" : ""string"", ""index"" : ""not_analyzed""}, " + _
                    """FileName"" : {""type"" : ""string"", ""index"" : ""not_analyzed""}, " + _
                    """EventID"" : {""type"" : ""string"", ""index"" : ""not_analyzed""}, " + _
                    """t_computerName"" : {""type"" : ""string"", ""index"" : ""not_analyzed""}, " + _
                    """t_applicationName"" : {""type"" : ""string"", ""index"" : ""not_analyzed""}}}}"

                Dim data = Encoding.UTF8.GetBytes(jsonSring)
                Dim result_post = SendRequest(URLElastic + "/" + IndexName + "/" + ObjectName + "/_mapping?ignore_conflicts=true", data, "application/json", "PUT")

            Catch ex As Exception
            End Try

        End Sub

        Public Function Calculate(_FileName As String) As String

            Dim objConn As New SqlConnection(ConnectionString)
            Dim command As New SqlCommand("INSERT INTO [dbo].[logs] ([DateTime],[Moment],[FileName],[Duration],[EventName]," +
                                          "[Process],[Level],[ProcessName],[text],[EventNumber],[t_clientID]" +
                                          ",[t_applicationName],[t_computerName],[t_connectID],[SessionID],[Usr]" +
                                          ",[AppID],[dbpid],[Sql],[TablesList],[Prm],[ILev],[Rows],[Context],[ContextLastRow],[Func],[Trans]" +
                                          ",[RowsAffected],[Descr],[planSQLText],[Exception]) " +
                                          "VALUES (@DateTime,@Moment,@FileName,@Duration,@EventName,@Process,@Level," +
                                          "@ProcessName,@text,@EventNumber,@t_clientID,@t_applicationName,@t_computerName," +
                                          "@t_connectID,@SessionID,@Usr,@AppID,@dbpid,@Sql,@TablesList,@Prm,@ILev,@Rows,@Context,@ContextLastRow,@Func," +
                                          "@Trans,@RowsAffected,@Descr,@planSQLText,@Exception)", objConn)

            If UploadToMSSQL Then
                objConn.Open()
            ElseIf UploadToElastic Then
                'UpdateElasticTypes()
            Else
                Return "Done"
            End If

            Dim textReader As StreamReader = New StreamReader(_FileName)

            Dim flatFileInfo As New FileInfo(_FileName)
            Dim fileName As String = flatFileInfo.Name

            If Not fileName Like "########*" Then
                fileName = flatFileInfo.LastWriteTime.ToString("yyMMddHH")
            End If

            Dim nextLine As String

            Dim FullString As String = ""
            Dim CountRowsEvent As Integer = 0
            Dim sb As New StringBuilder


            nextLine = textReader.ReadLine
            Do While nextLine IsNot Nothing

                If CountRowsEvent >= MaxCountRowFromEvent Then
                    'Если в событии очень много строк, то загружать их нет смысла, просто обрезаем,
                    'иначе рискуем получить переполнение буфера
                    sb.Clear()
                    CountRowsEvent = 0
                End If

                If nextLine.Length > 10 Then
                    If nextLine.Substring(0, 10) Like "##:##.####" Then
                        FullString = sb.GetString
                        If Not FullString = "" Then
                            Try
                                If UploadToMSSQL Then
                                    Dim Rec = ReadFullString(FullString, _FileName, fileName)
                                    WriteRecordToDB(command, Rec)
                                Else
                                    Dim Rec = ReadFullString_Elastic(FullString, _FileName, fileName)
                                    WriteRecordToDB_Elastic(command, Rec)
                                End If

                                FullString = ""

                            Catch ex As Exception
                                Dim a = ex.Message
                            End Try
                        End If

                        sb.Clear()
                        sb.Append(nextLine)
                        CountRowsEvent = 1
                    Else
                        Try
                            sb.Append(vbNewLine + nextLine)
                            CountRowsEvent += 1
                        Catch ex As Exception

                        End Try
                    End If
                Else
                    sb.Append(vbNewLine + nextLine)
                    CountRowsEvent += 1
                End If
                nextLine = textReader.ReadLine

            Loop

            FullString = sb.GetString

            If Not FullString = "" Then
                Try
                    If UploadToMSSQL Then
                        Dim Rec = ReadFullString(FullString, _FileName, fileName)
                        WriteRecordToDB(command, Rec)
                    Else
                        Dim Rec = ReadFullString_Elastic(FullString, _FileName, fileName)
                        WriteRecordToDB_Elastic(command, Rec)
                    End If
                Catch ex As Exception
                    Dim a = ex.Message
                End Try
            End If

            textReader.Close()

            If UploadToMSSQL Then
                objConn.Close()
            End If

            Return "Done"

        End Function

        Sub WriteRecordToDB(command As SqlCommand, Rec As TJRecord)

            command.Parameters.Clear()
            command.Parameters.Add("@DateTime", SqlDbType.Char, 25, "DateTime").Value = Rec.DateTime

            command.Parameters.Add("@FileName", SqlDbType.Char, 200, "FileName").Value = Rec.FileName
            command.Parameters.Add("@Moment", SqlDbType.Char, 12, "Moment").Value = Rec.Moment
            command.Parameters.Add("@Duration", SqlDbType.Decimal, 0, "Duration").Value = Rec.Duration
            command.Parameters.Add("@EventName", SqlDbType.Char, 15, "EventName").Value = IIf(Rec.EventName Is Nothing, "", Rec.EventName)
            command.Parameters.Add("@Process", SqlDbType.Char, 50, "Process").Value = IIf(Rec.Process Is Nothing, "", Rec.Process)

            command.Parameters.Add("@Level", SqlDbType.Char, 3, "Level").Value = IIf(Rec.Level Is Nothing, "", Rec.Level)
            command.Parameters.Add("@ProcessName", SqlDbType.Char, 50, "ProcessName").Value = IIf(Rec.ProcessName Is Nothing, "", Rec.ProcessName)
            command.Parameters.Add("@text", SqlDbType.VarChar, 0, "text").Value = IIf(Rec.text Is Nothing, "", Rec.text)
            command.Parameters.Add("@EventNumber", SqlDbType.Int, 0, "EventNumber").Value = IIf(Rec.EventNumber Is Nothing, 0, Rec.EventNumber)
            command.Parameters.Add("@t_clientID", SqlDbType.Char, 10, "t_clientID").Value = IIf(Rec.t_clientID Is Nothing, "", Rec.t_clientID)
            command.Parameters.Add("@t_applicationName", SqlDbType.Char, 50, "t_applicationName").Value = IIf(Rec.t_applicationName Is Nothing, "", Rec.t_applicationName)
            command.Parameters.Add("@t_computerName", SqlDbType.Char, 50, "t_computerName").Value = IIf(Rec.t_computerName Is Nothing, "", Rec.t_computerName)
            command.Parameters.Add("@t_connectID", SqlDbType.Char, 10, "t_connectID").Value = IIf(Rec.t_connectID Is Nothing, "", Rec.t_connectID)
            command.Parameters.Add("@SessionID", SqlDbType.Char, 10, "SessionID").Value = IIf(Rec.SessionID Is Nothing, "", Rec.SessionID)
            command.Parameters.Add("@Usr", SqlDbType.Char, 100, "Usr").Value = IIf(Rec.Usr Is Nothing, "", Rec.Usr)
            command.Parameters.Add("@AppID", SqlDbType.Char, 20, "AppID").Value = IIf(Rec.AppID Is Nothing, "", Rec.AppID)
            command.Parameters.Add("@dbpid", SqlDbType.Char, 10, "dbpid").Value = IIf(Rec.dbpid Is Nothing, "", Rec.dbpid)
            command.Parameters.Add("@Sql", SqlDbType.VarChar, 0, "Sql").Value = IIf(Rec.Sql Is Nothing, "", Rec.Sql)
            command.Parameters.Add("@TablesList", SqlDbType.VarChar, 0, "TablesList").Value = IIf(Rec.TablesList Is Nothing, "", Rec.TablesList)
            command.Parameters.Add("@Prm", SqlDbType.VarChar, 0, "Prm").Value = IIf(Rec.Prm Is Nothing, "", Rec.Prm)
            command.Parameters.Add("@ILev", SqlDbType.Char, 20, "ILev").Value = IIf(Rec.ILev Is Nothing, "", Rec.ILev)
            command.Parameters.Add("@Rows", SqlDbType.Char, 10, "Rows").Value = IIf(Rec.Rows Is Nothing, "", Rec.Rows)
            command.Parameters.Add("@Context", SqlDbType.VarChar, 0, "Context").Value = IIf(Rec.Context Is Nothing, "", Rec.Context)
            command.Parameters.Add("@ContextLastRow", SqlDbType.VarChar, 0, "ContextLastRow").Value = IIf(Rec.ContextLastRow Is Nothing, "", Rec.ContextLastRow)
            command.Parameters.Add("@Func", SqlDbType.Char, 50, "Func").Value = IIf(Rec.Func Is Nothing, "", Rec.Func)
            command.Parameters.Add("@Trans", SqlDbType.Char, 1, "Trans").Value = IIf(Rec.Trans Is Nothing, "", Rec.Trans)
            command.Parameters.Add("@RowsAffected", SqlDbType.Char, 10, "RowsAffected").Value = IIf(Rec.RowsAffected Is Nothing, "", Rec.RowsAffected)
            command.Parameters.Add("@Descr", SqlDbType.VarChar, 0, "Descr").Value = IIf(Rec.Descr Is Nothing, "", Rec.Descr)
            command.Parameters.Add("@planSQLText", SqlDbType.VarChar, 0, "planSQLText").Value = IIf(Rec.planSQLText Is Nothing, "", Rec.planSQLText)
            command.Parameters.Add("@Exception", SqlDbType.Char, 100, "Exception").Value = IIf(Rec.Exception Is Nothing, "", Rec.Exception)

            command.ExecuteNonQuery()



        End Sub

        Sub WriteRecordToDB_Elastic(command As SqlCommand, Rec As TJRecord)

            Dim jsonSring As String = JsonConvert.SerializeObject(Rec.Varibles, Formatting.None)

            Dim data = Encoding.UTF8.GetBytes(jsonSring)

            Dim result_post = SendRequest(URLElastic + "/" + IndexName + "/" + ObjectName + "/" + Rec.ID.ToString, data, "application/json", "PUT")

        End Sub




    End Class




End Class
