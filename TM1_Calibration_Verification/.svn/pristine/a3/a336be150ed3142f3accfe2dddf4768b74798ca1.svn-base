Imports System.Text.RegularExpressions
Imports System.IO
Imports System.IO.File

Public Class CommonLib
    Shared _ErrorMsg As String = ""

    Shared Property ErrorMsg() As String
        Get
            Return _ErrorMsg
        End Get
        Set(value As String)

        End Set
    End Property

    Shared Sub Delay(ByVal delay_time As Int16, Optional ByVal DisplayTimeout As Boolean = False)
        Dim Counts As Integer = delay_time / 0.2
        Dim Count As Integer = 0

        While Count < Counts
            If DisplayTimeout Then
                'Form1.TimeoutLabel.Text = Math.Round((Counts - Count) * 0.2).ToString
            End If
            System.Threading.Thread.Sleep(200)
            Count += 1
            Application.DoEvents()
        End While
    End Sub

    Shared Function ZuluToUTC(ByVal ZuluTime As String, ByRef UTC_DateTime As DateTime) As Boolean
        Dim TimeFields() As String

        Try
            TimeFields = Regex.Split(ZuluTime, "-|T|Z|:")
            UTC_DateTime = New DateTime(CInt(TimeFields(0)), CInt(TimeFields(1)), CInt(TimeFields(2)),
                                                    CInt(TimeFields(3)), CInt(TimeFields(4)), CInt(TimeFields(5)))
        Catch ex As Exception
            Return False
        End Try

        Return True
    End Function

    Shared Function GetLocalTime(ByRef CurrentTime As String, Optional ByVal Fmt As String = "STD") As Boolean
        Dim ServerTime As DateTime = Now
        Dim ServerTimeStr As String

        If Fmt = "ZULU" Then
            ServerTimeStr = ServerTime.ToString
            CurrentTime = Format(ServerTime.ToUniversalTime, "yyyy-MM-dd\THH:mm:ss\Z")
        ElseIf Fmt = "UTC" Then
            CurrentTime = ServerTime.ToUniversalTime.ToString
        Else
            CurrentTime = ServerTime.ToString
        End If

        Return True
    End Function

    Shared Function GetNetworkTime(ByRef CurrentTime As String, Optional ByVal Fmt As String = "STD") As Boolean
        Dim p As New Process
        Dim Line As String
        Dim startTime As String = Now
        Dim ServerTimeStr As DateTime
        Dim GotTime As Boolean = False
        Dim ServerTime As DateTime

        p.StartInfo.UseShellExecute = False
        p.StartInfo.CreateNoWindow = False
        p.StartInfo.RedirectStandardOutput = True
        p.StartInfo.FileName = "net"
        p.StartInfo.Arguments = "time \\d600loaner"

        p.Start()
        startTime = Now
        While (Not p.HasExited) And (Now.Subtract(startTime).TotalSeconds < 30)
            Application.DoEvents()
        End While
        If Not p.HasExited Then
            CurrentTime = "Timeout getting network time"
            p.Close()
            Return False
        End If
        Line = p.StandardOutput.ReadLine

        While (Not Line = Nothing)
            Console.WriteLine(Line)
            If (Line.Contains("Current time")) Then
                ServerTimeStr = Line.Substring(Line.IndexOf(" is") + 4)
                ServerTime = ServerTimeStr
                ServerTime.AddSeconds(Now.Subtract(startTime).TotalSeconds)
                ServerTimeStr = ServerTime.ToString
                'MsgBox(ServerTime.ToString + vbCr + Now.ToString + vbCr + Now.Subtract(startTime).TotalSeconds.ToString)
                If Fmt = "ZULU" Then
                    CurrentTime = Format(ServerTimeStr, "yyyy-MM-dd\THH:mm:ss\Z")
                ElseIf Fmt = "UTC" Then
                    CurrentTime = ServerTimeStr.ToUniversalTime.ToString()
                Else
                    CurrentTime = ServerTimeStr.ToString
                End If
                GotTime = True
            End If
            Line = p.StandardOutput.ReadLine
        End While
        If Not GotTime Then
            CurrentTime = "Problem getting server time"
        End If

        p.Close()

        Return GotTime
    End Function

    Shared Function ExportDataTableToCSV(ByVal DT As DataTable, ByVal Filename As String) As Boolean
        Dim CSVFile As FileStream
        Dim CSVFileWriter As StreamWriter
        Dim column As DataColumn
        Dim FirstCol As Boolean
        Dim row As DataRow

        Try
            CSVFile = New FileStream(Filename, FileMode.Create, FileAccess.Write)
            CSVFileWriter = New StreamWriter(CSVFile)
        Catch ex As Exception
            _ErrorMsg = "Problem opening " + Filename + " for writing"
            _ErrorMsg += ex.ToString
            Return False
        End Try


        FirstCol = True
        For Each column In DT.Columns
            If Not FirstCol Then
                CSVFileWriter.Write(",")
            Else
                FirstCol = False
            End If
            CSVFileWriter.Write(column.ColumnName)
        Next
        CSVFileWriter.Write(vbCr)

        For Each row In DT.Rows
            For i = 0 To DT.Columns.Count - 1
                CSVFileWriter.Write(row.Item(i).ToString)
                If i < DT.Columns.Count - 1 Then
                    CSVFileWriter.Write(",")
                End If
            Next
            CSVFileWriter.Write(vbCr)
        Next

        CSVFileWriter.Close()
        CSVFile.Close()


        Return True
    End Function

    Shared Function CreateDataTable(ByRef DT As DataTable, ByVal ColumnHeaders() As String) As Boolean
        Dim Name As String
        Dim TypeStr As String
        Dim Field As String

        DT = New DataTable()

        For Each Field In ColumnHeaders
            Name = Split(Field, "|")(0)
            TypeStr = Split(Field, "|")(1)
            Select Case TypeStr
                Case "DT"
                    DT.Columns.Add(Name, Type.GetType("System.DateTime"))
                Case "D"
                    DT.Columns.Add(Name, Type.GetType("System.Double"))
                Case "I"
                    DT.Columns.Add(Name, Type.GetType("System.Int32"))
                Case "UI"
                    DT.Columns.Add(Name, Type.GetType("System.UInt32"))
                Case "S"
                    DT.Columns.Add(Name, Type.GetType("System.String"))
            End Select
        Next
        Return True
    End Function

    'Shared Function FindThumbDrive(ByRef DriveLetter) As Boolean
    '    Dim allDrives() As DriveInfo
    '    Dim d As DriveInfo
    '    Dim ThumbDriveCnt As Integer = 0
    '    Dim startTime As DateTime = Now

    '    allDrives = DriveInfo.GetDrives
    '    For Each d In allDrives
    '        If (d.DriveType = DriveType.Removable) Then
    '            ThumbDriveCnt += 1
    '            DriveLetter = d.Name
    '            While Not d.IsReady And Now.Subtract(startTime).TotalSeconds < 20

    '            End While
    '            If Not d.IsReady Then
    '                _ErrorMsg = "Drive is not ready"
    '                Return False
    '            End If
    '        End If
    '    Next

    '    If Not ThumbDriveCnt = 1 Then
    '        _ErrorMsg = "Found " + ThumbDriveCnt.ToString + " thumb drives, expected 1"
    '        Return False
    '    End If

    '    Return True
    'End Function

    Shared Function ReadCSV(ByVal path As String, ByRef DT As DataTable) As Boolean
        Dim sr As StreamReader
        Dim fullFileStr As String
        Dim lines As String()
        Dim sArr As String()
        Dim line As String

        Try
            sr = New StreamReader(path)
            fullFileStr = sr.ReadToEnd()
            sr.Close()
            sr.Dispose()
            lines = fullFileStr.Split(ControlChars.Cr)
        Catch ex As Exception
            _ErrorMsg = ex.ToString
            Return False
        End Try

        DT = New DataTable
        sArr = lines(0).Split(","c)
        For Each s As String In sArr
            'DT.Columns.Add(New DataColumn())
            s = s.Replace("#", "_number")
            s = s.Replace(".", "_")
            Select Case s
                Case "run_number"
                    DT.Columns.Add(s, Type.GetType("System.Int32"))
                Case "Timestamp"
                    DT.Columns.Add(s, Type.GetType("System.DateTime"))
                Case "H2_PPM"
                    DT.Columns.Add(s, Type.GetType("System.Double"))
                Case "TM8_gas_ppm"
                    DT.Columns.Add(s, Type.GetType("System.Double"))
                Case Else
                    DT.Columns.Add(s)
            End Select
        Next
        Dim row As DataRow
        Dim finalLine As String = ""
        'For Each line As String In lines
        For i = 1 To UBound(lines) - 1
            Try
                line = lines(i)
                If line = "" Then Exit For
                row = DT.NewRow()
                finalLine = line.Replace(Convert.ToString(ControlChars.Cr), "")
                row.ItemArray = finalLine.Split(","c)
                DT.Rows.Add(row)
            Catch ex As Exception
                Form1.AppendText("ReadCSV(): Caught " + ex.ToString)
            End Try
        Next
        Return True
    End Function

    Shared Function GetAverageH2Ppms(ByVal TM1_SN As String, ByVal TargetPPM As Integer, ByRef AveTm1Ppm As Integer, ByRef AveTm8Ppm As Integer) As Boolean
        Dim UUT_Dir As String = ReportDir + "FINAL_TEST" + "\" + TM1_SN + "\"
        Dim di As DirectoryInfo
        Dim files() As FileSystemInfo
        Dim comparer As IComparer = New TimestampComparer()
        Dim CsvFilename As String
        Dim DT As DataTable
        Dim last_run As Integer
        Dim start_run As Integer
        Dim WhereClause As String

        If Not Directory.Exists(UUT_Dir) Then
            _ErrorMsg = "Folder '" + UUT_Dir + "' does not exist"
            Return False
        End If

        di = New DirectoryInfo(UUT_Dir)
        files = di.GetFileSystemInfos("VER_" + TargetPPM.ToString + "*")
        Array.Sort(files, comparer)
        If files.Length > 0 Then
            CsvFilename = files(0).FullName
            If Not ReadCSV(CsvFilename, DT) Then
                Return False
            End If
        Else
            _ErrorMsg = "No VER_" + TargetPPM.ToString + " files found in " + UUT_Dir
            Return False
        End If
        Try
            last_run = CInt(DT.Rows(DT.Rows.Count - 1)("run_number"))
        Catch ex As Exception
            _ErrorMsg = "Problem converting getting last run# from " + CsvFilename
            _ErrorMsg += ex.ToString
            Return False
        End Try
        start_run = last_run - 32
        WhereClause = "run_number > " + start_run.ToString
        Try
            AveTm1Ppm = DT.Compute("AVG(H2_PPM)", WhereClause)
        Catch ex As Exception
            MsgBox(WhereClause)
            _ErrorMsg = "Error calculating average H2.PPM for last 32 runs"
            _ErrorMsg += ex.ToString
            Return False
        End Try
        Try
            AveTm8Ppm = DT.Compute("AVG(TM8_gas_ppm)", WhereClause)
        Catch ex As Exception
            _ErrorMsg = "Error calculating average TM8 gas ppm for last 32 runs"
            _ErrorMsg += ex.ToString
            Return False
        End Try

        Return True
    End Function

    Private Class TimestampComparer
        Implements System.Collections.IComparer

        Public Function Compare(ByVal info1 As Object, ByVal info2 As Object) As Integer Implements System.Collections.IComparer.Compare
            Dim FileInfo1 As System.IO.FileInfo = DirectCast(info1, System.IO.FileInfo)
            Dim FileInfo2 As System.IO.FileInfo = DirectCast(info2, System.IO.FileInfo)

            Dim Timestamp1 As Long = Split(FileInfo1.Name, ".")(1)
            Dim Timestamp2 As Long = Split(FileInfo2.Name, ".")(1)

            If Timestamp1 > Timestamp2 Then Return -1
            If Timestamp1 < Timestamp2 Then Return 1

            Return 0
        End Function
    End Class
End Class
