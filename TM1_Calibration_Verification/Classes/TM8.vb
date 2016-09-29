Imports System.Text.RegularExpressions

Public Class TM8
    Private _ErrorMsg As String = ""
    Property ErrorMsg() As String
        Get
            Return _ErrorMsg
        End Get
        Set(value As String)

        End Set
    End Property

    Private _disruptHours As Integer = 0
    Property DisruptHours() As Integer
        Get
            Return _disruptHours
        End Get
        Set(value As Integer)
            _disruptHours = value
        End Set
    End Property

    Private Function ConnectToTm8(ByRef telnet As Telnet, ByVal allowedAttemps As Integer) As Boolean
        Dim connected As Boolean = False
        Try
            For attemp As Integer = 0 To allowedAttemps
                connected = telnet.Open(Form1.TM8_SN.Text, 23)
                If connected Then
                    Return connected
                Else
                    Threading.Thread.Sleep((attemp * 1000) + 1000)
                End If
            Next
        Catch ex As Exception
            Form1.AppendText("TM8.ConnectToTm8() caught exception: " + ex.Message)
            Return False
        End Try

        Return connected
    End Function

    Private Function TryConnectToTm8(ByRef telnet As Telnet) As Boolean
        Dim connected As Boolean = True
        Dim failStartTime As DateTime

        ' Try up to 5 times to connect to TM8. If fail then wait for the decision from the user
        ' on whether he/she abort or retry
        If Not ConnectToTm8(telnet, 5) Then
            _ErrorMsg = "Could not open telnet to system board at " + Form1.TM8_SN.Text

            failStartTime = Now
            Dim result As DialogResult = MessageBox.Show("Retry connecting to TM8 or Cancel the test?",
                                                               "TM8 Connection Issue",
                                                               MessageBoxButtons.RetryCancel)
            If result = DialogResult.Cancel Then
                Return False
            Else
                If Not ConnectToTm8(telnet, 5) Then
                    _ErrorMsg = "Again, could not open telnet to system board at " + Form1.TM8_SN.Text + "Test aborted!"
                    Return False
                End If
                Dim disruptTimeInMinutes As Double = Now.Subtract(failStartTime).TotalMinutes
                If disruptTimeInMinutes > 12 Then
                    disruptTimeInMinutes += 60
                    _disruptHours = disruptTimeInMinutes / 60
                End If

            End If
        End If

        Return connected
    End Function
    Private Function GetRecordsdata(ByRef RecData As DataTable, Optional ByVal numRecs As Integer = 24) As Boolean
        Dim TM8_Telnet As New Telnet
        Dim Line As String
        Dim Fields() As String
        Dim run As Integer
        Dim TimeFields() As String
        Dim TS As DateTime
        Dim pstZone As TimeZoneInfo = TimeZoneInfo.FindSystemTimeZoneById("Pacific Standard Time")
        Dim row As DataRow
        Dim ppm As Double

        RecData = New DataTable

        RecData.Columns.Add("RecTime", Type.GetType("System.DateTime"))
        RecData.Columns.Add("run#", Type.GetType("System.Int32"))
        RecData.Columns.Add("ppm", Type.GetType("System.Double"))
        RecData.Columns.Add("gas_ppm", Type.GetType("System.Double"))

        Try
            ' Try to connect to the TM8 unit. If failed get out of here
            If Not TryConnectToTm8(TM8_Telnet) Then
                Return False
            End If

            'If Not TM8_Telnet.Open(Form1.TM8_SN.Text, 23) Then
            '    _ErrorMsg = "Could not open telnet to system board at " + Form1.TM8_SN.Text
            '    Return False
            'End If

            If Not TM8_Telnet.Command("rec -L SAMPLE " + numRecs.ToString) Then
                _ErrorMsg = "Problem sending cmd: 'rec -L " + numRecs.ToString + "'"
                Return False
            End If
            For Each Line In TM8_Telnet.CmdResult.Split(Chr(13))
                Line = Line.Trim
                If Regex.IsMatch(Line, "run# = \d+, ts = \d\d\d\d-\d\d-\d\d \d\d:\d\d:\d\d, run_type = SAMPLE, aborted = false") Then
                    Fields = Regex.Split(Line, "run# = (\d+), ts = (\d\d\d\d-\d\d-\d\d \d\d:\d\d:\d\d), run_type = SAMPLE, aborted = false")
                    run = CInt(Fields(1))
                    TimeFields = Regex.Split(Fields(2), "-|\s|:")
                    TS = New DateTime(CInt(TimeFields(0)), CInt(TimeFields(1)), CInt(TimeFields(2)),
                                                        CInt(TimeFields(3)), CInt(TimeFields(4)), CInt(TimeFields(5)))
                    TS = TimeZoneInfo.ConvertTimeFromUtc(TS, pstZone)
                    row = RecData.NewRow()
                    row("RecTime") = TS
                    row("run#") = run
                    RecData.Rows.Add(row)
                End If
            Next

            If Not TM8_Telnet.Command("rec -p SAMPLE " + numRecs.ToString) Then
                _ErrorMsg = "Problem sending cmd 'rec -p SAMPLE " + numRecs.ToString + "'"
                Return False
            End If
            For Each Line In TM8_Telnet.CmdResult.Split(Chr(13))
                Line = Line.Trim
                If Regex.IsMatch(Line, "run# = \d+\s+oil_ppms = \d+\.\d+ \d+\.\d+ \d+\.\d+ \d+\.\d+ \d+\.\d+ \d+\.\d+ \d+\.\d+ \d+\.\d+ \d+\.\d+") Then
                    Fields = Regex.Split(Line, "run# = (\d+)\s+oil_ppms = \d+\.\d+ \d+\.\d+ \d+\.\d+ \d+\.\d+ (\d+\.\d+) \d+\.\d+ \d+\.\d+ \d+\.\d+ \d+\.\d+")
                    run = CInt(Fields(1))
                    ppm = CDbl(Fields(2))
                    For Each row In RecData.Rows
                        If row("run#") = run Then
                            row("ppm") = ppm
                        End If
                    Next
                End If
            Next

            If Not TM8_Telnet.Command("rec -g SAMPLE " + numRecs.ToString) Then
                _ErrorMsg = "Problem sending cmd 'rec -g SAMPLE " + numRecs.ToString + "'"
                Return False
            End If
            For Each Line In TM8_Telnet.CmdResult.Split(Chr(13))
                Line = Line.Trim
                If Regex.IsMatch(Line, "run# = \d+\s+gas_ppms = \d+\.\d+ \d+\.\d+ \d+\.\d+ \d+\.\d+ \d+\.\d+ \d+\.\d+ \d+\.\d+ \d+\.\d+ \d+\.\d+") Then
                    Fields = Regex.Split(Line, "run# = (\d+)\s+gas_ppms = \d+\.\d+ \d+\.\d+ \d+\.\d+ \d+\.\d+ (\d+\.\d+) \d+\.\d+ \d+\.\d+ \d+\.\d+ \d+\.\d+")
                    run = CInt(Fields(1))
                    ppm = CDbl(Fields(2))
                    For Each row In RecData.Rows
                        If row("run#") = run Then
                            row("gas_ppm") = ppm
                            Form1.AppendText("run=" + row("run#").ToString + ", timestamp=" + row("RecTime").ToString)
                        End If
                    Next
                End If
            Next

            TM8_Telnet.CloseTelnet()
        Catch ex As Exception
            Form1.AppendText("TM8.GetRecordsdata() caught exception: " + ex.Message)
            Return False
        End Try

        Return True
    End Function

    Public Function GetRecdata(ByRef RecData As DataTable, Optional ByVal numRecs As Integer = 24) As Boolean
        Dim done As Boolean = False

        For i As Integer = 0 To 3
            If Not GetRecordsdata(RecData, numRecs) Then
                Form1.AppendText("TM8.GetRecordsdata() failed. Try again.")
                System.Threading.Thread.Sleep(200)
            Else
                done = True
                Exit For
            End If
        Next

        Return done

    End Function

    'Public Shared
    Public Function CombineTm8_Tm1_Data(ByRef Tm1RecData As DataTable, ByVal Tm8RecData As DataTable) As Boolean
        Dim col_name As String
        Dim Tm1Row As DataRow
        Dim Tm8Row As DataRow
        Dim ClosestTm8Row As DataRow
        Dim Tm1DateTime As DateTime
        Dim Tm8DateTime As DateTime
        Dim PrevTm8DateTime As DateTime

        Try
            For Each Column In Tm8RecData.Columns
                col_name = "TM8_" + Column.ColumnName
                Column.ColumnName()
                If Not Tm8RecData.Columns.Contains(col_name) Then
                    Tm1RecData.Columns.Add(col_name, Column.DataType)
                End If
            Next

            For Each Tm1Row In Tm1RecData.Rows
                ClosestTm8Row = Tm8RecData.Rows(0)
                PrevTm8DateTime = ClosestTm8Row("RecTime")
                For i = 1 To Tm8RecData.Rows.Count - 1
                    Tm1DateTime = Tm1Row("Timestamp")
                    Tm8DateTime = Tm8RecData.Rows(i)("RecTime")
                    If Math.Abs(Tm1DateTime.Subtract(Tm8DateTime).TotalSeconds) < Math.Abs(Tm1DateTime.Subtract(PrevTm8DateTime).TotalSeconds) Then
                        ClosestTm8Row = Tm8RecData.Rows(i)
                        PrevTm8DateTime = ClosestTm8Row("RecTime")
                    End If
                Next
                'If Math.Abs(Tm1DateTime.Subtract(PrevTm8DateTime).TotalMinutes) < 120 Then
                If Math.Abs(Tm1DateTime.Subtract(PrevTm8DateTime).TotalMinutes) < 180 Then
                    For Each Column In Tm8RecData.Columns
                        col_name = "TM8_" + Column.ColumnName
                        Try
                            Tm1Row(col_name) = ClosestTm8Row(Column.ColumnName)
                        Catch ex As Exception
                            _ErrorMsg = ex.ToString
                        End Try
                    Next
                End If
            Next
        Catch ex As Exception
            Form1.AppendText("TM8.CombineTm8_Tm1_Data() caught exception: " + ex.Message)
            Return False
        End Try

        Return True
    End Function
End Class
