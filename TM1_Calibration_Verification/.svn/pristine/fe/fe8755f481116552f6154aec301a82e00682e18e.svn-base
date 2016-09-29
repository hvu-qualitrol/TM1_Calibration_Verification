Imports System.Text.RegularExpressions
Imports System.IO
Imports System.IO.File
Imports System.Data.OleDb

Public Class TM1
    Private _ErrorMsg As String = ""
    Property ErrorMsg() As String
        Get
            Return _ErrorMsg
        End Get
        Set(value As String)

        End Set
    End Property

    Private _Results As String
    Property Results() As String
        Get
            Return _Results
        End Get
        Set(value As String)

        End Set
    End Property

    Function GetVersionInfo(ByVal UUT As Hashtable, ByRef VersionInfo As Hashtable, Optional ByVal Quiet As Boolean = False) As Boolean
        Dim SF As New SerialFunctions
        Dim Response As String
        Dim Fields() As String
        Dim SerialPort As Object = UUT("SP")
        Dim ExpectedFields() As String = {"serial number", "firmware version", "BSP version", "application version",
                                          "hardware version", "assembly version", "sensor firmware version",
                                          "model ID", "hour meter", "embedded moisture"}
        Dim Success As Boolean = True

        Try
            VersionInfo = New Hashtable
            If Not SF.Cmd(UUT, Response, "ver", 5, GetPromptFirst:=True, Quiet:=Quiet) Then
                _Results = Response
                _ErrorMsg = "failed sending cmd ver" + vbCr
                _ErrorMsg += SF.ErrorMsg
                Return False
            End If
            _Results = Response + System.Environment.NewLine
            For Each Line In Response.Split(Chr(13))
                Line = Regex.Replace(Line, Chr(10), "")
                _Results += Line + System.Environment.NewLine
                If Not Quiet Then Form1.AppendText(Line, True, UUT:=UUT)
                Line = Regex.Replace(Line, "^\s+", "")
                Line = Regex.Replace(Line, "\s+=\s+", "=")
                Line = Regex.Replace(Line, "\s+$", "")
                Fields = Split(Line, "=")
                If Fields.Length = 2 Then
                    VersionInfo.Add(Fields(0), Fields(1))
                End If
            Next

            For Each ExpectedField In ExpectedFields
                If Not VersionInfo.Contains(ExpectedField) Then
                    Success = False
                    _Results += "Did not see '" + ExpectedField + "' in output of ver cmd" + System.Environment.NewLine
                    If Not Quiet Then Form1.AppendText("Did not see '" + ExpectedField + "' in output of ver cmd", UUT:=UUT)
                End If
            Next
        Catch ex As Exception
            Form1.AppendText("TM1.GetVersionInfo() caught exception: " + ex.Message, UUT:=UUT)
            Return False
        End Try

        Return Success

        'If Not VersionInfo.Contains("firmware version") Then
        '    Form1.AppendText("Did not see 'firmware version' in output of ver cmd", True, UUT:=UUT)
        'End If

        'If Not VersionInfo.Contains("hardware version") Then
        '    Form1.AppendText("Did not see 'hardwave version' in output of ver cmd", True, UUT:=UUT)
        'End If

        'If Not VersionInfo.Contains("assembly version") Then
        '    Form1.AppendText("Did not see 'assembly version' in output of ver cmd", True, UUT:=UUT)
        'End If

        'If Not VersionInfo.Contains("sensor firmware version") Then
        '    Form1.AppendText("Did not see 'sensor firmware version' in output of ver cmd", True, UUT:=UUT)
        'End If

        'Return True
    End Function

    Function GetSensors(ByVal UUT As Hashtable, ByRef SensorInfo As Hashtable) As Boolean
        Dim SF As New SerialFunctions
        Dim Response As String
        Dim Line As String
        Dim Fields() As String
        Dim Name, Value As String
        Dim SerialPort As Object = UUT("SP")

        Try
            SensorInfo = New Hashtable

            If Not SF.Cmd(UUT, Response, "sensor", 10) Then
                _ErrorMsg = "Problem sending command 'sensor'"
                _ErrorMsg += SF.ErrorMsg
                Return False
            End If

            For Each Line In Response.Split(Chr(10), Chr(13))
                If Not Line.Contains(":") Then
                    Continue For
                End If
                Line = Regex.Replace(Line, "\s+is\s+", ": ")
                Fields = Regex.Split(Line, ":")
                Try
                    Name = Fields(0).Trim
                    Value = Fields(1).Trim
                    SensorInfo.Add(Name, Value)
                Catch ex As Exception
                    _ErrorMsg = "Problem parsing sensor line" + Line
                    Return False
                End Try
            Next
        Catch ex As Exception
            Form1.AppendText("TM1.GetSensors() caught exception: " + ex.Message, UUT:=UUT)
            Return False
        End Try

        Return True
    End Function

    Function DumpConfig(ByVal UUT As Hashtable, Optional ByVal Quiet As Boolean = False)
        Dim SF As New SerialFunctions
        Dim Response As String
        Dim Line As String

        Dim LogFile As FileStream
        Dim LogFileWriter As StreamWriter
        Dim LogFilePath As String

        LogFilePath = ReportDir + "FINAL_TEST\" + UUT("SN").Text + "\CONFIG." + Format(Date.UtcNow, "yyyyMMddHHmmss") + ".txt"
        Try
            LogFile = New FileStream(LogFilePath, FileMode.Create, FileAccess.Write)
            LogFileWriter = New StreamWriter(LogFile)
        Catch ex As Exception
            _ErrorMsg = "Problem creating logfile " + LogFilePath + vbCr
            _ErrorMsg += ex.ToString
            Return False
        End Try

        Try
            LogFileWriter.Write("ver" + System.Environment.NewLine)
            If Not SF.Cmd(UUT, Response, "ver", 10, Quiet:=Quiet) Then
                _Results = Response
                _ErrorMsg = "Problem sending command 'ver'" + vbCr
                _ErrorMsg += SF.ErrorMsg
                LogFileWriter.Close()
                LogFile.Close()
                Return False
            End If
            _Results = Response
            For Each Line In Response.Split(Chr(13))
                LogFileWriter.Write(Line + System.Environment.NewLine)
            Next

            ' Log sensor info
            LogFileWriter.Write("sensor" + System.Environment.NewLine)
            If Not SF.Cmd(UUT, Response, "sensor", 10, Quiet:=Quiet) Then
                _Results = Response
                _ErrorMsg = "Problem sending command 'sensor'" + vbCr
                _ErrorMsg += SF.ErrorMsg
                LogFileWriter.Close()
                LogFile.Close()
                Return False
            End If
            _Results = Response
            For Each Line In Response.Split(Chr(13))
                LogFileWriter.Write(Line + System.Environment.NewLine)
            Next


            LogFileWriter.Write("date" + System.Environment.NewLine)
            LogFileWriter.Write(System.Environment.NewLine)
            If Not SF.Cmd(UUT, Response, "date", 10, Quiet:=Quiet) Then
                _Results += vbCr + Response
                _ErrorMsg = "Problem sending command 'date'" + vbCr
                _ErrorMsg += SF.ErrorMsg
                LogFileWriter.Close()
                LogFile.Close()
                Return False
            End If
            _Results += vbCr + Response
            LogFileWriter.Write(Response + System.Environment.NewLine)

            LogFileWriter.Write("config list" + System.Environment.NewLine)
            LogFileWriter.Write(System.Environment.NewLine)
            If Not SF.Cmd(UUT, Response, "co list", 10, Quiet:=Quiet) Then
                _Results += vbCr + Response
                _ErrorMsg = "Problem sending command 'date'" + vbCr
                _ErrorMsg += SF.ErrorMsg
                LogFileWriter.Close()
                LogFile.Close()
                Return False
            End If
            _Results += vbCr + Response
            LogFileWriter.Write(Response + System.Environment.NewLine)

            LogFileWriter.Write("config list -C" + System.Environment.NewLine)
            LogFileWriter.Write(System.Environment.NewLine)
            If Not SF.Cmd(UUT, Response, "co list -C", 10, Quiet:=Quiet) Then
                _Results += vbCr + Response
                _ErrorMsg = "Problem sending command 'date'" + vbCr
                _ErrorMsg += SF.ErrorMsg
                LogFileWriter.Close()
                LogFile.Close()
                Return False
            End If
            _Results += vbCr + Response
            LogFileWriter.Write(Response + System.Environment.NewLine)

            LogFileWriter.Close()
            LogFile.Close()
        Catch ex As Exception
            Form1.AppendText("TM1.DumpConfig() caught exception: " + ex.Message, UUT:=UUT)
            If LogFileWriter IsNot Nothing Then LogFileWriter.Close()
            If LogFile IsNot Nothing Then LogFile.Close()
            Return False
        End Try

        Return True
    End Function

    Function DumpRecs(ByVal UUT As Hashtable, Optional ByVal Quiet As Boolean = False)
        Dim SF As New SerialFunctions
        Dim Response As String
        Dim Line As String
        Dim Cmd As String

        Dim LogFile As FileStream
        Dim LogFileWriter As StreamWriter
        Dim LogFilePath As String
        Dim startTime As DateTime
        Dim done As Boolean
        Dim ReadTimeout As Boolean
        Dim LineCnt As Integer
        Dim last_rec As Integer
        Dim first_rec As Integer
        Dim rec As Integer
        Dim header_displayed As Boolean

        LogFilePath = ReportDir + "FINAL_TEST\" + UUT("SN").Text + "\REC." + Format(Date.UtcNow, "yyyyMMddHHmmss") + ".csv"
        Try
            LogFile = New FileStream(LogFilePath, FileMode.Create, FileAccess.Write)
            LogFileWriter = New StreamWriter(LogFile)
        Catch ex As Exception
            _ErrorMsg = "Problem creating logfile " + LogFilePath + vbCr
            _ErrorMsg += ex.ToString
            Return False
        End Try

        Try
            System.Threading.Thread.Sleep(150)
            If Not Quiet Then Form1.DebugLog("sending cmd rec")
            If Not SF.Cmd(UUT, Response, "rec", 10, Quiet:=Quiet) Then
                _Results = Response
                _ErrorMsg = "Problem sending command 'rec'"
                Return False
            End If
            _Results = Response
            For Each Line In Response.Split(Chr(10), Chr(13))
                Line = Line.Trim
                If Regex.IsMatch(Line, "first_rec_run_id = \d+") Then
                    first_rec = (Regex.Split(Line, "first_rec_run_id = (\d+)"))(1)
                End If
                If Regex.IsMatch(Line, "last_rec_run_id = \d+") Then
                    last_rec = (Regex.Split(Line, "last_rec_run_id = (\d+)"))(1)
                End If
            Next

            'For rec=last_rec to first_rec stop 100
            header_displayed = False
            For rec = last_rec To first_rec Step -100
                Cmd = "rec -D -CSV " + rec.ToString + " 100"
                If Not SF.Cmd(UUT, Response, Cmd, 15, Quiet:=Quiet) Then
                    _Results += vbCr + Response
                    _ErrorMsg = "Problem sending command '" + Cmd + "'"
                    _ErrorMsg += SF.ErrorMsg
                    LogFileWriter.Close()
                    LogFile.Close()
                    Return False
                End If
                _Results += vbCr + Response
                For Each Line In Response.Split(Chr(10), Chr(13))
                    If Regex.IsMatch(Line, "^\d+") Then
                        LogFileWriter.Write(Line + System.Environment.NewLine)
                    Else
                        If Not header_displayed Then
                            LogFileWriter.Write(Line + System.Environment.NewLine)
                            header_displayed = True
                        End If
                    End If
                Next
            Next rec


            'LogFileWriter.Write("rec –d –csv" + System.Environment.NewLine)
            'LogFileWriter.Write(System.Environment.NewLine)
            'Cmd = "rec -D -CSV 2206 100 "
            'System.Threading.Thread.Sleep(50)
            'UUT("SP").write(Cmd + Chr(13) + Chr(10))

            'done = False
            'ReadTimeout = False
            'LineCnt = 0
            'startTime = Now
            'UUT("SP").ReadTimeout = 2000
            'While (Not done And Now.Subtract(startTime).TotalSeconds < 180)
            '    Application.DoEvents()
            '    Try
            '        Line = UUT("SP").ReadLine()
            '        LogFileWriter.Write(Line + System.Environment.NewLine)
            '        If Line.StartsWith("> ") And LineCnt > 0 Then
            '            done = True
            '        End If
            '        LineCnt += 1
            '    Catch ex As Exception
            '        ReadTimeout = True
            '        Exit While
            '    End Try
            'End While
            'If Not done Then
            '    _ErrorMsg += "Problem sending command 'rec –d –csv'"
            '    Return False
            'End If

            'If Not SF.Cmd(UUT, Response, Cmd, 90) Then
            '    _ErrorMsg = "Problem sending command 'rec –d –csv'"
            '    _ErrorMsg += SF.ErrorMsg
            '    LogFileWriter.Close()
            '    LogFile.Close()
            '    Return False
            'End If
            'LogFileWriter.Write(Response + System.Environment.NewLine)

            LogFileWriter.Close()
            LogFile.Close()
        Catch ex As Exception
            Form1.AppendText("TM1.DumpRecs() caught exception: " + ex.Message, UUT:=UUT)
            If LogFileWriter IsNot Nothing Then LogFileWriter.Close()
            If LogFile IsNot Nothing Then LogFile.Close()
            Return False
        End Try

        Return True
    End Function

    Function DumpEvents(ByVal UUT As Hashtable, Optional ByVal Quiet As Boolean = False)
        Dim SF As New SerialFunctions
        Dim Response As String
        Dim Line As String
        Dim Cmd As String

        Dim LogFile As FileStream
        Dim LogFileWriter As StreamWriter
        Dim LogFilePath As String

        LogFilePath = ReportDir + "FINAL_TEST\" + UUT("SN").Text + "\EVENT." + Format(Date.UtcNow, "yyyyMMddHHmmss") + ".txt"
        Try
            LogFile = New FileStream(LogFilePath, FileMode.Create, FileAccess.Write)
            LogFileWriter = New StreamWriter(LogFile)
        Catch ex As Exception
            _ErrorMsg = "Problem creating logfile " + LogFilePath + vbCr
            _ErrorMsg += ex.ToString
            Return False
        End Try

        Try
            LogFileWriter.Write("event any" + System.Environment.NewLine)
            LogFileWriter.Write(System.Environment.NewLine)
            Cmd = "event any"
            If Not SF.Cmd(UUT, Response, Cmd, 180, Quiet:=Quiet) Then
                _Results = Response
                _ErrorMsg = "Problem sending command 'event any'" + vbCr
                _ErrorMsg += SF.ErrorMsg
                LogFileWriter.Close()
                LogFile.Close()
                Return False
            End If
            _Results = Response
            LogFileWriter.Write(Response + System.Environment.NewLine)

            LogFileWriter.Close()
            LogFile.Close()
        Catch ex As Exception
            Form1.AppendText("TM1.DumpEvents() caught exception: " + ex.Message, UUT:=UUT)
            Return False
        End Try

        Return True
    End Function

    Function SetHwRev2Config(ByVal UUT As Hashtable) As Boolean
        Dim SF As New SerialFunctions
        Dim Response As String = ""

        Try
            SF.Cmd(UUT, Response, "Config -S Set AUX1.NAME Relative_Saturation", 10)
            SF.Cmd(UUT, Response, "Config -S Set AUX2.NAME External_OilTemp", 10)
            If UUT("embedded moisture") = "yes" Then
                SF.Cmd(UUT, Response, "Config -S Set MOISTURE.ENABLE TRUE", 10)
            End If
        Catch ex As Exception
            Form1.AppendText("TM1.SetHwRev2Config caught exception: " + ex.Message, UUT:=UUT)
            Return False
        End Try

        Return True
    End Function

    Function GetConfig(ByVal UUT As Hashtable, ByRef Config As Hashtable) As Boolean
        Dim SF As New SerialFunctions
        Dim Response As String
        Dim Fields() As String
        Dim Name, Value As String
        Try
            Config = New Hashtable
            If Not SF.Cmd(UUT, Response, "config", 5, "> ", True, False, False) Then
                _ErrorMsg = "failed sending cmd config" + vbCr
                _ErrorMsg += SF.ErrorMsg
                Return False
            End If

            For Each Line In Response.Split(Chr(13))
                Line = Regex.Replace(Line, Chr(10), "")
                Fields = Regex.Split(Line, "=")
                Try
                    Name = Fields(0)
                    Value = Fields(1)
                Catch ex As Exception
                    _ErrorMsg = "Problem extracting name/value from " + Line
                    Return False
                End Try
                Name = Regex.Replace(Name, "^\s+", "")
                Name = Regex.Replace(Name, "\s+$", "")
                Value = Regex.Replace(Value, "^\s+", "")
                Value = Regex.Replace(Value, "\s+$", "")
                Config.Add(Name, Value)
            Next
        Catch ex As Exception
            Form1.AppendText("TM1.GetConfig() caught exception: " + ex.Message, UUT:=UUT)
            Return False
        End Try

        Return True
    End Function



    Public Function WaitForTm1PpmInSpec(ByVal ppm As Integer, tm8_gas_ppm_ul As Integer, tm8_gas_ppm_ll As Integer, ByRef AllTm1sInSpec As Boolean, ByVal ver_time As Integer, ByVal diffPpmSpec As Double, ByVal delPpmSpec As Double) As Boolean
        Dim Success As Boolean
        Dim results As ReturnResults
        Dim SF As New SerialFunctions
        'Dim AllTm1sInSpec As Boolean
        Dim n As Integer
        Dim h2scan As New H2SCAN_debug
        Dim Tm8RecData As DataTable
        Dim TM8 As New TM8
        Dim Check_ppm_start_DT As DateTime
        Dim Tm1InSpec As Boolean
        Dim last_row_index As Integer
        Dim start_4_hour_window As Integer
        Dim start_8_hour_window As Integer
        Dim csv_filepath As String
        Dim TimeStamp As String = Format(Date.UtcNow, "yyyyMMddHHmmss")
        Dim TM8_gas_ppm As Double
        Dim AllFailed As Boolean = True
        Dim retryCnt As Integer
        Dim getrec_success As Boolean
        Dim RecsInspecCnt As Integer
        'Dim ver_time As Integer = 20

        If Now.Subtract(TM8_gas_start_in_spec).TotalMinutes > 60 Then
            Check_ppm_start_DT = Now
        Else
            Check_ppm_start_DT = TM8_gas_start_in_spec
        End If

        'Reset passedVerify flags and use them to keep track of the ones have passed the test
        For Each UUT In UUTs
            UUT("passedVerify") = False
        Next

        Success = True
        AllTm1sInSpec = False
        TM8.DisruptHours = 0
        While Not AllTm1sInSpec And Now.Subtract(Check_ppm_start_DT).TotalHours < (ver_time + TM8.DisruptHours)
            'Get TM8 rec Data
            Form1.TimeoutLabel.Text = Math.Round(ver_time - Now.Subtract(Check_ppm_start_DT).TotalHours, 1).ToString + "h"
            If Not TM8.GetRecdata(Tm8RecData) Then
                Form1.AppendText(TM8.ErrorMsg)
                Return False
            End If
            TM8_gas_ppm = Tm8RecData.Rows(0)("gas_ppm")
            Form1.TM8_PPM.Text = TM8_gas_ppm.ToString
            Form1.AppendText("TM8 GAS PPM " + TM8_gas_ppm.ToString)
            If TM8_gas_ppm > tm8_gas_ppm_ul Or TM8_gas_ppm < tm8_gas_ppm_ll Then
                TM8_gas_in_spec = False
            End If
            If Not TM8_gas_in_spec Then
                Form1.AppendText("TM8 gas ppm out of spec, expected between " + tm8_gas_ppm_ll.ToString + " and " + tm8_gas_ppm_ul.ToString)
                Return False
            End If

            n = Now.Subtract(Check_ppm_start_DT).TotalMinutes / 15
            If Not n > 0 Then Continue While
            AllTm1sInSpec = True
            CommonLib.Delay(1)
            For Each UUT In UUTs
                Tm1InSpec = True
                If UUT("SN").Text = "" Or UUT("FAILED") Or UUT("passedVerify") Then Continue For

                Dim Tm1RecFields() As String = Tm1Rev0RecFields
                If UUT("TM1 Hardware Version") <> Products(Product)("hardware version 0") Then
                    Tm1RecFields = Tm1Rev2RecFields
                End If

                Try
                    Form1.AppendText(UUT("SN").Text + " verifying login", UUT:=UUT)
                    ' Allow up to three attempts
                    For attempt As Integer = 1 To 3
                        CommonLib.Delay(1)
                        results = SF.Connect(UUT)
                        If results.PassFail Then
                            Exit For
                        End If
                    Next
                    If Not results.PassFail Then
                        UUT("FAILED") = True
                        UUT("TAB").ImageIndex = StatusColor.FAILED
                        Success = False
                        Continue For
                    End If

                    retryCnt = 0
                    getrec_success = False
                    While (Not getrec_success And retryCnt < 3)
                        getrec_success = True
                        If retryCnt > 0 Then
                            CommonLib.Delay(30)
                        End If
                        retryCnt += 1
                        Form1.AppendText("Checking if " + UUT("SN").Text + " in spec", UUT:=UUT)
                        Form1.DebugLog("Creating data table for " + UUT("SN").Text)
                        If Not CommonLib.CreateDataTable(UUT("DT"), Tm1RecFields) Then
                            Form1.AppendText("Problem creating DT for " + UUT("SN").ToString, UUT:=UUT)
                            Form1.AppendText(CommonLib.ErrorMsg, UUT:=UUT)
                            getrec_success = False
                            Continue While
                        End If

                        Form1.DebugLog("Getting rec data for " + UUT("SN").Text)
                        If Not h2scan.GetRecData(UUT, n, UUT("DT"), True) Then
                            Form1.AppendText(h2scan.ErrorMsg, UUT:=UUT)
                            getrec_success = False
                            Continue While
                        End If

                        Form1.DebugLog("Combining TM8 rec data with " + UUT("SN").Text)
                        If Not TM8.CombineTm8_Tm1_Data(UUT("DT"), Tm8RecData) Then
                            Form1.AppendText(h2scan.ErrorMsg, UUT:=UUT)
                            getrec_success = False
                            Continue While
                        End If

                        Form1.AppendText(UUT("SN").Text + " rec returned " + UUT("DT").Rows.Count.ToString + " rows", UUT:=UUT)
                        If n > 2 And UUT("DT").Rows.Count < 2 Then
                            Form1.AppendText(UUT("SN").Text + " < 2 rows of data", UUT:=UUT)
                            getrec_success = False
                            Continue While
                        End If
                    End While
                    If Not getrec_success Then
                        Form1.AppendText(UUT("SN").Text + ":  Problem getting rec data", UUT:=UUT)
                        UUT("FAILED") = True
                        UUT("TAB").ImageIndex = StatusColor.FAILED
                        Success = False
                        Continue For
                    End If

                    Form1.DebugLog(UUT("SN").Text + " Adding error and TM8_ppm_at_temp columns to DT")
                    UUT("DT").Columns.Add("TM8_ppm_at_temp", Type.GetType("System.Double"))
                    UUT("DT").Columns.Add("PER_ERROR", Type.GetType("System.Double"))
                    Dim tm8PpmAtTemp As Double
                    For Each Row In UUT("DT").Rows
                        Try
                            If Not Row("TM8_ppm") Is Nothing Then
                                'Calculate TM8_ppm_at_temp
                                'Incorrect formula
                                'TM8 Gas PPM*(0.037*EXP((0.0078*( oil temperature))-(0.000000005*(oil temperature)^3)))
                                'tm8PpmAtTemp = Row("TM8_gas_ppm") * (0.037 * Math.Exp((0.0078 * Row("OilTemp")) - (0.000000005 * Row("OilTemp") ^ 3)))
                                'Correct formula
                                'Tm8GasPpm*(0.037*EXP((0.0078*Tm1OilTemp))-(0.000000005*Tm1OilTemp^3))
                                tm8PpmAtTemp = Row("TM8_gas_ppm") * (0.037 * Math.Exp((0.0078 * Row("OilTemp"))) - (0.000000005 * Row("OilTemp") ^ 3))
                                Row("TM8_ppm_at_temp") = tm8PpmAtTemp
                                If ppm = 1000 Then
                                    Row("PER_ERROR") = Math.Abs(tm8PpmAtTemp - Row("H2_OIL.PPM"))
                                Else
                                    Row("PER_ERROR") = Math.Abs(tm8PpmAtTemp - Row("H2_OIL.PPM")) * 100.0 / tm8PpmAtTemp
                                End If
                            End If
                        Catch ex As Exception
                            Form1.DebugLog("Error adding TM8_ppm_at_temp & PER_ERROR")
                            Continue For
                        End Try
                    Next

                    Form1.DebugLog(UUT("SN").Text + " Displaying data in gridview ")
                    UUT("GV").DataSource() = UUT("DT")
                    'UUT("GV").FirstDisplayedCell = UUT("GV").Rows(UUT("GV").Rows.Count - 1).Cells(0)
                    For Each col In UUT("GV").Columns
                        If Not (col.Name = "Timestamp" Or col.Name = "H2_OIL.PPM" Or col.Name = "H2.PPM" Or
                                col.Name = "TM8_ppm" Or col.Name = "TM8_gas_ppm" Or col.Name = "PER_ERROR") Then
                            UUT("GV").Columns(col.Name).visible = False
                        Else
                            UUT("GV").Columns(col.Name).visible = True
                        End If
                    Next
                    last_row_index = UUT("DT").Rows.Count - 1
                    ' Changed from 6 hour window to 8 hour window (04/16/15)
                    start_8_hour_window = last_row_index - 4 * 8
                    RecsInspecCnt = 0
                    If start_8_hour_window < 0 Then
                        AllTm1sInSpec = False
                        Tm1InSpec = False
                        Form1.AppendText(UUT("SN").Text + ":  not enough rec's", UUT:=UUT)
                    Else
                        ' Use double precision for spec checking
                        Dim lastPpm As Double = 0
                        Dim currentPpm As Double = 0
                        Dim delPpm As Double = 0
                        Dim diffPpm As Double = 0
                        Tm1InSpec = True

                        ' Reset the flag to restart the checking cycle
                        If ppm = 1000 Then
                            UUT("VERIFY_1000_FAILED") = False
                        ElseIf ppm = 6000 Then
                            UUT("VERIFY_6000_FAILED") = False
                        Else
                            UUT("VERIFY_10000_FAILED") = False
                        End If

                        For i = start_8_hour_window To last_row_index
                            Dim failFlag As Boolean = False
                            Try
                                ' Calculate and check for the delta between the current and the last measurements
                                currentPpm = UUT("DT").Rows(i)("H2_OIL.PPM")
                                If (lastPpm <> 0) Then
                                    If (ppm = 1000) Then
                                        delPpm = Math.Abs(lastPpm - currentPpm)
                                    Else
                                        delPpm = 100 * Math.Abs((lastPpm - currentPpm) / currentPpm)
                                    End If
                                    If (delPpm > delPpmSpec) Then
                                        Log.WriteLine(String.Format("SS = {0} failed: delPpm = {1}, delPpmSpec {2}", UUT("SN").Text, delPpm, delPpmSpec))
                                        failFlag = True
                                    End If
                                End If
                                lastPpm = currentPpm

                                ' Check for the different between TM8 & TM1 measurements
                                diffPpm = UUT("DT").Rows(i)("PER_ERROR")
                                If (diffPpm > diffPpmSpec) Then
                                    Log.WriteLine(String.Format("SS = {0} failed: diffPpm = {1}, diffPpmSpec {2}", UUT("SN").Text, diffPpm, diffPpmSpec))
                                    failFlag = True
                                End If

                                ' Set fail flag accordingly to the test type
                                If failFlag Then
                                    If ppm = 1000 Then
                                        UUT("VERIFY_1000_FAILED") = True
                                    ElseIf ppm = 6000 Then
                                        UUT("VERIFY_6000_FAILED") = True
                                    Else
                                        UUT("VERIFY_10000_FAILED") = True
                                    End If

                                    AllTm1sInSpec = False
                                    Tm1InSpec = False

                                    UUT("GV").Rows(i).DefaultCellStyle.ForeColor = Color.Red
                                    RecsInspecCnt = 0
                                Else
                                    UUT("GV").Rows(i).DefaultCellStyle.BackColor = Color.LightGreen
                                End If

                            Catch ex As Exception
                                Form1.AppendText("Error checking 'PER_ERROR' for row " + i.ToString)
                                AllTm1sInSpec = False
                                Tm1InSpec = False
                            End Try
                        Next
                    End If
                    If Not Tm1InSpec Then
                        Form1.AppendText(UUT("SN").Text + " not yet in spec for prev 6 hours", UUT:=UUT)
                        Form1.AppendText(UUT("SN").Text + RecsInspecCnt.ToString + " recs in spec", UUT:=UUT)
                    Else
                        Form1.AppendText(UUT("SN").Text + " in spec for prev 6 hours", UUT:=UUT)
                        UUT("passedVerify") = True
                    End If
                    Form1.AppendText(Tm1InSpec.ToString)
                Catch ex As Exception
                    Form1.AppendText(UUT("SN").Text + ": WaitForTm1PpmInSpec() caught" + ex.ToString, UUT:=UUT)
                    UUT("FAILED") = True
                    UUT("TAB").ImageIndex = StatusColor.FAILED
                    Success = False
                    Continue For
                End Try
            Next
            If Not AllTm1sInSpec Then
                CommonLib.Delay(60 * 15)
            End If
        End While

        For Each UUT In UUTs
            If UUT("SN").Text = "" Or UUT("DT") Is Nothing Or UUT("FAILED") Then Continue For

            Try
                csv_filepath = ReportDir + "FINAL_TEST" + "\" + UUT("SN").Text + "\VER_" + ppm.ToString + "." + TimeStamp + ".csv"
                If Not CommonLib.ExportDataTableToCSV(UUT("DT"), csv_filepath) Then
                    Form1.AppendText("Problem creating csv file " + csv_filepath, UUT:=UUT)
                    Success = False
                    UUT("FAILED") = True
                    UUT("TAB").ImageIndex = StatusColor.FAILED
                    Continue For
                End If

                Tm1InSpec = True
                If (ppm = 1000 And UUT("VERIFY_1000_FAILED")) Then
                    Log.WriteLine(String.Format("SN {0} failed TestVer1000", UUT("SN").Text))
                    Tm1InSpec = False
                ElseIf (ppm = 6000 And UUT("VERIFY_6000_FAILED")) Then
                    Log.WriteLine(String.Format("SN {0} failed TestVer6000", UUT("SN").Text))
                    Tm1InSpec = False
                ElseIf (ppm = 10000 And UUT("VERIFY_10000_FAILED")) Then
                    Log.WriteLine(String.Format("SN {0} failed TestVer10000", UUT("SN").Text))
                    Tm1InSpec = False
                End If

                If Not Tm1InSpec Then
                    Form1.AppendText(UUT("SN").Text + " NOT in spec", UUT:=UUT)
                    UUT("FAILED") = True
                    UUT("TAB").ImageIndex = StatusColor.FAILED
                Else
                    Form1.AppendText(UUT("SN").Text + " in spec", UUT:=UUT)
                    AllFailed = False
                End If
            Catch ex As Exception
                Form1.AppendText(UUT("SN").Text + ": WaitForTm1PpmInSpec() caught" + ex.ToString, UUT:=UUT)
                UUT("FAILED") = True
                UUT("TAB").ImageIndex = StatusColor.FAILED
                Success = False
            End Try
        Next

        If AllFailed Then
            Return False
        Else
            Return True
        End If
        'Return Success
    End Function

    Function VerifyDate(ByVal UUT As Hashtable) As Boolean
        Dim ServerTimeUTC_Str As String
        Dim ServerTimeUTC As DateTime
        Dim ZuluTime As String
        Dim UutTimeUTC As DateTime
        'Dim AllowedDiff As Integer = 10
        'Dim AllowedDiff As Integer = 30
        Dim AllowedDiff As Integer = 3

        Try
            If Not CommonLib.GetLocalTime(ServerTimeUTC_Str, "UTC") Then
                Form1.AppendText("Problem getting network time", UUT:=UUT)
                Form1.AppendText(ServerTimeUTC_Str, UUT:=UUT)
                Return False
            End If
            Form1.AppendText("Server Time (UTC) = " + ServerTimeUTC_Str, UUT:=UUT)
            ServerTimeUTC = ServerTimeUTC_Str

            If Not GetUutDate(UUT, ZuluTime) Then
                Form1.AppendText(ZuluTime, UUT:=UUT)
                Return False
            End If
            Form1.AppendText("UUT Time (ZULU) = " + ZuluTime, UUT:=UUT)

            If Not CommonLib.ZuluToUTC(ZuluTime, UutTimeUTC) Then
                Form1.AppendText("Problem converting ZuluTime to DateTime type", UUT:=UUT)
                Return False
            End If
            Form1.AppendText("UUT Time (UTC) = " + UutTimeUTC, UUT:=UUT)

            If Math.Abs(ServerTimeUTC.Subtract(UutTimeUTC).TotalSeconds) > AllowedDiff Then
                Form1.AppendText("Expected diff between server and uut time <= " + AllowedDiff.ToString + " seconds", UUT:=UUT)
                Return False
            End If
        Catch ex As Exception
            Form1.AppendText("TM1.VerifyDate() caught exception: " + ex.Message, UUT:=UUT)
            Return False
        End Try

        Return True
    End Function

    Public Function SetVerifyDate(ByVal UUT As Hashtable) As Boolean
        Dim SF As New SerialFunctions
        Dim ZuluTime As String
        Dim Response As String
        Dim Cmd As String

        Try
            Dim startTime As DateTime = Now
            If Not CommonLib.GetLocalTime(ZuluTime, "ZULU") Then
                Form1.AppendText("Problem getting network time", True)
                Form1.AppendText(ZuluTime, True)
                Return False
            End If

            Cmd = "date -s " + ZuluTime
            If Not SF.Cmd(UUT, Response, Cmd, 5) Then
                Form1.AppendText("failed sending cmd " + Cmd, True)
                Return False
            End If

            CommonLib.Delay(10)
            If Not VerifyDate(UUT) Then
                Return False
            End If
        Catch ex As Exception
            Form1.AppendText("TM1.SetVerifyDate() caught exception: " + ex.Message, UUT:=UUT)
            Return False
        End Try

        Return True

    End Function

    Function GetUutDate(ByVal UUT As Hashtable, ByRef ZuluTime As String) As Boolean
        Dim SF As New SerialFunctions
        Dim Response As String

        Try
            ZuluTime = "NOTFOUND"
            If Not SF.Cmd(UUT, Response, "date", 5) Then
                ZuluTime = "failed sending cmd date"
                Return False
            End If
            For Each Line In Response.Split(Chr(13))
                Line = Regex.Replace(Line, Chr(10), "")
                If Line.Contains("Z") Then
                    ZuluTime = Line
                End If
            Next
            If ZuluTime = "NOTFOUND" Then
                Return False
            End If
        Catch ex As Exception
            Form1.AppendText("TM1.GetUutDate() caught exception: " + ex.Message, UUT:=UUT)
            Return False
        End Try

        Return True
    End Function

    ' **************************************************************************************************************
    ' New code block starts
    ' **************************************************************************************************************
    Public Function TestNewCode() As Boolean
        Dim status As Boolean = False
        Dim atRow As Integer = 0
        Dim dt As New DataTable

        ReadCvsIntoTable(dt)
        atRow = DoneRefCycles(dt, 3)
        If (atRow > 0) Then status = True

        Return status
    End Function

    ''' <summary>
    ''' This is for debug only. It reads data from a CSV file into the specified data table
    ''' </summary>
    ''' <param name="dt"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function ReadCvsIntoTable(ByRef dt As DataTable) As Boolean

        Dim folder = "C:\Operations\Production\TM1\FINAL_TEST\TM1020115130133\"
        Dim con = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & folder & ";Extended Properties=""text;HDR=No;FMT=Delimited"";"
        'Dim dt As New DataTable
        Try
            Using Adp As New OleDbDataAdapter("Select * From [REC.csv]", con)
                Adp.Fill(dt)
            End Using
        Catch ex As Exception
            MessageBox.Show("TM1.ReadCvsIntoTable() caught " + ex.Message)
            Return False
        End Try

        Return True

    End Function

    Private Function DoneRefCycles(ByRef dt As DataTable, ByVal numRefCycles As Integer) As Integer
        Dim atRow As Integer = 0
        Dim refCyclesCount As Integer = 0

        For i As Integer = 0 To dt.Rows.Count - 2
            '"F20" = "Status" for improrted file and real data, respectively
            If dt.Rows(i)("Status").ToString.EndsWith("3") And
                Not dt.Rows(i + 1)("Status").ToString.EndsWith("3") Then
                refCyclesCount += 1

                If refCyclesCount = numRefCycles Then
                    atRow = i + 1
                    Exit For
                End If
            End If
        Next

        Return atRow
    End Function

    Public Function WaitForTm1PpmInSpecs(ByVal ppm As Integer, tm8_gas_ppm_ul As Integer, tm8_gas_ppm_ll As Integer, ByRef AllTm1sInSpec As Boolean, ByVal ver_time As Integer, ByVal diffPpmSpec As Double, ByVal delPpmSpec As Double) As Boolean
        Dim Success As Boolean
        Dim results As ReturnResults
        Dim SF As New SerialFunctions
        'Dim AllTm1sInSpec As Boolean
        Dim n As Integer
        Dim h2scan As New H2SCAN_debug
        Dim Tm8RecData As DataTable
        Dim TM8 As New TM8
        Dim Check_ppm_start_DT As DateTime
        Dim Tm1InSpec As Boolean
        Dim last_row_index As Integer
        Dim start_6_hour_window As Integer
        Dim csv_filepath As String
        Dim TimeStamp As String = Format(Date.UtcNow, "yyyyMMddHHmmss")
        Dim TM8_gas_ppm As Double
        Dim AllFailed As Boolean = True
        Dim retryCnt As Integer
        Dim getrec_success As Boolean
        Dim RecsInspecCnt As Integer
        'Dim ver_time As Integer = 20

        If Now.Subtract(TM8_gas_start_in_spec).TotalMinutes > 60 Then
            Check_ppm_start_DT = Now
        Else
            Check_ppm_start_DT = TM8_gas_start_in_spec
        End If

        'Reset passedVerify flags and use them to keep track of the ones have passed the test
        'Reset doneRefCyclesAt to keep track of whether the units has done the needed ref cycles
        For Each UUT In UUTs
            UUT("passedVerify") = False
            UUT("doneRefCyclesAt") = 0
        Next


        Success = True
        AllTm1sInSpec = False
        TM8.DisruptHours = 0
        While Not AllTm1sInSpec And Now.Subtract(Check_ppm_start_DT).TotalHours < (ver_time + TM8.DisruptHours)
            'Get TM8 rec Data
            Form1.TimeoutLabel.Text = Math.Round(ver_time - Now.Subtract(Check_ppm_start_DT).TotalHours, 1).ToString + "h"
            If Not TM8.GetRecdata(Tm8RecData) Then
                Form1.AppendText(TM8.ErrorMsg)
                Return False
            End If
            TM8_gas_ppm = Tm8RecData.Rows(0)("gas_ppm")
            Form1.TM8_PPM.Text = TM8_gas_ppm.ToString
            Form1.AppendText("TM8 GAS PPM " + TM8_gas_ppm.ToString)
            If TM8_gas_ppm > tm8_gas_ppm_ul Or TM8_gas_ppm < tm8_gas_ppm_ll Then
                TM8_gas_in_spec = False
            End If
            If Not TM8_gas_in_spec Then
                Form1.AppendText("TM8 gas ppm out of spec, expected between " + tm8_gas_ppm_ll.ToString + " and " + tm8_gas_ppm_ul.ToString)
                Return False
            End If

            n = Now.Subtract(Check_ppm_start_DT).TotalMinutes / 15
            If Not n > 0 Then Continue While
            AllTm1sInSpec = True
            CommonLib.Delay(1)
            For Each UUT In UUTs
                Tm1InSpec = True
                If UUT("SN").Text = "" Or UUT("FAILED") Or UUT("passedVerify") Then Continue For

                Dim Tm1RecFields() As String = Tm1Rev0RecFields
                If UUT("TM1 Hardware Version") <> Products(Product)("hardware version 0") Then
                    Tm1RecFields = Tm1Rev2RecFields
                End If

                Try
                    Form1.AppendText(UUT("SN").Text + " verifying login", UUT:=UUT)
                    CommonLib.Delay(1)
                    results = SF.Connect(UUT)
                    If Not results.PassFail Then
                        UUT("FAILED") = True
                        UUT("TAB").ImageIndex = StatusColor.FAILED
                        Success = False
                        Continue For
                    End If

                    retryCnt = 0
                    getrec_success = False
                    While (Not getrec_success And retryCnt < 3)
                        getrec_success = True
                        If retryCnt > 0 Then
                            CommonLib.Delay(30)
                        End If
                        retryCnt += 1
                        Form1.AppendText("Checking if " + UUT("SN").Text + " in spec", UUT:=UUT)
                        Form1.DebugLog("Creating data table for " + UUT("SN").Text)
                        If Not CommonLib.CreateDataTable(UUT("DT"), Tm1RecFields) Then
                            Form1.AppendText("Problem creating DT for " + UUT("SN").ToString, UUT:=UUT)
                            Form1.AppendText(CommonLib.ErrorMsg, UUT:=UUT)
                            getrec_success = False
                            Continue While
                        End If

                        Form1.DebugLog("Getting rec data for " + UUT("SN").Text)
                        If Not h2scan.GetRecData(UUT, n, UUT("DT"), True) Then
                            Form1.AppendText(h2scan.ErrorMsg, UUT:=UUT)
                            getrec_success = False
                            Continue While
                        End If

                        Form1.DebugLog("Combining TM8 rec data with " + UUT("SN").Text)
                        If Not TM8.CombineTm8_Tm1_Data(UUT("DT"), Tm8RecData) Then
                            Form1.AppendText(h2scan.ErrorMsg, UUT:=UUT)
                            getrec_success = False
                            Continue While
                        End If

                        Form1.AppendText(UUT("SN").Text + " rec returned " + UUT("DT").Rows.Count.ToString + " rows", UUT:=UUT)
                        If n > 2 And UUT("DT").Rows.Count < 2 Then
                            Form1.AppendText(UUT("SN").Text + " < 2 rows of data", UUT:=UUT)
                            getrec_success = False
                            Continue While
                        End If
                    End While
                    If Not getrec_success Then
                        Form1.AppendText(UUT("SN").Text + ":  Problem getting rec data", UUT:=UUT)
                        UUT("FAILED") = True
                        UUT("TAB").ImageIndex = StatusColor.FAILED
                        Success = False
                        Continue For
                    End If

                    Form1.DebugLog(UUT("SN").Text + " Adding error and TM8_ppm_at_temp columns to DT")
                    UUT("DT").Columns.Add("TM8_ppm_at_temp", Type.GetType("System.Double"))
                    UUT("DT").Columns.Add("PER_ERROR", Type.GetType("System.Double"))
                    Dim tm8PpmAtTemp As Double
                    For Each Row In UUT("DT").Rows
                        Try
                            If Not Row("TM8_ppm") Is Nothing Then
                                'Calculate TM8_ppm_at_temp
                                'Incorrect formula
                                'TM8 Gas PPM*(0.037*EXP((0.0078*( oil temperature))-(0.000000005*(oil temperature)^3)))
                                'tm8PpmAtTemp = Row("TM8_gas_ppm") * (0.037 * Math.Exp((0.0078 * Row("OilTemp")) - (0.000000005 * Row("OilTemp") ^ 3)))
                                'Correct formula
                                'Tm8GasPpm*(0.037*EXP((0.0078*Tm1OilTemp))-(0.000000005*Tm1OilTemp^3))
                                tm8PpmAtTemp = Row("TM8_gas_ppm") * (0.037 * Math.Exp((0.0078 * Row("OilTemp"))) - (0.000000005 * Row("OilTemp") ^ 3))
                                Row("TM8_ppm_at_temp") = tm8PpmAtTemp
                                If ppm = 1000 Then
                                    Row("PER_ERROR") = Math.Abs(tm8PpmAtTemp - Row("H2_OIL.PPM"))
                                Else
                                    Row("PER_ERROR") = Math.Abs(tm8PpmAtTemp - Row("H2_OIL.PPM")) * 100.0 / tm8PpmAtTemp
                                End If
                            End If
                        Catch ex As Exception
                            Form1.DebugLog("Error adding TM8_ppm_at_temp & PER_ERROR")
                            Continue For
                        End Try
                    Next

                    Form1.DebugLog(UUT("SN").Text + " Displaying data in gridview ")
                    UUT("GV").DataSource() = UUT("DT")
                    'UUT("GV").FirstDisplayedCell = UUT("GV").Rows(UUT("GV").Rows.Count - 1).Cells(0)
                    For Each col In UUT("GV").Columns
                        If Not (col.Name = "Timestamp" Or col.Name = "H2_OIL.PPM" Or col.Name = "H2.PPM" Or
                                col.Name = "TM8_ppm" Or col.Name = "TM8_gas_ppm" Or col.Name = "PER_ERROR") Then
                            UUT("GV").Columns(col.Name).visible = False
                        Else
                            UUT("GV").Columns(col.Name).visible = True
                        End If
                    Next

                    ' Check whether the units has done three referenced cycles
                    last_row_index = UUT("DT").Rows.Count - 1
                    If (last_row_index >= (10 * 4)) Then
                        UUT("doneRefCyclesAt") = DoneRefCycles(UUT("DT"), 4)
                    End If
                    ' Continue on the next unit for this one has not been done yet
                    If (UUT("doneRefCyclesAt") = 0) Then
                        Form1.AppendText(UUT("SN").Text + ":  has not done 3 ref cycles", UUT:=UUT)
                        Continue For
                    End If

                    start_6_hour_window = last_row_index - (4 * 6) - UUT("doneRefCyclesAt")
                    RecsInspecCnt = 0
                    If start_6_hour_window < 0 Then
                        AllTm1sInSpec = False
                        Tm1InSpec = False
                        Form1.AppendText(UUT("SN").Text + ":  not enough rec's", UUT:=UUT)
                    Else
                        ' Use double precision for spec checking
                        Dim lastPpm As Double = 0
                        Dim currentPpm As Double = 0
                        Dim delPpm As Double = 0
                        Dim diffPpm As Double = 0
                        Tm1InSpec = True

                        ' Reset the flag to restart the checking cycle
                        If ppm = 1000 Then
                            UUT("VERIFY_1000_FAILED") = False
                        ElseIf ppm = 6000 Then
                            UUT("VERIFY_6000_FAILED") = False
                        Else
                            UUT("VERIFY_10000_FAILED") = False
                        End If

                        For i = start_6_hour_window To last_row_index
                            Dim failFlag As Boolean = False
                            Try
                                ' Calculate and check for the delta between the current and the last measurements
                                currentPpm = UUT("DT").Rows(i)("H2_OIL.PPM")
                                If (lastPpm <> 0) Then
                                    If (ppm = 1000) Then
                                        delPpm = Math.Abs(lastPpm - currentPpm)
                                    Else
                                        delPpm = 100 * Math.Abs((lastPpm - currentPpm) / currentPpm)
                                    End If
                                    If (delPpm > delPpmSpec) Then
                                        Log.WriteLine(String.Format("SS = {0} failed: delPpm = {1}, delPpmSpec {2}", UUT("SN").Text, delPpm, delPpmSpec))
                                        failFlag = True
                                    End If
                                End If
                                lastPpm = currentPpm

                                ' Check for the different between TM8 & TM1 measurements
                                diffPpm = UUT("DT").Rows(i)("PER_ERROR")
                                If (diffPpm > diffPpmSpec) Then
                                    Log.WriteLine(String.Format("SS = {0} failed: diffPpm = {1}, diffPpmSpec {2}", UUT("SN").Text, diffPpm, diffPpmSpec))
                                    failFlag = True
                                End If

                                ' Set fail flag accordingly to the test type
                                If failFlag Then
                                    If ppm = 1000 Then
                                        UUT("VERIFY_1000_FAILED") = True
                                    ElseIf ppm = 6000 Then
                                        UUT("VERIFY_6000_FAILED") = True
                                    Else
                                        UUT("VERIFY_10000_FAILED") = True
                                    End If

                                    AllTm1sInSpec = False
                                    Tm1InSpec = False

                                    UUT("GV").Rows(i).DefaultCellStyle.ForeColor = Color.Red
                                    RecsInspecCnt = 0
                                Else
                                    UUT("GV").Rows(i).DefaultCellStyle.BackColor = Color.LightGreen
                                End If

                            Catch ex As Exception
                                Form1.AppendText("Error checking 'PER_ERROR' for row " + i.ToString)
                                AllTm1sInSpec = False
                                Tm1InSpec = False
                            End Try
                        Next
                    End If
                    If Not Tm1InSpec Then
                        Form1.AppendText(UUT("SN").Text + " not yet in spec for prev 6 hours", UUT:=UUT)
                        Form1.AppendText(UUT("SN").Text + RecsInspecCnt.ToString + " recs in spec", UUT:=UUT)
                    Else
                        Form1.AppendText(UUT("SN").Text + " in spec for prev 6 hours", UUT:=UUT)
                        UUT("passedVerify") = True
                    End If
                    Form1.AppendText(Tm1InSpec.ToString)
                Catch ex As Exception
                    Form1.AppendText(UUT("SN").Text + ": WaitForTm1PpmInSpec() caught" + ex.ToString, UUT:=UUT)
                    UUT("FAILED") = True
                    UUT("TAB").ImageIndex = StatusColor.FAILED
                    Success = False
                    Continue For
                End Try
            Next
            If Not AllTm1sInSpec Then
                CommonLib.Delay(60 * 15)
            End If
        End While

        For Each UUT In UUTs
            If UUT("SN").Text = "" Or UUT("DT") Is Nothing Or UUT("FAILED") Then Continue For

            Try
                csv_filepath = ReportDir + "FINAL_TEST" + "\" + UUT("SN").Text + "\VER_" + ppm.ToString + "." + TimeStamp + ".csv"
                If Not CommonLib.ExportDataTableToCSV(UUT("DT"), csv_filepath) Then
                    Form1.AppendText("Problem creating csv file " + csv_filepath, UUT:=UUT)
                    Success = False
                    UUT("FAILED") = True
                    UUT("TAB").ImageIndex = StatusColor.FAILED
                    Continue For
                End If

                Tm1InSpec = True
                If (ppm = 1000 And UUT("VERIFY_1000_FAILED")) Then
                    Log.WriteLine(String.Format("SN {0} failed TestVer1000", UUT("SN").Text))
                    Tm1InSpec = False
                ElseIf (ppm = 6000 And UUT("VERIFY_6000_FAILED")) Then
                    Log.WriteLine(String.Format("SN {0} failed TestVer6000", UUT("SN").Text))
                    Tm1InSpec = False
                ElseIf (ppm = 10000 And UUT("VERIFY_10000_FAILED")) Then
                    Log.WriteLine(String.Format("SN {0} failed TestVer10000", UUT("SN").Text))
                    Tm1InSpec = False
                End If

                If Not Tm1InSpec Then
                    Form1.AppendText(UUT("SN").Text + " NOT in spec", UUT:=UUT)
                    UUT("FAILED") = True
                    UUT("TAB").ImageIndex = StatusColor.FAILED
                Else
                    Form1.AppendText(UUT("SN").Text + " in spec", UUT:=UUT)
                    AllFailed = False
                End If
            Catch ex As Exception
                Form1.AppendText(UUT("SN").Text + ": WaitForTm1PpmInSpec() caught" + ex.ToString, UUT:=UUT)
                UUT("FAILED") = True
                UUT("TAB").ImageIndex = StatusColor.FAILED
                Success = False
            End Try
        Next

        If AllFailed Then
            Return False
        Else
            Return True
        End If
        'Return Success
    End Function

    ' **************************************************************************************************************
    ' New code block ends
    ' **************************************************************************************************************
End Class
'                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                   