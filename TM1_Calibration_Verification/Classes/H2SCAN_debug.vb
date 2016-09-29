Imports System.Text.RegularExpressions
Imports System.IO
Imports System.IO.File
Imports System.Threading

Public Class H2SCAN_debug
    Public Enum H2SCAN_SI_MODE
        CLI = 0
        MODBUS = 1
    End Enum

    Public Enum H2SCAN_OP_MODE
        FIELD = 0
        LAB = 1
    End Enum

    Private _ErrorMsg As String
    Private _Results As String

    Property ErrorMsg() As String
        Get
            Return _ErrorMsg
        End Get
        Set(value As String)

        End Set
    End Property

    Property Results() As String
        Get
            Return _Results
        End Get
        Set(value As String)

        End Set
    End Property

    Public Function Open(ByRef UUT As Hashtable, Optional ByVal Quiet As Boolean = False) As Boolean
        Dim SF As New SerialFunctions
        Dim Response As String
        Dim PromptFound As Boolean = False
        Dim retryCnt As Integer = 0
        Dim SerialPort As Object = UUT("SP")

        System.Threading.Thread.Sleep(200)

        If Not SF.Cmd(UUT, Response, "sensor -d", 50, "Press CTRL-D to break out of this mode", Quiet:=Quiet, GetPromptFirst:=False) Then
            _Results = Response
            _ErrorMsg = "Sent cmd 'sensor -d', did not see response 'Press CTRL-D to break out of this mode.'" + vbCr
            _ErrorMsg += SF.ErrorMsg
            Return False
        End If
        If Not Quiet Then Form1.AppendText(Response)
        _Results = Response

        SerialPort.Write(Chr(13) + Chr(10))
        While (Not PromptFound And retryCnt < 4)
            CommonLib.Delay(1)
            If SF.Cmd(UUT, Response, Chr(27), 30, "H2scan: ", Quiet, True, False, True) Then
                PromptFound = True
            End If
            _Results += vbCr + Response
            If Not Quiet Then Form1.AppendText(Response)
            retryCnt += 1
        End While
        If Not PromptFound Then
            _ErrorMsg = "Didn't get H2 scan prompt 'H2scan: '"
            Return False
        End If
        CommonLib.Delay(1)
        If Not SF.Cmd(UUT, Response, "=serv" + Chr(13) + Chr(10), 25, "H2scan: ", Quiet, True, False, True) Then
            _ErrorMsg = Response + vbCr + "Problem sending H2SCAN command '=serv'"
            Return False
        End If
        Results += vbCr + Response
        If Not Quiet Then Form1.AppendText(Response)

        Return True
    End Function

    Public Function Close(ByRef UUT As Hashtable, Optional ByVal Quiet As Boolean = False) As Boolean
        Dim SF As New SerialFunctions
        Dim Response As String
        Dim SerialPort As Object = UUT("SP")

        If Not SF.Cmd(UUT, Response, Chr(4), 50, "> ", Quiet, True, False, True) Then
            _Results = Response
            _ErrorMsg = "Couldn't exit from sensor debug mode, expected to see '> '" + vbCr
            _ErrorMsg += SF.ErrorMsg
            Return False
        End If
        _Results = Response

        Return True
    End Function

    Public Function GetOperatingMode(ByRef UUT As Hashtable, ByRef op_mode As H2SCAN_OP_MODE) As Boolean
        Dim T As New TM1
        Dim SensorInfo As Hashtable

        If Not T.GetSensors(UUT, SensorInfo) Then
            _ErrorMsg = T.ErrorMsg
            Return False
        End If
        If SensorInfo("Sensor Operating Mode").ToString.StartsWith("Field") Then
            op_mode = H2SCAN_OP_MODE.FIELD
        ElseIf SensorInfo("Sensor Operating Mode").ToString.StartsWith("Lab") Then
            op_mode = H2SCAN_OP_MODE.LAB
        Else
            _ErrorMsg = "Cant decode H2SCAN operating mode from '" + SensorInfo("Sensor Operating Mode") + "'"
            Return False
        End If

        Return True
    End Function

    Public Function SetSensorMode(ByRef UUT As Hashtable, ByVal op_mode As H2SCAN_OP_MODE, Optional ByVal Quiet As Boolean = False) As Boolean

        If Not SetOperatingMode(UUT, op_mode, Quiet) Then Return False
        CommonLib.Delay(10)
        'If Not Open(UUT, Quiet) Then Return False
        'CommonLib.Delay(3)
        'If Not ForceOperatingMode(UUT, op_mode, Quiet) Then Return False
        'CommonLib.Delay(3)
        'If Not Close(UUT, Quiet) Then Return False
        'CommonLib.Delay(10)

        Return True
    End Function

    Public Function SetOperatingMode(ByRef UUT As Hashtable, ByVal op_mode As H2SCAN_OP_MODE, Optional ByVal Quiet As Boolean = False) As Boolean
        Dim status As Boolean = False
        Dim SF As New SerialFunctions
        Dim Response As String = ""
        'Dim SerialPort As Object = UUT("SP")
        Dim Command As String
        Dim currentOpMode As H2SCAN_OP_MODE

        ' Set config sensor.mode
        Dim setConfigOpMode As String = "LAB"
        If op_mode = H2SCAN_OP_MODE.FIELD Then setConfigOpMode = "FIELD"


        ' Get the current config.sensor.mode setting.
        ' If it is configured for setting mode then check the actual setting
        Form1.AppendText("Check the current config.sensor.mode setting")
        Command = "config get sensor.mode"
        SF.Cmd(UUT, Response, Command, 10)
        If Response.ToUpper.Contains(setConfigOpMode) Then
            Form1.AppendText("current config mode = set config mode, so check the current sensor op mode setting")
            GetOperatingMode(UUT, currentOpMode)
            If currentOpMode = op_mode Then
                Form1.AppendText("currentOpMode = setOpMode")
                status = True
            End If
        Else
            Form1.AppendText("Change config.sensor.mode to " + setConfigOpMode + ". Wait 90s for the change takes effect")
            Command = "config -s set sensor.mode " + setConfigOpMode
            SF.Cmd(UUT, Response, Command, 10)
            CommonLib.Delay(70)
            Form1.AppendText("Check the current sensor op mode setting")

            For i As Integer = 1 To 3
                GetOperatingMode(UUT, currentOpMode)
                If currentOpMode = op_mode Then
                    status = True
                    Exit For
                Else
                    status = False
                    If i < 3 Then CommonLib.Delay(30)
                End If
            Next
        End If

        Form1.AppendText("H2SCAN Operating Mode = " + currentOpMode.ToString)

        Return status
    End Function

    Public Function ForceOperatingMode(ByRef UUT As Hashtable, ByRef op_mode As H2SCAN_OP_MODE, Optional ByVal Quiet As Boolean = False) As Boolean
        Dim SF As New SerialFunctions
        Dim Response As String
        Dim current_mode As H2SCAN_OP_MODE
        Dim SerialPort As Object = UUT("SP")

        Try
            If Not SF.Cmd(UUT, Response, "m" + Chr(13) + Chr(10), 20, "Change (Y/N)? ", Quiet, True, False, True) Then
                _Results = Response
                _ErrorMsg = "Problem sending h2scan m cmd, expected to see 'Change (Y/N)?  '" + vbCr
                _ErrorMsg += SF.ErrorMsg
                Return False
            End If
            _Results = Response

            If Response.Contains("enabled") Then
                current_mode = H2SCAN_OP_MODE.LAB
            ElseIf Response.Contains("disabled") Then
                current_mode = H2SCAN_OP_MODE.FIELD
            Else
                _ErrorMsg = "Cannot determine if lab mode is enabled or disabled"
                If Not SF.Cmd(UUT, Response, Chr(4), 20, "> ", Quiet, True, False) Then
                    _Results += vbCr + Response
                    _ErrorMsg += "Couldn't exit from sensor debug mode, expected to see '> '" + vbCr
                    _ErrorMsg += SF.ErrorMsg
                End If
                _Results += System.Environment.NewLine + Response
                Return False
            End If
            If Not Quiet Then Form1.AppendText("current_mode = " + current_mode.ToString)

            ' Regardless of the current mode, force to the specified mode
            If Not Quiet Then Form1.AppendText("changing operating mode")
            If Not SF.Cmd(UUT, Response, "y" + Chr(13) + Chr(10), 20, "(Y/N)? ", Quiet, True, False, True) Then
                _Results += System.Environment.NewLine + Response
                _ErrorMsg += "Problem answering yes to change mode" + vbCr
                _ErrorMsg += SF.ErrorMsg
                Return False
            End If
            _Results += System.Environment.NewLine + Response
            If Not SF.Cmd(UUT, Response, "y" + Chr(13) + Chr(10), 20, "(Y/N)? ", Quiet, True, False, True) Then
                _Results += vbCr + Response
                _ErrorMsg += "Problem answering yes to change mode" + vbCr
                _ErrorMsg += SF.ErrorMsg
                Return False
            End If
            _Results += System.Environment.NewLine + Response
            If Not SF.Cmd(UUT, Response, "n" + Chr(13) + Chr(10), 20, "Y/N)?", Quiet, True, False, True) Then
                _Results += vbCr + Response
                _ErrorMsg += "Problem answering no to change mode" + vbCr
                _ErrorMsg += SF.ErrorMsg
                Return False
            End If
            _Results += System.Environment.NewLine + Response
            If Not SF.Cmd(UUT, Response, "y" + Chr(13) + Chr(10), 20, "H2scan: ", Quiet, True, False, True) Then
                _Results += vbCr + Response
                _ErrorMsg += "Problem answering yes to change mode" + vbCr
                _ErrorMsg += SF.ErrorMsg
                Return False
            End If
            _Results += System.Environment.NewLine + Response
        Catch ex As Exception
            Form1.AppendText(UUT("SN").Text + ": H2SCAN_debug.ForceOperatingMode() caught" + ex.ToString, UUT:=UUT)
            Return False
        End Try

        Return True
    End Function

    Public Function SetOperatingMode0(ByRef UUT As Hashtable, ByRef op_mode As H2SCAN_OP_MODE, Optional ByVal Quiet As Boolean = False) As Boolean
        Dim SF As New SerialFunctions
        Dim Response As String
        Dim current_mode As H2SCAN_OP_MODE
        Dim SerialPort As Object = UUT("SP")

        If Not SF.Cmd(UUT, Response, "m" + Chr(13) + Chr(10), 20, "Change (Y/N)? ", Quiet, True, False, True) Then
            _Results = Response
            _ErrorMsg = "Problem sending h2scan m cmd, expected to see 'Change (Y/N)?  '" + vbCr
            _ErrorMsg += SF.ErrorMsg
            Return False
        End If
        _Results = Response

        '2scan: M
        'LabTest mode is enabled Change (Y/N)? n

        'H2scan: M()
        'LabTest mode is enabled Change (Y/N)? y
        'LabTest mode is disabled Change (Y/N)? y
        'LabTest mode is enabled Change (Y/N)? n
        'Save as Default (Y/N)? y
        '...wait...
        'H2scan:
        'Enabling Modbus on H2 Sensor.  Please wait...
        '> config list sensor.mode
        '        sensor.mode = LAB
        '>

        If Response.Contains("enabled") Then
            current_mode = H2SCAN_OP_MODE.LAB
        ElseIf Response.Contains("disabled") Then
            current_mode = H2SCAN_OP_MODE.FIELD
        Else
            _ErrorMsg = "Cannot determine if lab mode is enabled or disabled"
            If Not SF.Cmd(UUT, Response, Chr(4), 20, "> ", Quiet, True, False) Then
                _Results += vbCr + Response
                _ErrorMsg += "Couldn't exit from sensor debug mode, expected to see '> '" + vbCr
                _ErrorMsg += SF.ErrorMsg
            End If
            _Results += System.Environment.NewLine + Response
            Return False
        End If
        If Not Quiet Then Form1.AppendText("current_mode = " + current_mode.ToString)

        If Not current_mode = op_mode Then
            _Results += System.Environment.NewLine + "changing operating mode"
            If Not Quiet Then Form1.AppendText("changing operating mode")
            If Not SF.Cmd(UUT, Response, "y" + Chr(13) + Chr(10), 20, "(Y/N)? ", Quiet, True, False, True) Then
                _Results += System.Environment.NewLine + Response
                _ErrorMsg += "Problem answering yes to change mode" + vbCr
                _ErrorMsg += SF.ErrorMsg
            End If
            _Results += System.Environment.NewLine + Response
            If Not SF.Cmd(UUT, Response, "n" + Chr(13) + Chr(10), 20, "(Y/N)? ", Quiet, True, False, True) Then
                _Results += vbCr + Response
                _ErrorMsg += "Problem answering no to change mode" + vbCr
                _ErrorMsg += SF.ErrorMsg
            End If
            _Results += System.Environment.NewLine + Response
            If Not SF.Cmd(UUT, Response, "y" + Chr(13) + Chr(10), 20, "H2scan: ", Quiet, True, False, True) Then
                _Results += vbCr + Response
                _ErrorMsg += "Didn't see H2scan prompt" + vbCr
                _ErrorMsg += SF.ErrorMsg
            End If
            _Results += System.Environment.NewLine + Response
        Else
            If Not SF.Cmd(UUT, Response, "n" + Chr(13) + Chr(10), 20, "H2scan: ", Quiet, True, False, True) Then
                _Results += vbCr + Response
                _ErrorMsg += "Didn't see H2scan prompt" + vbCr
                _ErrorMsg += SF.ErrorMsg
            End If
            _Results += System.Environment.NewLine + Response
        End If

        Return True
    End Function

    Public Function GetSerialMode(ByRef UUT As Hashtable, ByRef si_mode As H2SCAN_SI_MODE) As Boolean
        Dim T As New TM1
        Dim SensorInfo As Hashtable

        If Not T.GetSensors(UUT, SensorInfo) Then
            _ErrorMsg = T.ErrorMsg
            Return False
        End If
        If SensorInfo("Sensor Model") = "" Or SensorInfo("Sensor Model") = "??" Then
            si_mode = H2SCAN_SI_MODE.CLI
        Else
            si_mode = H2SCAN_SI_MODE.MODBUS
        End If
        Return True
    End Function

    Public Function SetSerialMode(ByRef UUT As Hashtable, ByVal si_mode As H2SCAN_SI_MODE) As Boolean
        Dim SF As New SerialFunctions
        Dim Response As String
        Dim SerialPort As Object = UUT("SP")

        If Not SF.Cmd(UUT, Response, "si 2" + Chr(13) + Chr(10) + Chr(4), 20, "> ", False, True, False) Then
            _ErrorMsg = "Couldn't exit from sensor debug mode, expected to see '> '" + vbCr
            _ErrorMsg += SF.ErrorMsg
            Return False
        End If
        Return True
    End Function

    Public Function VerifyModbusID(ByRef UUT As Hashtable) As Boolean
        Dim SF As New SerialFunctions
        Dim Response As String
        Dim Line As String
        Dim ModbusID As Integer = 0
        Dim SerialPort As Object = UUT("SP")

        If Not SF.Cmd(UUT, Response, "mi" + Chr(13) + Chr(10), 5, "Y/N)? ", False, True, False) Then
            _ErrorMsg = "Sent mi command, did not see expected '(Y/N)? '"
            _ErrorMsg += SF.ErrorMsg
            Return False
        End If
        Form1.AppendText("Response = " + Response)
        For Each Line In Regex.Split(Response, "\r\n")
            If Regex.IsMatch(Line, "Modbus ID is \d+ Change") Then
                Try
                    ModbusID = CInt(Regex.Split(Line, "Modbus ID is (\d+) Change")(1))
                Catch ex As Exception
                    _ErrorMsg = "Problem extracting modbus ID from line '" + Line + "'"
                    _ErrorMsg += ex.ToString
                    Return False
                End Try
            End If
            Form1.AppendText("Line = " + Line)
        Next

        System.Threading.Thread.Sleep(50)
        If Not SF.Cmd(UUT, Response, "n " + Chr(13) + Chr(10), 5, "H2scan: ", False, True, False) Then
            _ErrorMsg = "Problem sending 'n' in response to 'Y/N)? '"
            _ErrorMsg += SF.ErrorMsg
            Return False
        End If

        Form1.AppendText("Modbus ID = " + ModbusID.ToString)
        If Not ModbusID = 1 Then
            _ErrorMsg = "Expecting MosbusID = 1"
            Return False
        End If


        Return True
    End Function

    Public Function UX(ByRef UUT As Hashtable) As Boolean
        Dim SF As New SerialFunctions
        Dim Response As String
        Dim Line As String
        Dim SerialPort As Object = UUT("SP")

        If Not SF.Cmd(UUT, Response, "ux" + Chr(13) + Chr(10), 5, "Y/N)? ", False, True, False) Then
            _ErrorMsg = "Sent ux command, did not see expected '(Y/N)? '"
            _ErrorMsg += SF.ErrorMsg
            Return False
        End If
        System.Threading.Thread.Sleep(50)
        If Not SF.Cmd(UUT, Response, "y " + Chr(13) + Chr(10), 20, "H2scan: ", False, True, False) Then
            _ErrorMsg = "Problem sending 'y' in response to 'Y/N)? '"
            _ErrorMsg += SF.ErrorMsg
            Return False
        End If

        Return True
    End Function

    Public Function U0(ByRef UUT As Hashtable) As Boolean
        Dim SF As New SerialFunctions
        Dim Response As String
        Dim Line As String
        Dim SerialPort As Object = UUT("SP")
        Dim U0_successful As Boolean = False

        If Not SF.Cmd(UUT, Response, "u0" + Chr(13) + Chr(10), 5, "Y/N)? ", False, True, False) Then
            _ErrorMsg = "Sent u0 command, did not see expected '(Y/N)? '"
            _ErrorMsg += SF.ErrorMsg
            Return False
        End If
        System.Threading.Thread.Sleep(50)
        If Not SF.Cmd(UUT, Response, "y " + Chr(13) + Chr(10), 20, "H2scan: ", False, True, False) Then
            _ErrorMsg = "Problem sending 'y' in response to 'Y/N)? '"
            _ErrorMsg += SF.ErrorMsg
            Return False
        End If
        If Not Response.Contains("Set condition #1 hydrogen and let sensor stabilize for 16 hours") Then
            _ErrorMsg = "Expected:  Set condition #1 hydrogen and let sensor stabilize for 16 hours"
            _ErrorMsg += SF.ErrorMsg
            Return False
        End If

        Return True
    End Function


    Public Function U1(ByRef UUT As Hashtable, ByVal tm8_gas_ppm As Double) As Boolean
        Dim SF As New SerialFunctions
        Dim Response As String
        Dim Line As String
        Dim SerialPort As Object = UUT("SP")
        Dim U0_successful As Boolean = False

        If Not SF.Cmd(UUT, Response, "u1" + Chr(13) + Chr(10), 5, "Y/N)? ", False, True, False) Then
            _ErrorMsg = "Sent u1 command, did not see expected '(Y/N)? '"
            _ErrorMsg += SF.ErrorMsg
            Return False
        End If
        System.Threading.Thread.Sleep(50)
        If Not SF.Cmd(UUT, Response, "y " + Chr(13) + Chr(10), 20, "Enter expected hydrogen (gas phase ppm): ", False, True, False) Then
            _ErrorMsg = "Problem sending 'y' in response to 'Y/N)? '"
            _ErrorMsg += SF.ErrorMsg
            Return False
        End If
        System.Threading.Thread.Sleep(50)
        If Not SF.Cmd(UUT, Response, tm8_gas_ppm.ToString + Chr(13) + Chr(10), 20, "H2scan: ", False, True, False) Then
            _ErrorMsg = "Problem sending 'y' in response to 'Y/N)? '"
            _ErrorMsg += SF.ErrorMsg
            Return False
        End If
        'If Not Response.Contains("Set condition #2 hydrogen and let sensor stabilize for 12 hours") Then
        '    _ErrorMsg = Response
        '    _ErrorMsg += "Expected:  Set condition #2 hydrogen and let sensor stabilize for 12 hours"
        '    _ErrorMsg += SF.ErrorMsg
        '    Return False
        'End If

        Return True
    End Function

    Public Function U2(ByRef UUT As Hashtable, ByVal tm8_gas_ppm As Double) As Boolean
        Dim SF As New SerialFunctions
        Dim Response As String
        Dim Line As String
        Dim SerialPort As Object = UUT("SP")
        Dim U0_successful As Boolean = False
        Dim YYYY As String = Format(Now, "yyyy")
        Dim MM As String = Format(Now, "MM")
        Dim DD As String = Format(Now, "dd")
        Dim CalDate As String = MM + "/" + DD + "/" + YYYY

        If Not SF.Cmd(UUT, Response, "u2" + Chr(13) + Chr(10), 5, "Y/N)? ", False, True, False) Then
            _ErrorMsg = "Sent u2 command, did not see expected '(Y/N)? '"
            _ErrorMsg += SF.ErrorMsg
            Return False
        End If
        Form1.AppendText("Response = " + Response, UUT:=UUT)
        System.Threading.Thread.Sleep(50)
        If Not SF.Cmd(UUT, Response, "y " + Chr(13) + Chr(10), 20, "Enter expected hydrogen (gas phase ppm): ", False, True, False) Then
            _ErrorMsg = "Problem sending 'y' in response to 'Y/N)? '"
            _ErrorMsg += SF.ErrorMsg
            Return False
        End If
        Form1.AppendText("Response = " + Response, UUT:=UUT)
        System.Threading.Thread.Sleep(50)
        If Not SF.Cmd(UUT, Response, tm8_gas_ppm.ToString + Chr(13) + Chr(10), 30, "Month: ", False, True, False) Then
            _ErrorMsg = "Problem entering gas ppm" + vbCr
            _ErrorMsg += SF.ErrorMsg
            Return False
        End If
        Form1.AppendText("Response = " + Response, UUT:=UUT)
        System.Threading.Thread.Sleep(50)
        If Not SF.Cmd(UUT, Response, MM + Chr(13) + Chr(10), 30, "Day: ", False, True, False) Then
            _ErrorMsg = "Problem entering calibration month" + vbCr
            _ErrorMsg += SF.ErrorMsg
            Return False
        End If
        Form1.AppendText("Response = " + Response, UUT:=UUT)
        System.Threading.Thread.Sleep(50)
        If Not SF.Cmd(UUT, Response, DD + Chr(13) + Chr(10), 30, "Year: ", False, True, False) Then
            _ErrorMsg = "Problem entering calibration day" + vbCr
            _ErrorMsg += SF.ErrorMsg
            Return False
        End If
        Form1.AppendText("Response = " + Response, UUT:=UUT)
        System.Threading.Thread.Sleep(50)
        If Not SF.Cmd(UUT, Response, YYYY + Chr(13) + Chr(10), 30, "H2scan: ", False, True, False) Then
            _ErrorMsg = "Problem entering calibration year" + vbCr
            _ErrorMsg += SF.ErrorMsg
            Return False
        End If
        Form1.AppendText("Response = " + Response, UUT:=UUT)
        If Not Response.Contains(CalDate) Then
            _ErrorMsg = Response
            _ErrorMsg += "Expected:  " + CalDate
            _ErrorMsg += SF.ErrorMsg
            Return False
        End If

        Return True
    End Function

    Public Function GetRecData(ByRef UUT As Hashtable, ByVal num_runs As Integer, ByRef RecData As DataTable, Optional ByVal Reverse As Boolean = False) As Boolean
        Dim SF As New SerialFunctions
        Dim Response As String
        Dim Line As String
        Dim Cmd As String
        Dim Fields() As String
        Dim row As DataRow
        Dim Name As String
        Dim TypeStr As String
        Dim pstZone As TimeZoneInfo = TimeZoneInfo.FindSystemTimeZoneById("Pacific Standard Time")
        Dim total_recs As Integer
        Dim n As Integer
        Dim success As Boolean
        Dim retryCnt As Integer

        Dim Tm1RecFields() As String = Tm1Rev0RecFields
        If UUT("TM1 Hardware Version") <> Products(Product)("hardware version 0") Then
            Tm1RecFields = Tm1Rev2RecFields
        End If

        If Reverse Then
            System.Threading.Thread.Sleep(150)
            Form1.DebugLog("sending cmd rec")
            If Not SF.Cmd(UUT, Response, "rec", 10) Then
                _ErrorMsg = "Problem sending command 'rec'"
                Return False
            End If
            For Each Line In Response.Split(Chr(10), Chr(13))
                Line = Line.Trim
                Form1.AppendText(Line, UUT:=UUT)
                If Regex.IsMatch(Line, "total_recs = \d+") Then
                    total_recs = (Regex.Split(Line, "total_recs = (\d+)"))(1)
                End If
            Next
            n = total_recs - num_runs
        End If

        If Reverse Then
            Cmd = "rec -N -D -CSV " + n.ToString + " " + num_runs.ToString
        Else
            Cmd = "rec -D -CSV " + num_runs.ToString
        End If

        success = False
        retryCnt = 0
        While Not success And retryCnt < 3
            retryCnt += 1
            System.Threading.Thread.Sleep(150)
            Form1.AppendText("Sending cmd " + Cmd)
            Try
                If Not SF.Cmd(UUT, Response, Cmd, 10) Then
                    _ErrorMsg = "Problem sending command '" + Cmd + "'"
                    Return False
                End If
                success = True
            Catch ex As Exception
                Form1.AppendText("Error sending cmd " + Cmd)
                Form1.AppendText(ex.ToString)
            End Try
        End While

        For Each Line In Response.Split(Chr(10), Chr(13))
            Line = Line.Trim
            If Regex.IsMatch(Line, "^\d+") Then
                Fields = Split(Line, ",")
                If Not Fields.Count = Tm1RecFields.Count Then
                    Continue For
                    '_ErrorMsg = "rec data line has " + Fields.Count.ToString + " fields, expecting " + Tm1RecFields.Count.ToString + vbCr
                    '_ErrorMsg += "rec line:  " + Line
                    'Return False
                End If
                row = RecData.NewRow()
                For i = 0 To UBound(Tm1RecFields)
                    Try
                        Name = Split(Tm1RecFields(i), "|")(0)
                        TypeStr = Split(Tm1RecFields(i), "|")(1)
                        Select Case TypeStr
                            Case "DT"
                                row(Name) = TimeZoneInfo.ConvertTimeFromUtc(DateTime.Parse(Fields(i)), pstZone)
                            Case "D"
                                row(Name) = CDbl(Fields(i))
                            Case "I"
                                row(Name) = CInt(Fields(i))
                            Case "UI"
                                row(Name) = CUInt(Fields(i))
                            Case "S"
                                row(Name) = Fields(i).ToString
                            Case Else
                                ErrorMsg = "No conversion for heater field " + Name + " with type " + TypeStr
                                Return False
                        End Select
                    Catch ex As Exception
                        _ErrorMsg = "Error parsing field " + Tm1RecFields(i).ToString + " from " + UUT("SN").Text + vbCr
                        _ErrorMsg += "Rec line:  " + Line + vbCr
                        _ErrorMsg += ex.ToString
                        Form1.AppendText(_ErrorMsg, UUT:=UUT)
                        Return False
                    End Try
                Next
                Try
                    RecData.Rows.Add(row)
                Catch ex As Exception
                    Form1.AppendText(UUT("SN") + ":  Problem adding tm1 rec line '" + Line + "' to table", UUT:=UUT)
                    Return False
                End Try
            End If
        Next

        Return True
    End Function

    Function ClearMemory(ByRef UUT As Hashtable, Optional ByVal Quiet As Boolean = False) As Boolean
        Dim SF As New SerialFunctions
        Dim Response As String
        Dim X_found As Boolean
        Dim retryCnt As Integer

        System.Threading.Thread.Sleep(50)
        If Not SF.Cmd(UUT, Response, "is" + Chr(13) + Chr(10), 30, "Y/N)? ", Quiet, True, False) Then
            _Results = Response
            _ErrorMsg = "Sent 'is' command, did not see expected '(Y/N)? '"
            _ErrorMsg += SF.ErrorMsg
            Return False
        End If
        _Results = Response
        If Not Quiet Then Form1.AppendText(Response, UUT:=UUT)
        System.Threading.Thread.Sleep(50)
        If Not SF.Cmd(UUT, Response, "y", 20, "Clearing log", Quiet, True, False, True) Then
            _Results += vbCr + Response
            _ErrorMsg = "Problem sending 'y' in response to 'Y/N)? '"
            _ErrorMsg += SF.ErrorMsg
            Return False
        End If
        _Results += vbCr + Response
        If Not Quiet Then Form1.AppendText(Response, UUT:=UUT)
        X_found = False
        retryCnt = 0
        While Not X_found And retryCnt < 3
            retryCnt += 1
            If SF.Cmd(UUT, Response, Chr(13) + Chr(10), 50, "X", Quiet, True, False, True) Then
                X_found = True
            End If
        End While
        If Not X_found Then
            If Not Quiet Then Form1.AppendText("Expected to see 'X&' to signal comletion of memory clearing", UUT:=UUT)
            _Results += vbCr + "Expected to see 'X&' to signal comletion of memory clearing"
            Return False
        End If

        Return True
    End Function

    Function ClearRecords(ByRef UUT As Hashtable, Optional ByVal Quiet As Boolean = False) As Boolean
        Dim SF As New SerialFunctions
        Dim Response As String

        System.Threading.Thread.Sleep(150)
        If Not SF.Cmd(UUT, Response, "t c" + Chr(13) + Chr(10), 20, "Y/N)? ", Quiet, True, False) Then
            _Results = Response
            _ErrorMsg = "Sent t c command, did not see expected '(Y/N)? '"
            _ErrorMsg += SF.ErrorMsg
            Return False
        End If
        _Results = Response
        If Not Quiet Then Form1.AppendText(Response, UUT:=UUT)
        System.Threading.Thread.Sleep(50)
        If Not SF.Cmd(UUT, Response, "y " + Chr(13) + Chr(10), 50, "H2scan: ", Quiet, True, False, True) Then
            _Results += vbCr + Response
            _ErrorMsg = "Problem sending 'y' in response to 'Y/N)? '"
            _ErrorMsg += SF.ErrorMsg
            Return False
        End If
        _Results += vbCr + Response
        If Not Quiet Then Form1.AppendText(Response, UUT:=UUT)

        Return True
    End Function

    Function DumpRecords(ByRef UUT As Hashtable, ByRef productInfo As Hashtable, Optional ByVal Quiet As Boolean = False) As Boolean
        Dim SF As New SerialFunctions
        Dim Response As String
        Dim LogFile As FileStream
        Dim LogFileWriter As StreamWriter
        Dim LogFilePath As String
        Dim done As Boolean = False
        Dim startTime As DateTime = Now
        Dim Line As String
        Dim ReadTimeout As Boolean
        Dim TimeoutLabelUpdateTime As DateTime

        LogFilePath = ReportDir + "FINAL_TEST\" + UUT("SN").Text + "\H2SCAN_DATA." + Format(Date.UtcNow, "yyyyMMddHHmmss") + ".txt"
        Try
            LogFile = New FileStream(LogFilePath, FileMode.Create, FileAccess.Write)
            LogFileWriter = New StreamWriter(LogFile)
        Catch ex As Exception
            _ErrorMsg = "Problem creating logfile " + LogFilePath + vbCr
            _ErrorMsg += ex.ToString
            Return False
        End Try

        ' Record the H2Scan product info
        Try
            LogFileWriter.WriteLine("TM1 SN: " + UUT("SN").Text)
            For Each record In productInfo
                LogFileWriter.WriteLine(record.key.ToString + ": " + record.value.ToString)
            Next
        Catch ex As Exception
            _ErrorMsg += ex.Message
        End Try

        If Not SF.Cmd(UUT, Response, "t d" + Chr(13) + Chr(10), 50, "Y/N)? ", Quiet, True, False, True) Then
            _Results = Response
            _ErrorMsg = "Sent t d command, did not see expected '(Y/N)? '"
            _ErrorMsg += SF.ErrorMsg
            LogFileWriter.Close()
            LogFile.Close()
            Return False
        End If
        _Results = Response
        System.Threading.Thread.Sleep(50)
        UUT("SP").write("y" + Chr(13) + Chr(10))

        done = False
        UUT("SP").ReadTimeout = 2000
        ReadTimeout = False
        TimeoutLabelUpdateTime = Now
        While (Not done And Now.Subtract(startTime).TotalMinutes < 30)
            Application.DoEvents()
            If Not Quiet Then
                If Now.Subtract(TimeoutLabelUpdateTime).TotalSeconds > 30 Then
                    Form1.TimeoutLabel.Text = Math.Round(30 - Now.Subtract(startTime).TotalMinutes, 1).ToString + "m"
                    TimeoutLabelUpdateTime = Now
                End If
            End If
            Try
                Line = UUT("SP").ReadLine()
                LogFileWriter.Write(Line + System.Environment.NewLine)
                If Line.Contains("H2scan: ") Then
                    done = True
                End If
            Catch ex As Exception
                ReadTimeout = True
                Exit While
            End Try
        End While
        If Not done Then
            If ReadTimeout Then
                _Results += vbCr + "serial port read line timeout"
                If Not Quiet Then Form1.AppendText("serial port read line timeout", UUT:=UUT)
            Else
                _Results += vbCr + "Timeout dumping H2SCAN logs"
                If Not Quiet Then Form1.AppendText("Timeout dumping H2SCAN logs", UUT:=UUT)
            End If
            Return False
        End If

        ' Flush and close the streams
        Try
            If LogFileWriter IsNot Nothing Then LogFileWriter.Close()
            If LogFile IsNot Nothing Then LogFile.Close()
        Catch ex As Exception
        End Try

        Return True
    End Function

    Public Function GetProductInfo(ByRef UUT As Hashtable, ByRef ProductInfo As Hashtable, Optional ByVal Quiet As Boolean = False) As Boolean
        Dim SF As New SerialFunctions
        Dim Response As String
        Dim Line As String
        Dim SerialPort As Object = UUT("SP")
        Dim Fields() As String
        Dim ExpectedFields() As String = {"Model Number", "Serial Number", "Sensor Number", "Firmware Rev",
                                          "Table Version", "Hardware Version", "Factory", "TouchUp", "DGA"}
        Dim Success As Boolean

        If Not SF.Cmd(UUT, Response, "d0" + Chr(13) + Chr(10), 5, "H2scan: ", Quiet, True, False) Then
            _Results = Response
            _ErrorMsg = "Problem sending 'd0' command to H2SCAN, didn't see prompt 'H2scan: '"
            _ErrorMsg += SF.ErrorMsg
            Return False
        End If
        _Results = Response
        If Not Quiet Then Form1.AppendText("Response = " + Response)
        ProductInfo = New Hashtable
        For Each Line In Regex.Split(Response, "\r\n")
            If Not Line.Contains(":") Then Continue For
            Line = Regex.Replace(Line, "^\s+", "")
            Line = Regex.Replace(Line, "\s+:\s+", "=")
            Line = Regex.Replace(Line, "\s+$", "")
            Fields = Split(Line, ":")
            If Fields.Length = 2 Then
                ProductInfo.Add(Fields(0), Fields(1))
            End If
        Next

        Success = True
        _ErrorMsg = ""
        For Each ExpectedField In ExpectedFields
            If Not ProductInfo.Contains(ExpectedField) Then
                _ErrorMsg += "H2SCAN product info missing " + ExpectedField + vbCr
                Success = False
            End If
        Next

        Return Success
    End Function
End Class
'