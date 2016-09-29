Imports System.Text.RegularExpressions
Imports System.IO.Ports

Public Class SerialFunctions
    Private _ErrorMsg As String = ""
    Property ErrorMsg() As String
        Get
            Return _ErrorMsg
        End Get
        Set(value As String)

        End Set
    End Property

    Private _RtnResults As String
    Property RtnResults() As String
        Get
            Return _RtnResults
        End Get
        Set(value As String)

        End Set
    End Property


    Public Function GetPrompt(ByVal SerialPort As Object, Optional ByVal Prompt As String = "> ") As Boolean
        Dim Data As Integer
        Dim PromptFound As Boolean = False
        Dim StartTime As DateTime = Now
        Dim Buffer As String
        Dim results As ReturnResults
        Dim SF As New SerialFunctions
        Dim CommandTerm = Chr(10)

        If Prompt = "H2scan: " Then
            CommandTerm = Chr(13) + Chr(10)
        End If

        Try
            'System.Threading.Thread.Sleep(50)
            'SerialPort.ReadExisting()
            SerialPort.DiscardInBuffer()
            'SerialPort.Write(Chr(10))
            System.Threading.Thread.Sleep(50)
            'SerialPort.Write(Chr(13) + Chr(10))
            'SerialPort.Write(Chr(10))
            SerialPort.Write(CommandTerm)
            While (Not PromptFound And Now.Subtract(StartTime).TotalSeconds < 2)
                Try
                    Application.DoEvents()
                    Data = SerialPort.ReadChar()
                    Buffer += Chr(Data)
                    'If (Buffer.EndsWith("> ")) Then
                    If (Buffer.EndsWith(Prompt)) Then
                        PromptFound = True
                    End If
                Catch ex As Exception
                    'SerialPort.Write(Chr(13) + Chr(10))
                    'Buffer = ""
                    'System.Threading.Thread.Sleep(200)
                End Try
            End While

            If Not PromptFound Then
                results = SF.Login(SerialPort)
                If Not results.PassFail Then
                    _ErrorMsg = "Did not see prompt"
                    Return False
                End If
                System.Threading.Thread.Sleep(200)
            End If
        Catch ex As Exception
            Form1.AppendText("SerialFunctions.GetPrompt() caught exception: " + ex.Message)
            Return False
        End Try

        Return True
    End Function

    Public Function Cmd(ByVal UUT As Hashtable, ByRef Response As String, ByVal Command As String, ByVal Timeout As Integer, Optional ByVal Prompt As String = "> ", Optional ByVal Quiet As Boolean = False,
                        Optional ByVal NoCR As Boolean = False, Optional ByVal GetPromptFirst As Boolean = True,
                        Optional ByVal DisplayResults As Boolean = False, Optional ByVal DisplayTimeout As Boolean = False)
        Dim PromptFound As Boolean
        Dim StartTime As DateTime
        Dim Data As Integer
        Dim PostCommandInput As String = "NONE"
        Dim CommandTerm = Chr(10)
        'Dim CommandTerm = Chr(13) + Chr(10)
        Dim SerialPort As Object = UUT("SP")

        If Prompt = "H2scan: " Then
            CommandTerm = Chr(13) + Chr(10)
        End If

        Try
            If GetPromptFirst Then
                If Not GetPrompt(SerialPort, Prompt) Then
                    _ErrorMsg = "Did not see prompt '" + Prompt + "' before sending cmd"
                    Return False
                End If
            End If
            'SerialPort.ReadExisting()
            SerialPort.DiscardInBuffer()

            If (Not Quiet) Then Form1.AppendText(Command + System.Environment.NewLine, True, UUT:=UUT)
            'System.Threading.Thread.Sleep(250)
            System.Threading.Thread.Sleep(50)
            For Each C As Char In Command.ToCharArray
                SerialPort.Write(C)
                System.Threading.Thread.Sleep(2)
            Next
            'SerialPort.Write(Command)
            If Not NoCR Then
                'System.Threading.Thread.Sleep(25)
                SerialPort.Write(CommandTerm)
            End If
            PromptFound = False
            Response = ""
            If Command = "config reset factory" Then
                Prompt = "(y/n)? "
                PostCommandInput = "y"
            End If
            If Command = "config set CLI_OVER_TMCOM1.ENABLE true" Or Command = "fr -E" Or Command = "fr -A" Then
                Prompt = "reboot the system now? "
                PostCommandInput = "y"
            End If
            If Command = "reboot" Then
                PromptFound = True
                CommonLib.Delay(5)
            End If
            StartTime = Now
            ' Need to handle case where unit is not logged in
            While (Not PromptFound And Now.Subtract(StartTime).TotalSeconds < Timeout)
                Try
                    If DisplayTimeout Then
                        'Form1.TimeoutLabel.Text = Math.Round(Timeout - Now.Subtract(StartTime).TotalSeconds).ToString
                    End If
                    Application.DoEvents()
                    Data = SerialPort.ReadChar()
                    If DisplayResults Then
                        If Not Data = 10 And Not Data = 13 Then
                            Form1.AppendText(Chr(Data), True, False, UUT:=UUT)
                        ElseIf Data = 13 Then
                            Form1.AppendText(vbCr, True, False, UUT:=UUT)
                        End If
                    End If
                    Response += Chr(Data)
                    If (Response.EndsWith(Prompt)) Then
                        PromptFound = True
                    End If
                Catch ex As Exception

                End Try
            End While
            If Not PromptFound Then
                _ErrorMsg = "Did not see prompt:  " + Prompt
                Return False
            End If
            If Not PostCommandInput = "NONE" Then
                SerialPort.Write(PostCommandInput)
            End If
            'If Command = "config reset factory" Then
            '    SerialPort.Write("y")
            '    Return True
            'End If

            ' Trim command & prompt from response
            Response = Regex.Replace(Response, "^" + Command + ".*" + Chr(10), "")
            If Quiet Then
                Try
                    Response = Regex.Replace(Response, Chr(13) + Chr(10) + Prompt + ".*$", "")
                Catch ex As Exception
                    '_ErrorMsg = "Problem parsing '" + Response + "'"
                    '_ErrorMsg += vbCr + ex.ToString
                    'Return False
                End Try
            End If
        Catch ex As Exception
            Form1.AppendText("SerialFunctions.Cmd() caught exception: " + ex.Message)
            Return False
        End Try

        Return True
    End Function

    Public Function LogOffTm1(ByRef SP As SerialPort) As Boolean
        Try
            If SP.IsOpen Then
                SP.WriteLine("Exit")
                Threading.Thread.Sleep(250)
                SP.Close()
                Threading.Thread.Sleep(250)
            End If
        Catch ex As Exception
            Form1.AppendText("SerialFunctions.LogOffTm1() caught " + ex.Message)
            Return False
        End Try

        Return True
    End Function

    Public Function Login(SP As Object, Optional ByVal WaitRecordsFile As Boolean = False, Optional ByVal Quiet As Boolean = False) As ReturnResults
        Dim prompt_found As Boolean
        Dim Results As New ReturnResults
        Dim startTime As DateTime
        Dim Data As Integer
        Dim Buffer As String
        Dim Timeout As Integer = 10
        Dim NotLoggedIn As Boolean = False
        Dim DataReceivedTime As DateTime = Now

        Results.PassFail = False

        If WaitRecordsFile Then Timeout = 300

        If Not SP.IsOpen Then
            Results.PassFail = False
            Results.Result = "Serial port is not open"
            Return Results
        End If

        prompt_found = False
        startTime = Now
        'SP.RtsEnable = True
        While (Not prompt_found And Now.Subtract(startTime).TotalSeconds < Timeout And Not Stopped)
            Application.DoEvents()
            Try
                Data = SP.ReadChar()
                DataReceivedTime = Now
                If Not Data = 13 Then
                    _RtnResults += Chr(Data)
                    If Not Quiet Then Form1.AppendText(Chr(Data), True, False)
                    If WaitRecordsFile Then
                        Timeout = 10
                        WaitRecordsFile = False
                        startTime = Now
                    End If
                End If
                Buffer += Chr(Data)
                If (Buffer.EndsWith("username: ")) Then
                    System.Threading.Thread.Sleep(200)
                    SP.Write("BPLG" + Chr(10) + "CommandTimeout" + Chr(10))
                    NotLoggedIn = True
                End If
                If (Buffer.EndsWith("> ")) Then
                    prompt_found = True
                End If
                If (Buffer.EndsWith("H2scan: ")) Then
                    SP.Write(Chr(4))
                End If
                If (Buffer.Contains(": ") Or Buffer.Contains("> ")) Then
                    Buffer = ""
                End If

            Catch ex As Exception
                Try
                    SP.Write(Chr(13) + Chr(10))
                Catch exx As Exception
                    Results.Result = exx.ToString
                    Exit While
                End Try
                Buffer = ""
                If Now.Subtract(DataReceivedTime).TotalSeconds > 10 Then
                    CommonLib.Delay(10)
                Else
                    System.Threading.Thread.Sleep(200)
                End If
            End Try
        End While
        Results.PassFail = prompt_found
        If NotLoggedIn Then System.Threading.Thread.Sleep(300)

        Return Results
    End Function

    'Public Function Connect(ByRef SerialPort As Object) As ReturnResults
    Public Function Connect(ByRef UUT As Hashtable, Optional ByVal Quiet As Boolean = False) As ReturnResults
        Dim ftdi_device As New FT232R
        Dim Comport As String
        Dim Results As New ReturnResults
        Dim retryCnt As Integer
        Dim success As Boolean

        Results.PassFail = False
        _RtnResults = ""

        ' End the CLI session so the new one can be freshly start
        If Not LogOffTm1(UUT("SP")) Then Return Results

        If Not UUT("SP").IsOpen Then
            _RtnResults += "Opening serial port"
            If Not Quiet Then Form1.AppendText("Opening serial port", True, UUT:=UUT)
            retryCnt = 0
            success = False
            While retryCnt < 10 And Not success
                If retryCnt > 0 Then
                    CommonLib.Delay(10)
                End If
                retryCnt += 1
                Try
                    UUT("SP").Open()
                    success = True
                Catch ex As Exception
                End Try
            End While
            If Not success Then
                _RtnResults += "Problem opening serial port" + System.Environment.NewLine
                If Not Quiet Then Form1.AppendText("Problem opening serial port", True, UUT:=UUT)
                Return Results
            End If
        End If

        _RtnResults += "logging in to TM1" + System.Environment.NewLine
        If Not Quiet Then Form1.AppendText("logging in to TM1", True, UUT:=UUT)
        UUT("SP").ReadTimeout = 1000
        ' Allow up to three tries
        For attempt As Integer = 1 To 5
            CommonLib.Delay(3)
            Results = Login(UUT("SP"), Quiet:=Quiet)
            If Results.PassFail Then
                Exit For
            End If
        Next
        If Not Results.PassFail Then
            If Not Quiet Then
                Form1.AppendText("Login failed", True, UUT:=UUT)
                Form1.AppendText(Results.Result, True, UUT:=UUT)
            End If
            _RtnResults += "Login failed" + System.Environment.NewLine
            _RtnResults += Results.Result + System.Environment.NewLine
            Return Results
        End If
        'Form1.AppendText("Logged in", True)

        Return Results
    End Function

    Public Function Reboot(ByRef UUT As Hashtable, Optional ByVal Quiet As Boolean = False) As ReturnResults
        Dim Results As New ReturnResults
        Dim Response As String
        Dim SerialPort As Object = UUT("SP")

        Results.PassFail = False
        If Not SerialPort.IsOpen Then
            Results.PassFail = False
            Results.Result = "Serial port is not open"
            Return Results
        End If
        If Not Cmd(UUT, Response, "reboot", 10, Quiet:=Quiet) Then
            _RtnResults = Response
            Return Results
        End If
        _RtnResults = Response
        Results = Login(SerialPort, Quiet:=Quiet)

        Return Results
    End Function

    Public Function Close(ByVal UUT As Hashtable) As Boolean
        Dim SerialPort As Object = UUT("SP")
        If Not SerialPort Is Nothing Then
            If SerialPort.IsOpen Then
                Try
                    SerialPort.Close()
                Catch ex As Exception

                End Try
            End If
        End If

        Return True
    End Function


End Class
