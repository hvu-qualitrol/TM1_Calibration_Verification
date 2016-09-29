Imports System.Text

Public Class Telnet
    WithEvents mTelNet As TelNetSocketLayer
    Private Inbuffer As String
    Private LastCommand As String
    Public IPAddress As String
    Public Port As Int16
    Private Buffer As StringBuilder = New StringBuilder(32000)
    Private LoggedIn As Boolean
    Private CarrotFound As Boolean
    Private Prompt As String = ">"

    Event Connected(ByVal State As Boolean)
    Event UserLoggedIn()
    Event GotData(ByVal Data As String)
    Event Excpt(ByVal Evnt As String)
    Event Abort()
    Event ReSend()
    Event Connecting()

    Public ReadOnly Property PortState() As Boolean
        Get
            Return mTelNet.tns.PortState
        End Get
    End Property
    Public ReadOnly Property UserLogin() As Boolean
        Get
            Return mTelNet.tns.UserLogin
        End Get
    End Property
    Public ReadOnly Property UserClose() As Boolean
        Get
            Return mTelNet.tns.CloseRead
        End Get
    End Property
    Public ReadOnly Property CmdResult() As String
        Get
            Return Buffer.ToString()
        End Get
    End Property

    Public Function ReturnResult() As String
        Return Buffer.ToString
    End Function
    Public Function Open(ByVal IPAddr As String, ByVal Prt As Int16)
        Dim startTime As DateTime = Now

        If Not (mTelNet Is Nothing) Then
            mTelNet.Dispose()
        End If
        mTelNet = New TelNetSocketLayer
        IPAddress = IPAddr
        Port = Prt
        If Not mTelNet.Open(IPAddr, Prt) Then
            Return False
        End If

        While (Now.Subtract(startTime).TotalSeconds < 30) And Not Me.UserLogin
            Application.DoEvents()
            System.Threading.Thread.Sleep(300)
        End While
        Return Me.UserLogin
        Return True
    End Function
    Public Function Open()
        RaiseEvent Connecting()
        If Not (mTelNet Is Nothing) And PortState = True Then
            mTelNet.Dispose()
            'mTelNet = Nothing
        End If
        mTelNet = New TelNetSocketLayer
        If Not mTelNet.Open(IPAddress, Port) Then
            Return False
        End If
        Return True
    End Function
    Public Function Command(ByVal cmd As String, Optional ByVal timeout As Integer = 10, Optional ByVal PromptChar As String = ">", Optional add_cr As Boolean = True) As Boolean
        Dim startTime As DateTime = Now

        Prompt = PromptChar

        Buffer.Length = 0
        If add_cr Then
            cmd += vbCr
        End If
        SendData(cmd)
        If timeout = 0 Then
            Return True
        End If
        'SendData(cmd & vbCr)
        CarrotFound = False
        Do
            Application.DoEvents()
        Loop Until CarrotFound = True Or Now.Subtract(startTime).TotalSeconds > timeout
        Return CarrotFound
    End Function

    Public Sub SendData(ByVal data As String, Optional ByVal Drop As Boolean = False)
        If data = "" Then Exit Sub
        If PortState = True Then LastCommand = data
        mTelNet.SendData(data)
        'If data.ToLower() = "reboot" & vbCr And Drop = False Then
        '    LastCommand = ""
        '    Pause(120)
        '    ReConnect()
        '    SendData(" " & vbCr)
        'End If
    End Sub
    Public Function CloseTelnet() As Boolean
        mTelNet.Dispose()

        mTelNet = Nothing
        Return True
    End Function

    Private Sub ConnectionEstablished(ByVal Connected As Boolean) Handles mTelNet.Connected
        mTelNet.tns.UserLogin = False
        If mTelNet.tns.Connected = False Then
            MsgBox("Telnet failed to establish a connection. Please check your IP Address.")
            RaiseEvent Abort()
            Exit Sub
        End If
        If Login() = False Then
            mTelNet.tns.UserLogin = False
            MsgBox("Unable to Login to the analyzer.")
            RaiseEvent Abort()
            Exit Sub
        End If
        mTelNet.tns.UserLogin = True
        RaiseEvent Connected(True)
    End Sub
    Private Sub EventPortState(ByVal Port As Boolean) Handles mTelNet.PortState
        mTelNet.tns.PortState = Port
    End Sub
    Private Function Login() As Boolean
        ' Send Login Name
        Pause(2)
        If ResponseOrTimeOut("username:") = True Then
            Return mTelNet.tns.UserLogin
        End If
        Inbuffer = String.Empty
        mTelNet.SendData("Serveron" & vbCr)

        If ResponseOrTimeOut("password:") = True Then
            Return mTelNet.tns.UserLogin
        End If
        Inbuffer = String.Empty
        mTelNet.SendData("LoginFailed" & vbCr)

        If ResponseOrTimeOut(">") = True Then
            Return mTelNet.tns.UserLogin
        End If
        mTelNet.tns.UserLogin = True
        RaiseEvent UserLoggedIn()
        If LastCommand <> "" Then
            RaiseEvent ReSend()
        End If
        Return mTelNet.tns.UserLogin
    End Function
    Private Sub EventGotData(ByVal data As String) Handles mTelNet.GotData
        If mTelNet.tns.UserLogin = False Then
            Inbuffer += data
        Else
            Inbuffer = String.Empty
        End If

        Buffer.Append(data)
        If UserLogin = True Then
            Call CommandSet()
        Else
            Call EstablishConn()
        End If
    End Sub
    Private Sub EstablishConn()
        If InStr(Buffer.ToString, ">", CompareMethod.Text) > 0 Then
            LoggedIn = True
            Buffer.Length = 0
            CarrotFound = True
        End If
    End Sub
    Private Sub CommandSet()
        'If InStr(Buffer.ToString, ">", CompareMethod.Text) > 0 Then
        If InStr(Buffer.ToString, Prompt, CompareMethod.Text) > 0 Then
            CarrotFound = True
        End If
    End Sub
    Private Function ResponseOrTimeOut(ByVal Response As String) As Boolean
        Dim TimedOut As Boolean
        Dim Timr As Date
        Timr = Now.AddSeconds(30)
        TimedOut = True
        While DateTime.Compare(DateTime.Now, Timr) < 0
            Application.DoEvents()
            If InStr(Inbuffer, Response, CompareMethod.Text) > 0 Then
                TimedOut = False
                Exit While
            End If
        End While
        Return TimedOut
    End Function
    Private Sub ReConnect() Handles mTelNet.ReConnect
        Open()
    End Sub
    Private Sub ReSendCommand()
        Dim Timr As Date
        Timr = Now.AddSeconds(300)
        While UserLogin = False And DateTime.Compare(DateTime.Now, Timr) < 0
            Application.DoEvents()
        End While
        If UserLogin = True Then
            SendData(LastCommand)
        End If
    End Sub
    Private Sub Pause(ByVal Seconds As Int32)
        Static start_time As DateTime
        Static stop_time As DateTime
        Dim elapsed_time As TimeSpan
        start_time = Now
        Do
            System.Windows.Forms.Application.DoEvents()
            stop_time = Now
            elapsed_time = stop_time.Subtract(start_time)
            System.Threading.Thread.Sleep(50)
            Application.DoEvents()
        Loop Until elapsed_time.TotalSeconds >= Seconds
    End Sub
End Class
