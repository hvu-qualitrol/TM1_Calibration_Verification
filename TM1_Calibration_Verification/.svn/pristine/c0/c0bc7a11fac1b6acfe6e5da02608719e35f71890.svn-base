Imports System.Net.Sockets
Imports System.Runtime.Serialization 
Imports System.Runtime.Serialization.Formatters.Binary
Imports System.Threading

Public Class TelNetSocketLayer
    Const READ_BUFFER_SIZE As Integer = 4096
    Public tns As New TelNetState

    Private client As TcpClient
    Private readBuffer(READ_BUFFER_SIZE) As Byte
    Private Lastdatasent As String
    Private Conn As Thread
    Private IPAddress As String
    Private Port As Int16

    Event Connected(ByVal State As Boolean)
    Event ReConnect()
    Event PortState(ByVal Port As Boolean)
    Event UserLoggedIn()
    Event GotData(ByVal Data As String)
    Event Excpt(ByVal Evnt As String)
    WithEvents TCPTimer As System.timers.Timer

    Public Function Open(ByVal IPAddr As String, ByVal Prt As Int16)
        Try
            tns.Connecting = True
            If IPAddr = String.Empty Or Prt = 0 Then
                Return False
            End If
            IPAddress = IPAddr
            Port = Prt
            ThreadPool.QueueUserWorkItem(AddressOf Connection)
        Catch ex As Exception
            Return False
        End Try
        Return True
    End Function

    Public Sub SendData(ByVal data As String)
        Dim bytCmd() As Byte = System.Text.Encoding.ASCII.GetBytes(data)
        Try
            client.GetStream.Write(bytCmd, 0, bytCmd.Length)
        Catch ex As Exception
            RaiseEvent PortState(False)
        End Try
    End Sub

    Private Sub WriteSent(ByVal i As IAsyncResult)

    End Sub
    Public Sub DoRead(ByVal ar As IAsyncResult)
        Dim BytesRead As Integer
        Dim strMessage As String
        Try
            ' Finish asynchronous read into readBuffer and return number of bytes read. 
            BytesRead = client.GetStream.EndRead(ar)

            ' Convert the byte array the message was saved into 
            strMessage = System.Text.Encoding.ASCII.GetString(readBuffer, 0, BytesRead)

            RaiseEvent GotData(strMessage)

            If tns.CloseRead = True Then
                MsgBox("DoRead - CloseRead Found. Exit DoRead")
                Exit Sub
            End If
            If tns.PortState = False Then
                If tns.Connecting = False Then
                    MsgBox("DoRead - Port State = False, Connecting = False")
                    Exit Sub
                End If
            End If
            ' Start a new asynchronous read into readBuffer. 
            client.GetStream.BeginRead(readBuffer, 0, READ_BUFFER_SIZE, AddressOf DoRead, Nothing)
        Catch e As Exception
            If tns.CloseRead = False Then
                RaiseEvent PortState(False)
                RaiseEvent ReConnect()
            End If
        End Try
    End Sub

    Private Sub Connection(ByVal state As Object)
        Dim TimedOut As Boolean
        Dim Timr As Date
        Dim Retry As Int32
        Dim TimeInSeconds As Int32

        'Thread.CurrentThread.Name = "Connection for " & IPAddress
        Try
DoRetry:
            Do
                TimeInSeconds = Convert.ToInt32(2 ^ Retry)
                Pause(TimeInSeconds)
                Retry += 1
                TimedOut = False

                ' The TcpClient is a subclass of Socket, providing higher level 
                ' functionality like streaming. 
                Try
                    TCPTimer = New System.Timers.Timer(60000)
                    TCPTimer.Start()
                    client = New TcpClient
                    Dim LOp As New LingerOption(False, 0)
                    client.LingerState = LOp
                    client.Connect(IPAddress, Port)

                    TCPTimer.Stop()
                    TCPTimer.Close()

                Catch ex As Exception
                    MsgBox("Telnetsocketlayer Connection - " & ex.Message)
                End Try

                Timr.AddSeconds(10)
                Timr = Now
                Do
                    Application.DoEvents()
                    If tns.CloseRead = True Then
                        tns.Failed = True
                        Exit Sub
                    End If
                Loop Until DateTime.Compare(DateTime.Now, Timr) < 0 Or client.GetStream.DataAvailable() = True
            Loop Until client.GetStream.DataAvailable() = True Or Retry > 6

            If client.GetStream.DataAvailable() = True Then
                client.GetStream.BeginRead(readBuffer, 0, READ_BUFFER_SIZE, AddressOf DoRead, Nothing)
                tns.PortState = True
                RaiseEvent PortState(True)
                tns.Connected = True
                tns.Connecting = False
                RaiseEvent Connected(True)
            Else
                tns.PortState = False
                RaiseEvent PortState(False)
                tns.Connected = False
                tns.Connecting = False
                RaiseEvent Connected(False)
            End If

        Catch E As Exception
            If Retry > 6 Then
                MsgBox(E.ToString & ". This form will now shut down", MsgBoxStyle.OkOnly, "Error")
                tns.Failed = True
                RaiseEvent Excpt(E.ToString)
            Else
                GoTo DoRetry
            End If
        End Try
    End Sub

    Private Sub TCPtimeouthandler(ByVal source As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles TCPTimer.Elapsed
        Try
            tns.TimedOut = True
            client.LingerState.Enabled = False
            Dim NetStream As NetworkStream = client.GetStream
            NetStream.Close()
            client.Close()
        Catch ex As Exception

        End Try
    End Sub

    Public Sub Dispose()
        Try
            tns.CloseRead = True
            Dim NetStream As NetworkStream = client.GetStream
            If tns.PortState = True Then
                NetStream.Close()
            End If
            tns.PortState = False

            RemoveHandler TCPTimer.Elapsed, AddressOf TCPtimeouthandler

            client.Close()
            tns.UserLogin = False

        Catch ex As Exception
            MsgBox("Telnetsocketlayer Dispose - " & ex.Message)
        End Try
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
            If tns.CloseRead = True Then
                tns.Failed = True
                Exit Sub
            End If
        Loop Until elapsed_time.TotalSeconds >= Seconds
    End Sub
End Class
