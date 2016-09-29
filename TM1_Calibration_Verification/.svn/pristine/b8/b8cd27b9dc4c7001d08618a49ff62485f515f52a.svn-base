Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions

Public Class config
    Private _Error_Message As String
    Property ErrorMsg() As String
        Get
            Return _Error_Message
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

    Public Function ReadExpectedConfig(ByVal ConfigName As String, ByRef Config As Hashtable) As Boolean
        Dim res() As String
        Dim swriter As Stream
        Dim ConfigData As String
        Dim ConfigFile As String = "NOT FOUND"
        Dim Line As String
        Dim Fields() As String
        Dim Name, Value As String

        Config = New Hashtable

        res = Me.GetType.Assembly.GetManifestResourceNames()
        For i = 0 To UBound(res)
            If res(i).EndsWith(".txt") Then
                If res(i).Contains(ConfigName) Then
                    ConfigFile = res(i)
                End If
            End If
        Next
        If ConfigFile = "NOT FOUND" Then
            _Error_Message = "Can't find embedded resource file " + ConfigName
            Return False
        End If

        Try
            swriter = Me.GetType.Assembly.GetManifestResourceStream(ConfigFile)
            Dim Bytes(swriter.Length) As Byte
            swriter.Read(Bytes, 0, swriter.Length)
            ConfigData = Encoding.ASCII.GetString(Bytes)
            'ConfigData = Regex.Replace(ConfigData, "\?", "")
            ConfigData = Regex.Replace(ConfigData, "\?\?\?", "")
        Catch ex As Exception
            _Error_Message = "Problem reading embedded resource file " + ConfigName + vbCr
            _Error_Message += ex.ToString
            Return False
        End Try

        For Each Line In Regex.Split(ConfigData.Substring(0, ConfigData.Length - 1), "\r\n")
            Fields = Regex.Split(Line, "=")
            Try
                Name = Fields(0)
                Value = Fields(1)
            Catch ex As Exception
                _Error_Message = "Problem extracting name/value from " + Line
                Return False
            End Try
            Name = Regex.Replace(Name, "^\s+", "")
            Name = Regex.Replace(Name, "\s+$", "")
            Value = Regex.Replace(Value, "^\s+", "")
            Value = Regex.Replace(Value, "\s+$", "")
            Config.Add(Name, Value)
        Next

        Return True
    End Function

    Public Function ConfigFactoryReset(ByVal UUT As Hashtable, Optional ByVal Quiet As Boolean = False) As Boolean
        Dim SF As New SerialFunctions
        Dim results As ReturnResults
        Dim Response As String
        Dim ConfigHash As Hashtable
        Dim config_name As String
        Dim T As New TM1
        Dim Tm1Config As Hashtable
        Dim ConfigNames As ICollection
        Dim ConfigNamesArray() As String

        config_name = "default_config_" + Products(Product)("FW_VERSION")

        If Not SF.Cmd(UUT, Response, "config reset factory", 20, Quiet:=Quiet) Then
            _Results = Response + System.Environment.NewLine
            If Not Quiet Then
                Form1.AppendText("failed sending cmd 'config reset factory'")
                Form1.AppendText(SF.ErrorMsg)
            End If
            _Results += "failed sending cmd 'config reset factory'" + System.Environment.NewLine
            _Results += SF.ErrorMsg + System.Environment.NewLine
            Return False
        End If
        _Results = Response

        CommonLib.Delay(20)
        results = SF.Connect(UUT, Quiet)
        _Results += SF.RtnResults + System.Environment.NewLine
        If Not results.PassFail Then
            _Error_Message = SF.ErrorMsg
            Return False
        End If

        If Not ReadExpectedConfig(config_name, ConfigHash) Then
            Return False
        End If

        If Not ConfigHash.Contains("board.serial_number") Then
            _Error_Message = "default config missing 'board.serial_number'"
            Return False
        End If
        If Not UUT("LI").Contains("CONTROLLER_BOARD") Then
            _Results += "Missing link data for " + UUT("SN").Text + System.Environment.NewLine
            If Not Quiet Then Form1.AppendText("Missing link data for " + UUT("SN").Text, UUT:=UUT)
            Return False
        End If
        ConfigHash("board.serial_number") = UUT("LI")("CONTROLLER_BOARD")

        'TODO:  Get board serial number from parent child database
        CommonLib.Delay(2)
        If Not T.GetConfig(UUT, Tm1Config) Then
            _Results += T.ErrorMsg + System.Environment.NewLine
            If Not Quiet Then Form1.AppendText(T.ErrorMsg)
            Return False
        End If

        ConfigNames = ConfigHash.Keys
        ReDim ConfigNamesArray(ConfigHash.Count - 1)
        ConfigNames.CopyTo(ConfigNamesArray, 0)
        Array.Sort(ConfigNamesArray)

        For Each k In Tm1Config.Keys
            If Not ConfigHash.Contains(k) Then
                _Error_Message = "Unexpected config:  " + k
                Return False
            End If
        Next

        _Results += System.Environment.NewLine
        For Each k In ConfigNamesArray
            If Not Tm1Config.Contains(k) Then
                _Error_Message = "Did not find config entry for " + k
                Return False
            End If
            _Results += k + " = " + Tm1Config(k) + System.Environment.NewLine
            If Not Quiet Then Form1.AppendText(k + " = " + Tm1Config(k), UUT:=UUT)
            If Not Tm1Config(k) = ConfigHash(k) Then
                _Error_Message = "   Expected:  " + ConfigHash(k)
                Return False
            End If
        Next

        Return True
    End Function

End Class
