Imports FTD2XX_NET

Public Class FT232R
    Private _failure_message As String

    Public Property failure_message As String
        Get
            Return _failure_message
        End Get
        Set(value As String)

        End Set
    End Property

    Public Function Rescan() As Boolean
        Dim ftdi_device As New FTDI
        Dim ftstatus As FTDI.FT_STATUS

        ftstatus = ftdi_device.Rescan()
        If (Not ftstatus = FTDI.FT_STATUS.FT_OK) Then
            _failure_message = ftstatus.ToString
            Return False
        End If

        Return True
    End Function


    Function FtdiDeviceCount(ByRef DeviceCount As UInteger) As Boolean
        Dim ftdi_device As New FTDI
        Dim ftstatus As FTDI.FT_STATUS

        _failure_message = ""
        Try
            ftstatus = ftdi_device.GetNumberOfDevices(DeviceCount)
        Catch ex As Exception
            _failure_message = "Exception getting FTDI device count:  " + ex.ToString
            Return False
        End Try
        If (ftstatus = FTDI.FT_STATUS.FT_OK) Then
            Return True
        Else
            _failure_message = "Error getting FTDI device count: " + ftstatus.ToString
            Return False
        End If
    End Function

    ' Returns the location of ID of FT232R device.    
    Function GetLocation(ByRef Location As UInteger) As Boolean
        Dim DeviceCount As UInteger
        Dim ftstatus As FTDI.FT_STATUS
        Dim ftdi_device As New FTDI
        Dim FoundFT232R As Boolean = False

        _failure_message = ""
        If Not FtdiDeviceCount(DeviceCount) Then
            _failure_message = "FtdiDeviceCount() failed"
            Return False
        End If
        If Not (DeviceCount > 0) Then
            _failure_message = "No FTDI devices found"
            Return False
        End If

        Dim ftdiDeviceList(DeviceCount) As FTDI.FT_DEVICE_INFO_NODE
        ftstatus = ftdi_device.GetDeviceList(ftdiDeviceList)
        If Not ftstatus = FTDI.FT_STATUS.FT_OK Then
            _failure_message = "Problem getting FTDI device info"
            Return False
        End If

        For i = 0 To DeviceCount - 1
            If ftdiDeviceList(i).Type = FTDI.FT_DEVICE.FT_DEVICE_232R Then
                Form1.AppendText("device " + i.ToString + "location id " + ftdiDeviceList(i).LocId.ToString)
                ' If ftdiDeviceList(i).LocId < &H100 Then
                If ftdiDeviceList(i).LocId < &H1000 Then
                    If FoundFT232R Then
                        _failure_message = "Multiple FT232R devices found with a single port, expected 1"
                        Return False
                    End If
                    Location = ftdiDeviceList(i).LocId
                    FoundFT232R = True
                Else
                    Form1.AppendText("Skipping " + ftdiDeviceList(i).LocId.ToString)
                End If
            End If
        Next
        Return FoundFT232R
    End Function

    Public Function FindLocationForSN(ByVal SN As String, ByRef LocID As Integer) As Boolean
        Dim DeviceCount As Integer
        Dim ftstatus As FTDI.FT_STATUS
        Dim ftdi_device As New FTDI

        If Not FtdiDeviceCount(DeviceCount) Then
            Return False
        End If
        If Not DeviceCount > 0 Then
            Return False
        End If

        _failure_message = ""
        Dim startTime As DateTime = Now
        Dim ftdiDeviceList(DeviceCount) As FTDI.FT_DEVICE_INFO_NODE
        ftstatus = ftdi_device.GetDeviceList(ftdiDeviceList)
        If Not ftstatus = FTDI.FT_STATUS.FT_OK Then
            _failure_message = "Error getting FTDI device list:  " + ftstatus.ToString
            Return False
        End If
        For i = 0 To DeviceCount - 1
            Form1.AppendText("device " + i.ToString + " LocID = " + String.Format("{0:x}", ftdiDeviceList(i).LocId))
            Form1.AppendText("Checking device " + i.ToString + ", SN=" + ftdiDeviceList(i).SerialNumber + vbCr)
            Form1.AppendText("description = " + ftdiDeviceList(i).Description.ToString)
            If ftdiDeviceList(i).SerialNumber = SN Then
                Form1.AppendText("Location found for " + SN)
                LocID = ftdiDeviceList(i).LocId
                Return True
            End If
        Next
        _failure_message = "Did not find FTDI device with SN " + SN

        Return False
    End Function

    Public Function FindComportForSN(ByVal UUT As Hashtable, ByRef ComPort As String) As Boolean
        Dim DeviceCount As Integer
        Dim ftstatus As FTDI.FT_STATUS
        Dim ftdi_device As New FTDI
        Dim found_blank_SN As Boolean = False
        Dim retryCnt As Integer = 0
        Dim found_SN As Boolean = False
        Dim SN As String = UUT("SN").Text

        ComPort = "UNKNOWN"

        If Not FtdiDeviceCount(DeviceCount) Then
            Return False
        End If
        If Not DeviceCount > 0 Then
            Return False
        End If

        _failure_message = ""
        Dim startTime As DateTime = Now
        Dim ftdiDeviceList(DeviceCount) As FTDI.FT_DEVICE_INFO_NODE

        found_blank_SN = True
        While (Not found_SN And found_blank_SN And retryCnt < 3)
            ftstatus = ftdi_device.GetDeviceList(ftdiDeviceList)
            If Not ftstatus = FTDI.FT_STATUS.FT_OK Then
                _failure_message = "Error getting FTDI device list:  " + ftstatus.ToString
                Return False
            End If

            found_blank_SN = False
            If retryCnt < 0 Then
                CommonLib.Delay(10)
                Form1.AppendText("RETRY", UUT:=UUT)
            End If
            retryCnt += 1

            For i = 0 To DeviceCount - 1
                Application.DoEvents()
                Form1.AppendText("device " + i.ToString + " LocID = " + String.Format("{0:x}", ftdiDeviceList(i).LocId), True, UUT:=UUT)
                Form1.AppendText("Checking device " + i.ToString + ", SN=" + ftdiDeviceList(i).SerialNumber + vbCr, True, UUT:=UUT)
                Form1.AppendText("description = " + ftdiDeviceList(i).Description.ToString, UUT:=UUT)
                If ftdiDeviceList(i).SerialNumber = "" Then
                    Form1.AppendText("SN not programmed in this device?", UUT:=UUT)
                    found_blank_SN = True
                End If
                If ftdiDeviceList(i).SerialNumber = SN Then
                    found_SN = True
                    Try
                        ftstatus = ftdi_device.OpenBySerialNumber(ftdiDeviceList(i).SerialNumber)
                    Catch ex As Exception
                        _failure_message = "Exception opening FTDI device:  " + ex.ToString
                        Return False
                    End Try
                    If Not ftstatus = FTDI.FT_STATUS.FT_OK Then
                        _failure_message = "Error opening FTDI device:  " + ftstatus.ToString
                        Return False
                    End If
                    ftstatus = ftdi_device.GetCOMPort(ComPort)
                    If Not ftstatus = FTDI.FT_STATUS.FT_OK Then
                        _failure_message = "Error getting FTDI comport:  " + ftstatus.ToString
                        Return False
                    End If
                    ftstatus = ftdi_device.Close()
                    If Not ftstatus = FTDI.FT_STATUS.FT_OK Then
                        _failure_message = "Error closing FTDI device  " + ftstatus.ToString
                        Return False
                    End If
                    Form1.AppendText("resetting " + ftdiDeviceList(i).SerialNumber, UUT:=UUT)
                    ftdi_device.ResetPort()
                    Form1.AppendText("cycling " + ftdiDeviceList(i).SerialNumber, UUT:=UUT)
                    ftdi_device.OpenByIndex(i)
                    ftdi_device.CyclePort()
                    ftdi_device.Close()

                    'For j = 0 To DeviceCount - 1
                    '    'If j = i Then Continue For
                    '    Form1.AppendText("cycling " + ftdiDeviceList(j).SerialNumber, UUT_Result:=UUT("RESULT"))
                    '    ftdi_device.OpenByIndex(j)
                    '    ftdi_device.CyclePort()
                    '    ftdi_device.Close()
                    'Next
                    Return True
                End If
            Next
        End While

        For i = 0 To DeviceCount - 1
            Form1.AppendText("device " + i.ToString + " LocID = " + String.Format("{0:x}", ftdiDeviceList(i).LocId), True, UUT:=UUT)
            Form1.AppendText("Checking device " + i.ToString + ", SN=" + ftdiDeviceList(i).SerialNumber + vbCr, True, UUT:=UUT)
            Form1.AppendText("description = " + ftdiDeviceList(i).Description.ToString, UUT:=UUT)
            If ftdiDeviceList(i).SerialNumber = "" Then
                found_blank_SN = True
            End If
            If ftdiDeviceList(i).SerialNumber = SN Then
                found_SN = True
                Try
                    ftstatus = ftdi_device.OpenBySerialNumber(ftdiDeviceList(i).SerialNumber)
                Catch ex As Exception
                    _failure_message = "Exception opening FTDI device:  " + ex.ToString
                    Return False
                End Try
                If Not ftstatus = FTDI.FT_STATUS.FT_OK Then
                    _failure_message = "Error opening FTDI device:  " + ftstatus.ToString
                    Return False
                End If
                ftstatus = ftdi_device.GetCOMPort(ComPort)
                If Not ftstatus = FTDI.FT_STATUS.FT_OK Then
                    _failure_message = "Error getting FTDI comport:  " + ftstatus.ToString
                    Return False
                End If
                ftstatus = ftdi_device.Close()
                If Not ftstatus = FTDI.FT_STATUS.FT_OK Then
                    _failure_message = "Error closing FTDI device  " + ftstatus.ToString
                    Return False
                End If
                Form1.AppendText("resetting " + ftdiDeviceList(i).SerialNumber, UUT:=UUT)
                ftdi_device.ResetPort()

                For j = 0 To DeviceCount - 1
                    'If j = i Then Continue For
                    Form1.AppendText("cycling " + ftdiDeviceList(j).SerialNumber, UUT:=UUT)
                    ftdi_device.OpenByIndex(j)
                    ftdi_device.CyclePort()
                    ftdi_device.Close()
                Next
                Return True
            End If
        Next


        Return False
    End Function

    Function SetSN(ByVal ftdi_device As FTDI, SN As String) As Boolean
        Dim myEEData As FTDI.FT232R_EEPROM_STRUCTURE
        Dim ftstatus As FTDI.FT_STATUS

        _failure_message = ""
        If Not ftdi_device.IsOpen Then
            _failure_message = "FTDI device is not open"
            Return False
        End If

        myEEData = New FTDI.FT232R_EEPROM_STRUCTURE()
        Try
            ftstatus = ftdi_device.ReadFT232REEPROM(myEEData)
        Catch ex As Exception
            _failure_message = "Exception reading FTDI eeprom:  " + ex.ToString
            Return False
        End Try

        If Not ftstatus = FTDI.FT_STATUS.FT_OK Then
            _failure_message = "Error reading FTDI eeprom:  " + ftstatus.ToString
            Return False
        End If

        myEEData.SerialNumber = SN
        Try
            ftstatus = ftdi_device.WriteFT232REEPROM(myEEData)
        Catch ex As Exception
            _failure_message = "Exception writing FTDI eeprom:  " + ex.ToString
            Return (False)
        End Try

        If Not ftstatus = FTDI.FT_STATUS.FT_OK Then
            _failure_message = "Error writing FTDI eeprom:  " + ftstatus.ToString
            Return False
        End If

        Return True
    End Function

    Function CycleDevice(ByVal ftdi_device As FTDI, LocId As UInteger) As Boolean
        Dim ftstatus As FTDI.FT_STATUS
        Dim startTime As DateTime = Now

        _failure_message = ""
        Try
            ftstatus = ftdi_device.CyclePort()
        Catch ex As Exception
            _failure_message = "Exception cycling FTDI device:  " + ex.ToString
            Return False
        End Try
        If Not ftstatus = FTDI.FT_STATUS.FT_OK Then
            _failure_message = "Error cycling FTDI device:  " + ftstatus.ToString
            Return False
        End If
        Do
            ftstatus = ftdi_device.OpenByLocation(LocId)
            System.Threading.Thread.Sleep(1000)
        Loop Until (ftstatus = FTDI.FT_STATUS.FT_OK) Or (Now.Subtract(startTime).TotalSeconds > 30)

        If Not ftstatus = FTDI.FT_STATUS.FT_OK Then
            _failure_message = "Timeout resetting FTDI device "
            Return False
        End If
        Return True
    End Function
End Class
