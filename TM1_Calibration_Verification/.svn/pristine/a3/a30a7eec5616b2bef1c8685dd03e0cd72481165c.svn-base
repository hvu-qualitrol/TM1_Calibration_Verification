Imports System.IO.Ports
Imports System.Text.RegularExpressions


Partial Class Tests

    Public Shared Function Test_INITIAL() As Boolean
        Dim Success As Boolean = True

        ' Get LinkInfo
        If Not GetLinkInfo() Then Success = False

        ' Find and login to UUT's
        If Not LoginUuts() Then Success = False

        ' Check firmware versions
        If Not CheckFirmwareVersions() Then Success = False

        '' For debug only
        'GetData()
        'Return True

        ' Quit If report generating only
        If (Form1.PartialRun) Then Return Success
 
        ' Force H2Scan to Lab mode
        If Not ForceH2ScanToLabMode() Then Success = False

        ' Set and verify date
        If Not SetAndVerifyDate() Then Success = False

        Return Success

    End Function

    ''' <summary>
    ''' This is to check out the capability fo parsing records from differnt hardware versions
    ''' whose record length and fileds are different. This is intented for debug only
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function GetData() As Boolean
        Dim Success As Boolean
        Dim SF As New SerialFunctions
        Dim h2scan As New H2SCAN_debug
        Dim Tm8RecData As DataTable
        Dim TM8 As New TM8
        Dim csv_filepath As String
        Dim TimeStamp As String = Format(Date.UtcNow, "yyyyMMddHHmmss")
        Dim LastRowDateTime As DateTime
        Dim ts_3h As New TimeSpan(3, 0, 0)
        Dim WhereClause As String
        Dim TM8_gas_ppm As Double
        Dim AllFailed As Boolean = True
        Dim startTime As DateTime = Now
        Dim u1_time As Integer = 28
        Dim dt As DataTable

        Log.WriteLine("Getting TM8 record data")
        'Get TM8 rec Data
        If Not TM8.GetRecdata(Tm8RecData) Then
            Form1.AppendText(TM8.ErrorMsg)
            Return False
        End If

        For Each UUT In UUTs
            Application.DoEvents()
            If UUT("SN").Text = "" Or UUT("FAILED") Then Continue For

            Dim Tm1RecFields() As String = Tm1Rev0RecFields
            If UUT("TM1 Hardware Version") <> Products(Product)("hardware version 0") Then
                Form1.AppendText("Use TM1Rev2RecFields", UUT:=UUT)
                Tm1RecFields = Tm1Rev2RecFields
            End If
            Try
                Log.WriteLine("Creating datatable for " + UUT("SN").Text)
                If Not CommonLib.CreateDataTable(UUT("DT"), Tm1RecFields) Then
                    UUT("FAILED") = True
                    UUT("TAB").ImageIndex = StatusColor.FAILED
                    Success = False
                    Continue For
                Else
                    Log.WriteLine("Table Created")
                End If

                Log.WriteLine("Getting TM1 data for " + UUT("SN").Text)
                Dim recordCount = 112 ' It used to be  Timeout_U1 * 4
                If Not h2scan.GetRecData(UUT, recordCount, UUT("DT")) Then
                    Form1.AppendText(h2scan.ErrorMsg, UUT:=UUT)
                    UUT("FAILED") = True
                    UUT("TAB").ImageIndex = StatusColor.FAILED
                    Success = False
                    Continue For
                Else
                    Log.WriteLine("Data retrieved")
                End If

                ' add TM8 data
                Log.WriteLine("Combining TM1 and TM8 datatable for " + UUT("SN").Text)
                If Not TM8.CombineTm8_Tm1_Data(UUT("DT"), Tm8RecData) Then
                    Form1.AppendText(h2scan.ErrorMsg, UUT:=UUT)
                    UUT("FAILED") = True
                    UUT("TAB").ImageIndex = StatusColor.FAILED
                    Success = False
                    Continue For
                Else
                    Log.WriteLine("Data combining finished")
                End If

                Log.WriteLine("Binding to the GridView")
                UUT("GV").DataSource() = UUT("DT")
                'UUT("GV").FirstDisplayedCell = UUT("GV").Rows(UUT("GV").Rows.Count - 1).Cells(0)
                Log.WriteLine("Walking the columns and adding to the grid view")
                For Each col In UUT("GV").Columns
                    Log.WriteLine("Adding column : " + col.Name)
                    If Not (col.Name = "Timestamp" Or col.Name = "H2_OIl.PPM" Or col.Name = "H2.PPM" Or
                            col.Name = "TM8_ppm" Or col.Name = "TM8_gas_ppm") Then
                        UUT("GV").Columns(col.Name).visible = False
                    End If
                Next
                Log.WriteLine("Writing results to the csv file for unit " + UUT("SN").Text)
                csv_filepath = ReportDir + "FINAL_TEST" + "\" + UUT("SN").Text + "\U1." + TimeStamp + ".csv"
                If Not CommonLib.ExportDataTableToCSV(UUT("DT"), csv_filepath) Then
                    Form1.AppendText("Problem creating csv file " + csv_filepath, UUT:=UUT)
                    UUT("FAILED") = True
                    UUT("TAB").ImageIndex = StatusColor.FAILED
                    Success = False
                    Continue For
                Else
                    Log.WriteLine("File Created")
                End If
            Catch ex As Exception
                Form1.AppendText(UUT("SN").Text + ": Test_U1() caught" + ex.ToString, UUT:=UUT)
                UUT("FAILED") = True
                UUT("TAB").ImageIndex = StatusColor.FAILED
                Success = False
                Continue For
            End Try

            Try
                LastRowDateTime = UUT("DT").Rows(0)("Timestamp")
                WhereClause = "TM8_RecTime > '" + LastRowDateTime.Subtract(ts_3h).ToString + "'"
                Log.WriteLine("LastRowDateTime Value is " + LastRowDateTime.ToString)
            Catch ex As Exception
                Log.WriteLine("Execption throwning during calculating the TM8 record times.")
                Log.WriteLine("Exception message is " + ex.Message)
                Log.WriteLine("Inner Exception is " + ex.InnerException.Message)
            End Try

            Try
                TM8_gas_ppm = UUT("DT").Compute("AVG(TM8_gas_ppm)", WhereClause)
                Log.WriteLine("Average TM8 Gas Value : " + TM8_gas_ppm.ToString)
            Catch ex As Exception
                Log.WriteLine("Exception in calculating the average : " + ex.Message)
                Log.WriteLine("Exception message is " + ex.Message)
                Log.WriteLine("Inner Exception is " + ex.InnerException.Message)
                Form1.AppendText("Error calculating average TM8 ppm for prev 3 hours", UUT:=UUT)
                Form1.AppendText(ex.ToString)
                UUT("FAILED") = True
                UUT("TAB").ImageIndex = StatusColor.FAILED
                Success = False
                Continue For
            End Try
            'TM8_gas_ppm = UUT("DT").Rows(0)("TM8_gas_ppm")
            Form1.AppendText("Average TM8 gas ppm to be used for calibration:  " + TM8_gas_ppm.ToString, UUT:=UUT)
            Log.WriteLine("Checking that the average PPM value is in spec")
            If TM8_gas_ppm < 800 Or TM8_gas_ppm > 1200 Then
                Form1.AppendText("Expected TM8 gas ppm to be between 800 & 1200", UUT:=UUT)
                UUT("FAILED") = True
                UUT("TAB").ImageIndex = StatusColor.FAILED
                Success = False
                Continue For
            End If

            Try
                Log.WriteLine("Opening the H2 Scan device")
                If Not h2scan.Open(UUT) Then
                    Form1.AppendText(h2scan.ErrorMsg, UUT:=UUT)
                    UUT("FAILED") = True
                    UUT("TAB").ImageIndex = StatusColor.FAILED
                    Success = False
                    Continue For
                Else
                    Log.WriteLine("Done opening")
                End If

                Log.WriteLine("Setting the U1 cal value")
                If Not h2scan.U1(UUT, TM8_gas_ppm) Then
                    Form1.AppendText(h2scan.ErrorMsg, UUT:=UUT)
                    UUT("FAILED") = True
                    UUT("TAB").ImageIndex = StatusColor.FAILED
                    Success = False
                    h2scan.Close(UUT)
                    Continue For
                End If

                If Not h2scan.Close(UUT) Then
                    Form1.AppendText(h2scan.ErrorMsg, UUT:=UUT)
                    UUT("FAILED") = True
                    UUT("TAB").ImageIndex = StatusColor.FAILED
                    Success = False
                    Continue For
                End If
                AllFailed = False
            Catch ex As Exception
                Form1.AppendText(UUT("SN").Text + ": Test_U1() caught" + ex.ToString, UUT:=UUT)
                UUT("FAILED") = True
                UUT("TAB").ImageIndex = StatusColor.FAILED
                Success = False
                Continue For
            End Try
        Next

        Return True
    End Function

    Private Shared Function GetLinkInfo() As Boolean
        Dim LinkInfo As Hashtable
        Dim DB As New DB
        Dim Success As Boolean = True

        For Each UUT In UUTs
            Application.DoEvents()
            If UUT("SN").Text = "" Then Continue For
            Try
                If Not DB.GetLinkData(UUT("SN").Text, AssemblyType.TM1, LinkInfo) Then
                    Form1.AppendText("Problem getting SN link data for " + UUT("SN").Text, UUT:=UUT)
                    UUT("FAILED") = True
                    UUT("TAB").ImageIndex = StatusColor.FAILED
                    Success = False
                Else
                    UUT("LI") = LinkInfo
                    For Each k In LinkInfo.Keys
                        Form1.AppendText(k + "=" + LinkInfo(k), UUT:=UUT)
                    Next
                    If Not DB.CheckPrevTests(UUT) Then
                        Form1.AppendText(DB.ErrorMsg, UUT:=UUT)
                        UUT("FAILED") = True
                        UUT("TAB").ImageIndex = StatusColor.FAILED
                        Return False
                    End If
                End If
            Catch ex As Exception
                Form1.AppendText(UUT("SN").Text + ": GetLinkInfo() caught" + ex.ToString, UUT:=UUT)
                UUT("FAILED") = True
                UUT("TAB").ImageIndex = StatusColor.FAILED
                Success = False
            End Try
        Next

        Return Success
    End Function

    Private Shared Function LoginUuts() As Boolean
        Dim Success As Boolean = True
        Dim ftdi_device As New FT232R
        Dim SF As New SerialFunctions
        Dim SN As String
        Dim ComPort As String = ""
        Dim results As ReturnResults

        For Each UUT In UUTs
            Application.DoEvents()
            If UUT("SN").Text = "" Or UUT("FAILED") Then Continue For
            Try
                SN = UUT("SN").Text
                If Not ftdi_device.FindComportForSN(UUT, ComPort) Then
                    Form1.AppendText("Could not find comport for " + SN, UUT:=UUT)
                    UUT("FAILED") = True
                    UUT("TAB").ImageIndex = StatusColor.FAILED
                    Success = False
                Else
                    Form1.AppendText(UUT("SN").Text, UUT:=UUT)
                    UUT("COM").Text = ComPort
                    UUT("SP") = New SerialPort(ComPort, 115200, 0, 8, 1)
                    UUT("SP").Handshake = Handshake.RequestToSend
                    results = SF.Connect(UUT)
                    If Not results.PassFail Then
                        UUT("FAILED") = True
                        UUT("TAB").ImageIndex = StatusColor.FAILED
                        Success = False
                    End If
                End If
            Catch ex As Exception
                Form1.AppendText(UUT("SN").Text + ": LoginUuts() caught" + ex.ToString, UUT:=UUT)
                UUT("FAILED") = True
                UUT("TAB").ImageIndex = StatusColor.FAILED
                Success = False
            End Try
        Next

        Return Success
    End Function

    Private Shared Function CheckFirmwareVersions() As Boolean
        Dim Success As Boolean = True
        Dim VersionInfo As Hashtable = New Hashtable
        Dim FW_Version As String = "UNKNOWN"
        Dim hardwareVersion As String = "UNKNOWN"
        Dim results As ReturnResults
        Dim SF As New SerialFunctions
        Dim T As New TM1

        For Each UUT In UUTs
            Application.DoEvents()
            If UUT("SN").Text = "" Or UUT("FAILED") Then Continue For

            Try
                results = SF.Connect(UUT)
                If Not results.PassFail Then
                    UUT("FAILED") = True
                    UUT("TAB").ImageIndex = StatusColor.FAILED
                    Success = False
                    Continue For
                End If

                If Not T.GetVersionInfo(UUT, VersionInfo) Then
                    UUT("FAILED") = True
                    UUT("TAB").ImageIndex = StatusColor.FAILED
                    Success = False
                    Continue For
                End If

                FW_Version = Regex.Split(VersionInfo("firmware version"), "\s+")(1)
                UUT("TM1 Firmware Version") = FW_Version
                UUT("TM1 Hardware Version") = VersionInfo("hardware version")
                If (VersionInfo("hardware version") <> Products(Product)("hardware version 0") And
                    VersionInfo("hardware version") <> Products(Product)("hardware version 1") And
                    VersionInfo("hardware version") <> Products(Product)("hardware version 2")) Then
                    Form1.AppendText("Expected 'hardware version' = " + Products(Product)("hardware version 0") +
                                     " or " + Products(Product)("hardware version 1") +
                                     " or " + Products(Product)("hardware version 2"), True)
                    UUT("FAILED") = True
                    UUT("TAB").ImageIndex = StatusColor.FAILED
                    Success = False
                    Continue For
                End If

                UUT("embedded moisture") = VersionInfo("embedded moisture")
                'If VersionInfo("hardware version") <> Products(Product)("hardware version 0") Then
                '    If Not T.SetHwRev2Config(UUT) Then
                '        UUT("FAILED") = True
                '        UUT("TAB").ImageIndex = StatusColor.FAILED
                '        Success = False
                '        Continue For
                '    End If
                'End If

                If Not VersionInfo("sensor firmware version") = Products(Product)("sensor firmware version") Then
                    Form1.AppendText("Expected 'sensor firmware version' = " + Products(Product)("sensor firmware version"), True, UUT:=UUT)
                    UUT("FAILED") = True
                    UUT("TAB").ImageIndex = StatusColor.FAILED
                    Success = False
                    Continue For
                End If

            Catch ex As Exception
                Form1.AppendText(UUT("SN").Text + ": CheckFirmwareVersions() caught" + ex.ToString, UUT:=UUT)
                UUT("FAILED") = True
                UUT("TAB").ImageIndex = StatusColor.FAILED
                Success = False
            End Try
        Next

        Return Success
    End Function

    Private Shared Function ForceH2ScanToLabMode() As Boolean
        Dim Success As Boolean = True
        Dim SF As New SerialFunctions
        Dim h2scan As New H2SCAN_debug
        Dim results As ReturnResults

        For Each UUT In UUTs
            Application.DoEvents()
            If UUT("SN").Text = "" Or UUT("FAILED") Then Continue For

            Try
                results = SF.Connect(UUT)
                If Not results.PassFail Then
                    UUT("FAILED") = True
                    UUT("TAB").ImageIndex = StatusColor.FAILED
                    Success = False
                    Continue For
                End If

                If Not h2scan.SetSensorMode(UUT, H2SCAN_debug.H2SCAN_OP_MODE.LAB) Then
                    h2scan.Close(UUT, True)
                    Form1.AppendText(h2scan.ErrorMsg, UUT:=UUT)
                    UUT("FAILED") = True
                    UUT("TAB").ImageIndex = StatusColor.FAILED
                    Success = False
                    Continue For
                End If
            Catch ex As Exception
                Form1.AppendText(UUT("SN").Text + ": ForceH2ScanToLabMode() caught" + ex.ToString, UUT:=UUT)
                UUT("FAILED") = True
                UUT("TAB").ImageIndex = StatusColor.FAILED
                Success = False
            End Try
        Next

        Return Success
    End Function

    Private Shared Function SetAndVerifyDate() As Boolean
        Dim Success As Boolean = True
        Dim SF As New SerialFunctions
        Dim T As New TM1
        Dim results As ReturnResults

        For Each UUT In UUTs
            Application.DoEvents()
            If UUT("SN").Text = "" Or UUT("FAILED") Then Continue For

            Try
                results = SF.Connect(UUT)
                If Not results.PassFail Then
                    UUT("FAILED") = True
                    UUT("TAB").ImageIndex = StatusColor.FAILED
                    Success = False
                    Continue For
                End If

                If Not T.SetVerifyDate(UUT) Then
                    UUT("FAILED") = True
                    UUT("TAB").ImageIndex = StatusColor.FAILED
                    Success = False
                End If
            Catch ex As Exception
                Form1.AppendText(UUT("SN").Text + ": SetAndVerifyDate() caught" + ex.ToString, UUT:=UUT)
                UUT("FAILED") = True
                UUT("TAB").ImageIndex = StatusColor.FAILED
                Success = False
            End Try
        Next

        Return Success
    End Function

    Public Shared Function Test_INITIAL0() As Boolean
        Dim ftdi_device As New FT232R
        Dim SN As String
        Dim ComPort As String
        Dim Success As Boolean
        Dim SF As New SerialFunctions
        Dim results As ReturnResults
        Dim T As New TM1
        Dim VersionInfo As Hashtable
        Dim FW_Version As String = "UNKNOWN"
        Dim h2scan As New H2SCAN_debug
        Dim op_mode As H2SCAN_debug.H2SCAN_OP_MODE
        Dim DB As New DB
        Dim LinkInfo As Hashtable

        Success = True

        ' Get LinkInfo
        For Each UUT In UUTs
            Application.DoEvents()
            If UUT("SN").Text = "" Then Continue For
            Try
                If Not DB.GetLinkData(UUT("SN").Text, AssemblyType.TM1, LinkInfo) Then
                    Form1.AppendText("Problem getting SN link data for " + UUT("SN").Text, UUT:=UUT)
                    UUT("FAILED") = True
                    UUT("TAB").ImageIndex = StatusColor.FAILED
                    Success = False
                Else
                    UUT("LI") = LinkInfo
                    For Each k In LinkInfo.Keys
                        Form1.AppendText(k + "=" + LinkInfo(k), UUT:=UUT)
                    Next
                    If Not DB.CheckPrevTests(UUT) Then
                        Form1.AppendText(DB.ErrorMsg, UUT:=UUT)
                        UUT("FAILED") = True
                        UUT("TAB").ImageIndex = StatusColor.FAILED
                        Return False
                    End If
                End If
            Catch ex As Exception
                Form1.AppendText(UUT("SN").Text + ": Test_Initial() caught" + ex.ToString, UUT:=UUT)
                UUT("FAILED") = True
                UUT("TAB").ImageIndex = StatusColor.FAILED
                Success = False
            End Try
        Next

        ' Find and login to UUT's
        For Each UUT In UUTs
            Application.DoEvents()
            If UUT("SN").Text = "" Or UUT("FAILED") Then Continue For
            Try
                SN = UUT("SN").Text
                If Not ftdi_device.FindComportForSN(UUT, ComPort) Then
                    Form1.AppendText("Could not find comport for " + SN, UUT:=UUT)
                    UUT("FAILED") = True
                    UUT("TAB").ImageIndex = StatusColor.FAILED
                    Success = False
                Else
                    Form1.AppendText(UUT("SN").Text, UUT:=UUT)
                    UUT("COM").Text = ComPort
                    UUT("SP") = New SerialPort(ComPort, 115200, 0, 8, 1)
                    UUT("SP").Handshake = Handshake.RequestToSend
                    results = SF.Connect(UUT)
                    If Not results.PassFail Then
                        UUT("FAILED") = True
                        UUT("TAB").ImageIndex = StatusColor.FAILED
                        Success = False
                    End If
                End If
            Catch ex As Exception
                Form1.AppendText(UUT("SN").Text + ": Test_Initial() caught" + ex.ToString, UUT:=UUT)
                UUT("FAILED") = True
                UUT("TAB").ImageIndex = StatusColor.FAILED
                Success = False
            End Try
        Next

        ' Check H2SCAN Operating Mode
        For Each UUT In UUTs
            Application.DoEvents()
            If UUT("SN").Text = "" Or UUT("FAILED") Then Continue For
            Try
                If Not (h2scan.GetOperatingMode(UUT, op_mode)) Then
                    Form1.AppendText(h2scan.ErrorMsg, UUT:=UUT)
                    UUT("FAILED") = True
                    UUT("TAB").ImageIndex = StatusColor.FAILED
                    Success = False
                    Continue For
                End If
                Form1.AppendText("H2SCAN Operating Mode = " + op_mode.ToString, UUT:=UUT)
                If Not op_mode = H2SCAN_debug.H2SCAN_OP_MODE.LAB Then
                    Form1.AppendText("Expected sensor to be in LAB mode", UUT:=UUT)
                    UUT("FAILED") = True
                    UUT("TAB").ImageIndex = StatusColor.FAILED
                    Success = False
                End If
            Catch ex As Exception
                Form1.AppendText(UUT("SN").Text + ": Test_Initial() caught" + ex.ToString, UUT:=UUT)
                UUT("FAILED") = True
                UUT("TAB").ImageIndex = StatusColor.FAILED
                Success = False
            End Try
        Next

        ' check versions
        For Each UUT In UUTs
            Application.DoEvents()
            If UUT("SN").Text = "" Or UUT("FAILED") Then Continue For

            Try
                If Not T.GetVersionInfo(UUT, VersionInfo) Then
                    UUT("FAILED") = True
                    UUT("TAB").ImageIndex = StatusColor.FAILED
                    Success = False
                    Continue For
                End If

                FW_Version = Regex.Split(VersionInfo("firmware version"), "\s+")(1)

                ' TO DO : Disabled the firmware version to do testing for the next verions.

                'If Not FW_Version = Products(Product)("FW_VERSION") Then
                '    Form1.AppendText("Expected 'firmware version' = " + Products(Product)("FW_VERSION"), True, UUT:=UUT)
                '    UUT("FAILED") = True
                '    UUT("TAB").ImageIndex = StatusColor.FAILED
                '    Success = False
                '    Continue For
                'End If

                UUT("TM1 Firmware Version") = FW_Version
                If Not VersionInfo("hardware version") = Products(Product)("hardware version") Then
                    Form1.AppendText("Expected 'hardware version' = " + Products(Product)("hardware version"), True, UUT:=UUT)
                    UUT("FAILED") = True
                    UUT("TAB").ImageIndex = StatusColor.FAILED
                    Success = False
                    Continue For
                End If
                'If Not VersionInfo("assembly version") = Products(Product)("assembly version") Then
                '    Form1.AppendText("Expected 'assembly version' = " + Products(Product)("assembly version"), True, UUT:=UUT)
                '    UUT("FAILED") = True
                '    UUT("TAB").ImageIndex = StatusColor.FAILED
                '    Success = False
                '    Continue For
                'End If
                If Not VersionInfo("sensor firmware version") = Products(Product)("sensor firmware version") Then
                    Form1.AppendText("Expected 'sensor firmware version' = " + Products(Product)("sensor firmware version"), True, UUT:=UUT)
                    UUT("FAILED") = True
                    UUT("TAB").ImageIndex = StatusColor.FAILED
                    Success = False
                    Continue For
                End If
            Catch ex As Exception
                Form1.AppendText(UUT("SN").Text + ": Test_Initial() caught" + ex.ToString, UUT:=UUT)
                UUT("FAILED") = True
                UUT("TAB").ImageIndex = StatusColor.FAILED
                Success = False
            End Try
        Next

        ' Check H2SCAN Operating Mode
        For Each UUT In UUTs
            Application.DoEvents()
            If UUT("SN").Text = "" Or UUT("FAILED") Then Continue For

            Try
                If Not (h2scan.GetOperatingMode(UUT, op_mode)) Then
                    Form1.AppendText(h2scan.ErrorMsg, UUT:=UUT)
                    UUT("FAILED") = True
                    UUT("TAB").ImageIndex = StatusColor.FAILED
                    Success = False
                    Continue For
                End If
                Form1.AppendText("H2SCAN Operating Mode = " + op_mode.ToString, UUT:=UUT)
                If Not op_mode = H2SCAN_debug.H2SCAN_OP_MODE.LAB Then
                    Form1.AppendText("Expected sensor to be in LAB mode", UUT:=UUT)
                    UUT("FAILED") = True
                    UUT("TAB").ImageIndex = StatusColor.FAILED
                    Success = False
                End If
            Catch ex As Exception
                Form1.AppendText(UUT("SN").Text + ": Test_Initial() caught" + ex.ToString, UUT:=UUT)
                UUT("FAILED") = True
                UUT("TAB").ImageIndex = StatusColor.FAILED
                Success = False
            End Try
        Next

        ' Set and verify date
        For Each UUT In UUTs
            Application.DoEvents()
            If UUT("SN").Text = "" Or UUT("FAILED") Then Continue For

            Try
                If Not T.SetVerifyDate(UUT) Then
                    UUT("FAILED") = True
                    UUT("TAB").ImageIndex = StatusColor.FAILED
                    Success = False
                End If
            Catch ex As Exception
                Form1.AppendText(UUT("SN").Text + ": Test_Initial() caught" + ex.ToString, UUT:=UUT)
                UUT("FAILED") = True
                UUT("TAB").ImageIndex = StatusColor.FAILED
                Success = False
            End Try
        Next

        ' Check for UUT's stopped by stop button
        For Each UUT In UUTs
            If UUT("FAILED") Then Success = False
        Next

        Return Success
    End Function

    Public Shared Function Test_INITIAL_Debug() As Boolean
        Dim ftdi_device As New FT232R
        Dim SN As String
        Dim ComPort As String
        Dim Success As Boolean
        Dim SF As New SerialFunctions
        Dim results As ReturnResults
        Dim T As New TM1
        Dim VersionInfo As Hashtable
        Dim FW_Version As String = "UNKNOWN"
        Dim h2scan As New H2SCAN_debug
        Dim op_mode As H2SCAN_debug.H2SCAN_OP_MODE
        Dim DB As New DB
        Dim LinkInfo As Hashtable

        Success = True

        '' Get LinkInfo
        'For Each UUT In UUTs
        '    Application.DoEvents()
        '    If UUT("SN").Text = "" Then Continue For
        '    Try
        '        If Not DB.GetLinkData(UUT("SN").Text, AssemblyType.TM1, LinkInfo) Then
        '            Form1.AppendText("Problem getting SN link data for " + UUT("SN").Text, UUT:=UUT)
        '            UUT("FAILED") = True
        '            UUT("TAB").ImageIndex = StatusColor.FAILED
        '            Success = False
        '        Else
        '            UUT("LI") = LinkInfo
        '            For Each k In LinkInfo.Keys
        '                Form1.AppendText(k + "=" + LinkInfo(k), UUT:=UUT)
        '            Next
        '            If Not DB.CheckPrevTests(UUT) Then
        '                Form1.AppendText(DB.ErrorMsg, UUT:=UUT)
        '                UUT("FAILED") = True
        '                UUT("TAB").ImageIndex = StatusColor.FAILED
        '                Return False
        '            End If
        '        End If
        '    Catch ex As Exception
        '        Form1.AppendText(UUT("SN").Text + ": Test_Initial() caught" + ex.ToString, UUT:=UUT)
        '        UUT("FAILED") = True
        '        UUT("TAB").ImageIndex = StatusColor.FAILED
        '        Success = False
        '    End Try
        'Next

        ' Find and login to UUT's
        For Each UUT In UUTs
            Application.DoEvents()
            If UUT("SN").Text = "" Or UUT("FAILED") Then Continue For
            Try
                SN = UUT("SN").Text
                If Not ftdi_device.FindComportForSN(UUT, ComPort) Then
                    Form1.AppendText("Could not find comport for " + SN, UUT:=UUT)
                    UUT("FAILED") = True
                    UUT("TAB").ImageIndex = StatusColor.FAILED
                    Success = False
                Else
                    Form1.AppendText(UUT("SN").Text, UUT:=UUT)
                    UUT("COM").Text = ComPort
                    UUT("SP") = New SerialPort(ComPort, 115200, 0, 8, 1)
                    UUT("SP").Handshake = Handshake.RequestToSend
                    results = SF.Connect(UUT)
                    If Not results.PassFail Then
                        UUT("FAILED") = True
                        UUT("TAB").ImageIndex = StatusColor.FAILED
                        Success = False
                    End If

                    ' For debug on 02/27/2015: Change sensor mode
                    If Not h2scan.SetSensorMode(UUT, H2SCAN_debug.H2SCAN_OP_MODE.LAB) Then
                        Form1.AppendText(h2scan.Results, UUT:=UUT, LogToResults:=False)
                        Form1.AppendText("Problem setting " + UUT("SN").Text + " to field operating mode", UUT:=UUT, LogToResults:=False)
                        Form1.AppendText(h2scan.ErrorMsg, UUT:=UUT, LogToResults:=False)
                        h2scan.Close(UUT, True)
                        UUT("FAILED") = True
                        UUT("TAB").ImageIndex = StatusColor.FAILED
                        Success = False
                    End If
                    Form1.AppendText(h2scan.Results, UUT:=UUT, LogToResults:=False)
                End If
            Catch ex As Exception
                Form1.AppendText(UUT("SN").Text + ": Test_Initial() caught" + ex.ToString, UUT:=UUT)
                UUT("FAILED") = True
                UUT("TAB").ImageIndex = StatusColor.FAILED
                Success = False
            End Try
        Next

        Return False
    End Function
End Class