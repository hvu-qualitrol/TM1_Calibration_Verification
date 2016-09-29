Partial Class Tests
    ' Capture last 24 hours of TM8 and TM1 data
    Public Shared Function Test_U1() As Boolean
        Dim Success As Boolean
        Dim results As ReturnResults
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

        Log.WriteLine("Entering U1 Class")

        While Now.Subtract(startTime).TotalHours < Timeout_U1
            If Not TM8.GetRecdata(Tm8RecData) Then
                Form1.AppendText(TM8.ErrorMsg)
                Log.WriteLine("Error : " & TM8.ErrorMsg)
                Return False
            End If
            TM8_gas_ppm = Tm8RecData.Rows(0)("gas_ppm")
            Form1.TM8_PPM.Text = TM8_gas_ppm.ToString
            Log.WriteLine("TM8 Gas PPM value is : " & TM8_gas_ppm.ToString)
            If TM8_gas_ppm < 800 Or TM8_gas_ppm > 1200 Then
                TM8_gas_in_spec = False
            Else
                TM8_gas_in_spec = True
            End If

            If Not TM8_gas_in_spec Then
                Form1.AppendText("TM8 gas ppm out of spec, expected between 800 & 1200")
                Return False
            End If
            Form1.TimeoutLabel.Text = Math.Round(Timeout_U1 - Now.Subtract(startTime).TotalHours, 1).ToString + "h"
            CommonLib.Delay(60 * 15)
        End While

        '' ***************************** HACK ****************************
        '' Adding 30 mins to the end of this cycle as per requested by Thomas.

        'Log.WriteLine("Staring 30 min. extra delay")
        ''CommonLib.Delay(60 * 30)

        ' ***************************************************************
        Log.WriteLine("Getting TM8 record data")
        'Get TM8 rec Data
        If Not TM8.GetRecdata(Tm8RecData) Then
            Form1.AppendText(TM8.ErrorMsg)
            Return False
        End If

        ' login to UUT's
        CommonLib.Delay(10)
        Log.WriteLine("Logging into the Units")
        Success = True
        For Each UUT In UUTs
            ' Skip failing units
            If UUT("SN").Text = "" Or UUT("FAILED") Then Continue For

            Dim Tm1RecFields() As String = Tm1Rev0RecFields
            If UUT("TM1 Hardware Version") <> Products(Product)("hardware version 0") Then
                Tm1RecFields = Tm1Rev2RecFields
            End If

            Try
                Log.WriteLine("Attempting login for " + UUT("SN").Text)
                results = SF.Connect(UUT)
                If Not results.PassFail Then
                    UUT("FAILED") = True
                    UUT("TAB").ImageIndex = StatusColor.FAILED
                    Success = False
                    Continue For
                Else
                    Log.WriteLine("Login success")
                End If
                CommonLib.Delay(1)

                Log.WriteLine("Getting TM1 record data")
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
                Dim recordCount = Timeout_U1 * 4
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

                LastRowDateTime = UUT("DT").Rows(0)("Timestamp")
                WhereClause = "TM8_RecTime > '" + LastRowDateTime.Subtract(ts_3h).ToString + "'"
                Log.WriteLine("LastRowDateTime Value is " + LastRowDateTime.ToString)

                TM8_gas_ppm = UUT("DT").Compute("AVG(TM8_gas_ppm)", WhereClause)
                Log.WriteLine("Average TM8 Gas Value : " + TM8_gas_ppm.ToString)

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

        If AllFailed Then
            Return False
        Else
            Return True
        End If
    End Function

    Public Shared Function Test_U10() As Boolean
        Dim Success As Boolean
        Dim results As ReturnResults
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

        Log.WriteLine("Entering U1 Class")

        While Now.Subtract(startTime).TotalHours < Timeout_U1
            If Not TM8.GetRecdata(Tm8RecData) Then
                Form1.AppendText(TM8.ErrorMsg)
                Log.WriteLine("Error : " & TM8.ErrorMsg)
                Return False
            End If
            TM8_gas_ppm = Tm8RecData.Rows(0)("gas_ppm")
            Form1.TM8_PPM.Text = TM8_gas_ppm.ToString
            Log.WriteLine("TM8 Gas PPM value is : " & TM8_gas_ppm.ToString)
            If TM8_gas_ppm < 800 Or TM8_gas_ppm > 1200 Then
                TM8_gas_in_spec = False
            Else
                TM8_gas_in_spec = True
            End If

            If Not TM8_gas_in_spec Then
                Form1.AppendText("TM8 gas ppm out of spec, expected between 800 & 1200")
                Return False
            End If
            Form1.TimeoutLabel.Text = Math.Round(Timeout_U1 - Now.Subtract(startTime).TotalHours, 1).ToString + "h"
            CommonLib.Delay(60 * 15)
        End While

        ' ***************************** HACK ****************************
        ' Adding 30 mins to the end of this cycle as per requested by Thomas.

        Log.WriteLine("Staring 30 min. extra delay")
        'CommonLib.Delay(60 * 30)

        ' ***************************************************************

        ' login to UUT's
        CommonLib.Delay(10)
        Log.WriteLine("Logging into the Units")
        Success = True
        For Each UUT In UUTs
            Log.WriteLine("Attempting login for " + UUT("SN").Text)
            If UUT("SN").Text = "" Or UUT("FAILED") Then Continue For

            Try
                results = SF.Connect(UUT)
                If Not results.PassFail Then
                    UUT("FAILED") = True
                    UUT("TAB").ImageIndex = StatusColor.FAILED
                    Success = False
                Else
                    Log.WriteLine("Login success")
                End If
                CommonLib.Delay(30)
            Catch ex As Exception
                Form1.AppendText(UUT("SN").Text + ": Test_U1() caught" + ex.ToString, UUT:=UUT)
                UUT("FAILED") = True
                UUT("TAB").ImageIndex = StatusColor.FAILED
                Success = False
            End Try
        Next

        Log.WriteLine("Getting TM8 record data")
        'Get TM8 rec Data
        If Not TM8.GetRecdata(Tm8RecData) Then
            Form1.AppendText(TM8.ErrorMsg)
            Return False
        End If

        Log.WriteLine("Getting TM1 record data")
        'Get TM1 rec data
        For Each UUT In UUTs
            CommonLib.Delay(1)
            If UUT("SN").Text = "" Or UUT("FAILED") Then Continue For

            Dim Tm1RecFields() As String = Tm1Rev0RecFields
            If UUT("TM1 Hardware Version") <> Products(Product)("hardware version 0") Then
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

        If AllFailed Then
            Return False
        Else
            Return True
        End If
        ' Return Success
    End Function
End Class