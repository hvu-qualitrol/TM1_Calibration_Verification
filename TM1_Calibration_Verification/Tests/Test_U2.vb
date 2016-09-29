Partial Class Tests
    Public Shared Function Test_U2() As Boolean
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
        Dim TM8_gas_ppm_valid As Boolean
        Dim AllFailed As Boolean = True
        Dim startTime As DateTime = Now
        Dim u2_time As Integer = 20

        While Now.Subtract(startTime).TotalHours < Timeout_U2
            If Not TM8.GetRecdata(Tm8RecData) Then
                Form1.AppendText(TM8.ErrorMsg)
                Return False
            End If
            TM8_gas_ppm = Tm8RecData.Rows(0)("gas_ppm")
            Form1.TM8_PPM.Text = TM8_gas_ppm.ToString
            If TM8_gas_ppm < 8500 Or TM8_gas_ppm > 11500 Then
                TM8_gas_in_spec = False
            Else
                TM8_gas_in_spec = True
            End If
            If Not TM8_gas_in_spec Then
                Form1.AppendText("TM8 gas ppm out of spec, expected between 8500 & 11500")
                Return False
            End If
            Form1.TimeoutLabel.Text = Math.Round(Timeout_U2 - Now.Subtract(startTime).TotalHours, 1).ToString + "h"
            CommonLib.Delay(60 * 15)
        End While

        ' ***************************** HACK ****************************
        ' Adding 30 mins to the end of this cycle as per requested by Thomas.
        ' This does not seem to help in reducing the negative spikes, so it is 
        ' taken out per Bruce's request (05/29/2013).

        'CommonLib.Delay(60 * 30)

        ' ***************************************************************

        'Get TM8 rec Data
        If Not TM8.GetRecdata(Tm8RecData) Then
            Form1.AppendText(TM8.ErrorMsg)
            Return False
        End If

         CommonLib.Delay(10)
        Success = True
        For Each UUT In UUTs
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

                If Not CommonLib.CreateDataTable(UUT("DT"), Tm1RecFields) Then
                    UUT("FAILED") = True
                    UUT("TAB").ImageIndex = StatusColor.FAILED
                    Success = False
                    Continue For
                End If
                If Not h2scan.GetRecData(UUT, Timeout_U2 * 4, UUT("DT")) Then
                    Form1.AppendText(h2scan.ErrorMsg, UUT:=UUT)
                    UUT("FAILED") = True
                    UUT("TAB").ImageIndex = StatusColor.FAILED
                    Success = False
                    Continue For
                End If
                ' add TM8 data
                If Not TM8.CombineTm8_Tm1_Data(UUT("DT"), Tm8RecData) Then
                    Form1.AppendText(h2scan.ErrorMsg, UUT:=UUT)
                    UUT("FAILED") = True
                    UUT("TAB").ImageIndex = StatusColor.FAILED
                    Success = False
                    Continue For
                End If
                UUT("GV").DataSource() = UUT("DT")
                'UUT("GV").FirstDisplayedCell = UUT("GV").Rows(UUT("GV").Rows.Count - 1).Cells(0)
                For Each col In UUT("GV").Columns
                    If Not (col.Name = "Timestamp" Or col.Name = "H2_OIl.PPM" Or col.Name = "H2.PPM" Or
                            col.Name = "TM8_ppm" Or col.Name = "TM8_gas_ppm") Then
                        UUT("GV").Columns(col.Name).visible = False
                    End If
                Next
                csv_filepath = ReportDir + "FINAL_TEST" + "\" + UUT("SN").Text + "\U2." + TimeStamp + ".csv"
                If Not CommonLib.ExportDataTableToCSV(UUT("DT"), csv_filepath) Then
                    Form1.AppendText("Problem creating csv file " + csv_filepath, UUT:=UUT)
                    UUT("FAILED") = True
                    UUT("TAB").ImageIndex = StatusColor.FAILED
                    Success = False
                    Continue For
                End If

                LastRowDateTime = UUT("DT").Rows(0)("Timestamp")
                WhereClause = "TM8_RecTime > '" + LastRowDateTime.Subtract(ts_3h).ToString + "'"
                TM8_gas_ppm = UUT("DT").Compute("AVG(TM8_gas_ppm)", WhereClause)

                'TM8_gas_ppm = UUT("DT").Rows(0)("TM8_gas_ppm")
                Form1.AppendText("Average TM8 gas ppm to be used for calibration:  " + TM8_gas_ppm.ToString, UUT:=UUT)
                If TM8_gas_ppm < 8500 Or TM8_gas_ppm > 11500 Then
                    Form1.AppendText("Expected TM8 gas ppm between 8500 & 11500", UUT:=UUT)
                    UUT("FAILED") = True
                    UUT("TAB").ImageIndex = StatusColor.FAILED
                    Success = False
                    Continue For
                End If

                If Not h2scan.Open(UUT) Then
                    Form1.AppendText(h2scan.ErrorMsg, UUT:=UUT)
                    UUT("FAILED") = True
                    UUT("TAB").ImageIndex = StatusColor.FAILED
                    Success = False
                    Continue For
                End If
                If Not h2scan.U2(UUT, TM8_gas_ppm) Then
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
                Form1.AppendText(UUT("SN").Text + ": Test_U2() caught" + ex.ToString, UUT:=UUT)
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

    Public Shared Function Test_U20() As Boolean
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
        Dim TM8_gas_ppm_valid As Boolean
        Dim AllFailed As Boolean = True
        Dim startTime As DateTime = Now
        Dim u2_time As Integer = 20

        While Now.Subtract(startTime).TotalHours < Timeout_U2
            If Not TM8.GetRecdata(Tm8RecData) Then
                Form1.AppendText(TM8.ErrorMsg)
                Return False
            End If
            TM8_gas_ppm = Tm8RecData.Rows(0)("gas_ppm")
            Form1.TM8_PPM.Text = TM8_gas_ppm.ToString
            If TM8_gas_ppm < 8500 Or TM8_gas_ppm > 11500 Then
                TM8_gas_in_spec = False
            Else
                TM8_gas_in_spec = True
            End If
            If Not TM8_gas_in_spec Then
                Form1.AppendText("TM8 gas ppm out of spec, expected between 8500 & 11500")
                Return False
            End If
            Form1.TimeoutLabel.Text = Math.Round(Timeout_U2 - Now.Subtract(startTime).TotalHours, 1).ToString + "h"
            CommonLib.Delay(60 * 15)
        End While

        ' ***************************** HACK ****************************
        ' Adding 30 mins to the end of this cycle as per requested by Thomas.
        ' This does not seem to help in reducing the negative spikes, so it is 
        ' taken out per Bruce's request (05/29/2013).

        'CommonLib.Delay(60 * 30)

        ' ***************************************************************

        ' login to UUT's
        CommonLib.Delay(10)
        Success = True
        For Each UUT In UUTs
            If UUT("SN").Text = "" Or UUT("FAILED") Then Continue For

            Try
                results = SF.Connect(UUT)
                If Not results.PassFail Then
                    UUT("FAILED") = True
                    UUT("TAB").ImageIndex = StatusColor.FAILED
                    Success = False
                End If
                CommonLib.Delay(1)
            Catch ex As Exception
                Form1.AppendText(UUT("SN").Text + ": Test_U2() caught" + ex.ToString, UUT:=UUT)
                UUT("FAILED") = True
                UUT("TAB").ImageIndex = StatusColor.FAILED
                Success = False
            End Try
        Next

        'Get TM8 rec Data
        If Not TM8.GetRecdata(Tm8RecData) Then
            Form1.AppendText(TM8.ErrorMsg)
            Return False
        End If

        CommonLib.Delay(10)
        'Get TM1 rec data
        For Each UUT In UUTs
            CommonLib.Delay(1)
            If UUT("SN").Text = "" Or UUT("FAILED") Then Continue For

            Dim Tm1RecFields() As String = Tm1Rev0RecFields
            If UUT("TM1 Hardware Version") <> Products(Product)("hardware version 0") Then
                Tm1RecFields = Tm1Rev2RecFields
            End If

            Try
                If Not CommonLib.CreateDataTable(UUT("DT"), Tm1RecFields) Then
                    UUT("FAILED") = True
                    UUT("TAB").ImageIndex = StatusColor.FAILED
                    Success = False
                    Continue For
                End If
                If Not h2scan.GetRecData(UUT, Timeout_U2 * 4, UUT("DT")) Then
                    Form1.AppendText(h2scan.ErrorMsg, UUT:=UUT)
                    UUT("FAILED") = True
                    UUT("TAB").ImageIndex = StatusColor.FAILED
                    Success = False
                    Continue For
                End If
                ' add TM8 data
                If Not TM8.CombineTm8_Tm1_Data(UUT("DT"), Tm8RecData) Then
                    Form1.AppendText(h2scan.ErrorMsg, UUT:=UUT)
                    UUT("FAILED") = True
                    UUT("TAB").ImageIndex = StatusColor.FAILED
                    Success = False
                    Continue For
                End If
                UUT("GV").DataSource() = UUT("DT")
                'UUT("GV").FirstDisplayedCell = UUT("GV").Rows(UUT("GV").Rows.Count - 1).Cells(0)
                For Each col In UUT("GV").Columns
                    If Not (col.Name = "Timestamp" Or col.Name = "H2_OIl.PPM" Or col.Name = "H2.PPM" Or
                            col.Name = "TM8_ppm" Or col.Name = "TM8_gas_ppm") Then
                        UUT("GV").Columns(col.Name).visible = False
                    End If
                Next
                csv_filepath = ReportDir + "FINAL_TEST" + "\" + UUT("SN").Text + "\U2." + TimeStamp + ".csv"
                If Not CommonLib.ExportDataTableToCSV(UUT("DT"), csv_filepath) Then
                    Form1.AppendText("Problem creating csv file " + csv_filepath, UUT:=UUT)
                    UUT("FAILED") = True
                    UUT("TAB").ImageIndex = StatusColor.FAILED
                    Success = False
                    Continue For
                End If
            Catch ex As Exception
                Form1.AppendText(UUT("SN").Text + ": Test_U2() caught" + ex.ToString, UUT:=UUT)
                UUT("FAILED") = True
                UUT("TAB").ImageIndex = StatusColor.FAILED
                Success = False
                Continue For
            End Try

            LastRowDateTime = UUT("DT").Rows(0)("Timestamp")
            WhereClause = "TM8_RecTime > '" + LastRowDateTime.Subtract(ts_3h).ToString + "'"
            Try
                TM8_gas_ppm = UUT("DT").Compute("AVG(TM8_gas_ppm)", WhereClause)
            Catch ex As Exception
                Form1.AppendText("Error calculating average TM8 ppm over prev 3 hours", UUT:=UUT)
                Form1.AppendText(ex.ToString)
                UUT("FAILED") = True
                UUT("TAB").ImageIndex = StatusColor.FAILED
                Success = False
                Continue For
            End Try
            'TM8_gas_ppm = UUT("DT").Rows(0)("TM8_gas_ppm")
            Form1.AppendText("Average TM8 gas ppm to be used for calibration:  " + TM8_gas_ppm.ToString, UUT:=UUT)
            If TM8_gas_ppm < 8500 Or TM8_gas_ppm > 11500 Then
                Form1.AppendText("Expected TM8 gas ppm between 8500 & 11500", UUT:=UUT)
                UUT("FAILED") = True
                UUT("TAB").ImageIndex = StatusColor.FAILED
                Success = False
                Continue For
            End If

            Try
                If Not h2scan.Open(UUT) Then
                    Form1.AppendText(h2scan.ErrorMsg, UUT:=UUT)
                    UUT("FAILED") = True
                    UUT("TAB").ImageIndex = StatusColor.FAILED
                    Success = False
                    Continue For
                End If
                If Not h2scan.U2(UUT, TM8_gas_ppm) Then
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
                Form1.AppendText(UUT("SN").Text + ": Test_U2() caught" + ex.ToString, UUT:=UUT)
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