Partial Class Tests
    Public Shared Function Test_VER_1000() As Boolean
        Dim Success As Boolean
        Dim AllTm1sInSpec As Boolean
        Dim T As New TM1

        If Not TM8_gas_in_spec Then
            Form1.AppendText("TM8 gas ppm not in spec")
            Return False
        End If

        ' diffPpmSpec = 15ppm: The difference between the TM8 & TM1 measurements
        ' delPpmSpec = 10ppm: The delta between the two consecutive of TM1 measurements (changed from 7 to 10ppm 4/16/15)
        Dim diffPpmSpec As Double = 15.0
        Dim delPpmSpec As Double = 10.0
        Success = T.WaitForTm1PpmInSpec(1000, 1200, 800, AllTm1sInSpec, Timeout_VER_1000, diffPpmSpec, delPpmSpec)
        'Success = T.WaitForTm1PpmInSpec(1000, 1200, 800, AllTm1sInSpec, 24, 10, 207)
        'Success = T.WaitForTm1PpmInSpec(1000, 1200, 800, AllTm1sInSpec, 16, 10, 207)
        'If AllTm1sInSpec Then
        '    Return Success
        'Else
        '    Return False
        'End If

        Return Success








        '' login to UUT's
        'Success = True
        'For Each UUT In UUTs
        '    If UUT("SN").Text = "" Or UUT("FAILED") Then Continue For
        '    results = SF.Connect(UUT)
        '    If Not results.PassFail Then
        '        UUT("FAILED") = True
        '        UUT("TAB").ImageIndex = StatusColor.FAILED
        '        Success = False
        '    End If
        'Next

        'If Now.Subtract(TM8_gas_start_in_spec).TotalMinutes > 60 Then
        '    Check_ppm_start_DT = Now
        'Else
        '    Check_ppm_start_DT = TM8_gas_start_in_spec
        'End If

        'AllTm1sInSpec = False
        'While Not AllTm1sInSpec And Now.Subtract(Check_ppm_start_DT).TotalHours < 12
        '    'Get TM8 rec Data
        '    If Not TM8.GetRecdata(Tm8RecData) Then
        '        Form1.AppendText(TM8.ErrorMsg)
        '        Return False
        '    End If

        '    n = Now.Subtract(Check_ppm_start_DT).TotalMinutes / 15
        '    If Not n > 0 Then Continue While
        '    AllTm1sInSpec = True
        '    For Each UUT In UUTs
        '        Tm1InSpec = True
        '        If UUT("SN").Text = "" Or UUT("FAILED") Then Continue For
        '        If Not CommonLib.CreateDataTable(UUT("DT"), Tm1RecFields) Then
        '            UUT("FAILED") = True
        '            UUT("TAB").ImageIndex = StatusColor.FAILED
        '            Success = False
        '            Continue For
        '        End If
        '        If Not h2scan.GetRecData(UUT, n, UUT("DT"), True) Then
        '            Form1.AppendText(h2scan.ErrorMsg, UUT:=UUT)
        '            UUT("FAILED") = True
        '            UUT("TAB").ImageIndex = StatusColor.FAILED
        '            Success = False
        '            Continue For
        '        End If
        '        If Not TM8.CombineTm8_Tm1_Data(UUT("DT"), Tm8RecData) Then
        '            Form1.AppendText(h2scan.ErrorMsg, UUT:=UUT)
        '            UUT("FAILED") = True
        '            UUT("TAB").ImageIndex = StatusColor.FAILED
        '            Success = False
        '            Continue For
        '        End If
        '        UUT("DT").Columns.Add("PER_ERROR", Type.GetType("System.Double"))
        '        For Each Row In UUT("DT").Rows
        '            If Not Row("TM8_ppm") = Nothing Then
        '                Row("PER_ERROR") = Math.Abs((Row("TM8_ppm") - Row("H2_OIL.PPM")) * 100.0 / Row("TM8_ppm"))
        '                'Row("PPM_Diff") = Math.Abs((Row("TM8_ppm") - Row("H2_OIL.PPM")))
        '            End If
        '        Next
        '        UUT("GV").DataSource() = UUT("DT")
        '        'UUT("GV").FirstDisplayedCell = UUT("GV").Rows(UUT("GV").Rows.Count - 1).Cells(0)
        '        For Each col In UUT("GV").Columns
        '            If Not (col.Name = "Timestamp" Or col.Name = "H2_OIL.PPM" Or col.Name = "H2.PPM" Or
        '                    col.Name = "TM8_ppm" Or col.Name = "TM8_gas_ppm" Or col.Name = "PER_ERROR") Then
        '                UUT("GV").Columns(col.Name).visible = False
        '            End If
        '        Next
        '        last_row_index = UUT("DT").Rows.Count - 1
        '        start_4_hour_window = last_row_index - 16
        '        If start_4_hour_window < 0 Then
        '            AllTm1sInSpec = False
        '            Tm1InSpec = False
        '        Else
        '            Dim lastPPM As Integer = 0
        '            For i = start_4_hour_window To last_row_index
        '                ' This would check for 10% error. That is incorrect on the 1000 ppm level.
        '                'If UUT("DT").Rows(i)("PER_ERROR") > 10 Then                        
        '                ppmDiff = Math.Abs((UUT("DT").Rows(i)("TM8_ppm") - UUT("DT").Rows(i)("H2_OIL.PPM")))
        '                Log.WriteLine(String.Format("SS = {0}: Last ppm = {1}", UUT("SN").Text, lastPPM))
        '                If (lastPPM <> 0) Then
        '                    Log.WriteLine(String.Format("SS = {0}: Current ppm = {1}", UUT("SN").Text, UUT("DT").Rows(i)("H2_OIL.PPM")))
        '                    If (Math.Abs((lastPPM - UUT("DT").Rows(i)("H2_OIL.PPM"))) > 7) Then
        '                        AllTm1sInSpec = False
        '                        Tm1InSpec = False
        '                        Log.WriteLine(String.Format("Precision spec of 7 failed. Actual value was {0}", Math.Abs((lastPPM - UUT("DT").Rows(i)("H2_OIL.PPM")))))
        '                    End If
        '                End If
        '                lastPPM = UUT("DT").Rows(i)("H2_OIL.PPM")
        '                Log.Write("PPM Diff calculated as : " + ppmDiff.ToString())
        '                If ppmDiff > 15 Then
        '                    AllTm1sInSpec = False
        '                    Tm1InSpec = False
        '                End If


        '            Next
        '        End If
        '        If Not Tm1InSpec Then
        '            Form1.AppendText(UUT("SN").Text + " not yet in spec for prev 4 hours", UUT:=UUT)
        '        End If
        '    Next
        '    If Not AllTm1sInSpec Then
        '        CommonLib.Delay(60 * 15)
        '    End If
        'End While

        'If AllTm1sInSpec Then
        '    Return Success
        'Else
        '    Return False
        'End If

        'Return Success
    End Function
End Class