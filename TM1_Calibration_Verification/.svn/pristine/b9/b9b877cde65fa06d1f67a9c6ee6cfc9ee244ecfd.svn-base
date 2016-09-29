Imports System.Threading.Thread
Imports System.Threading
Imports System.IO
'Imports System.ComponentModel

Partial Class Tests
    Public Shared Function Test_FINAL() As Boolean
        ' Prompt the user to drain oil out of all the passing units
        MessageBox.Show("Drain oil out of all the passing units now. Click OK to continue.", "Drain Oil Message")

        ' Dump data on pass units
        If DumpDataOnPassUnits() = False Then
            Form1.AppendText("DumpDataOnPassUnits() all failed. Test aborted!")
            Return False
        End If

        ' Do final configuration
        Dim DR As DialogResult
        DR = MessageBox.Show("Click yes when ready for final config or no to fail test" + vbCr +
                             "Units needed to be powered down after completion of this test." _
                             , "FINAL CONFIG?", MessageBoxButtons.YesNo)
        If DR = DialogResult.No Then
            Form1.AppendText("Operator not ready for final config.")
            ' LedTimer.Stop()
            Return False
        End If
        DoFinalConfigOnPassUnits()

        MessageBox.Show("Power down the rack. Click OK to continue.", "Final Configurations Complete")
        ' TODO: Need to test report generation for EM 5/27/2016
        ' Generate Test Report txt for the passing units
        If CreateTestReports() = False Then
            Form1.AppendText("CreateTestReports() all failed. Test aborted!")
            Return False
        End If

        ' Dump the data records on the failed units
        ' DumpDataOnFailUnits()

        ' Return true because at least some unit is passing
        Return True
    End Function

    'Private Shared Function CreateTestReports_Partial() As Boolean
    '    Dim TestReportData As Hashtable
    '    Dim testreport_filepath As String
    '    Dim TimeStamp As String
    '    Dim TestReport_File As FileStream
    '    Dim TestReport_FileWriter As StreamWriter
    '    Dim Get_PPMS_Success As Boolean
    '    Dim TargetPPM As Integer
    '    Dim TM8_PPM, TM1_PPM As Integer
    '    Dim TR As TestReport
    '    Dim somePassed As Boolean = False

    '    Form1.AppendText("CreateTestReports() is being invoked...")

    '    ' Extract the report template
    '    If CREATE_TEST_REPORT_DOC Then
    '        TR = New TestReport
    '        For i As Integer = 1 To 3
    '            Try
    '                If Not TR.ExtractTemplate Then
    '                    Form1.AppendText(TR.ErrorMsg)
    '                    MsgBox("Failed to extract the report template.")
    '                    CREATE_TEST_REPORT_DOC = False
    '                    Thread.Sleep(250)
    '                Else
    '                    Form1.AppendText("Test Report Template successfully extracted")
    '                    CREATE_TEST_REPORT_DOC = True
    '                    Exit For
    '                End If
    '            Catch ex As Exception
    '                CREATE_TEST_REPORT_DOC = False
    '                MsgBox("CreateTestReports().ExtractTemplate() Caught " + ex.Message)
    '            End Try
    '        Next
    '    End If

    '    If Not CREATE_TEST_REPORT_DOC Then
    '        Form1.AppendText("Error: Test Report Template was unsuccessfully extracted!!!")
    '    End If

    '    For Each UUT In UUTs
    '        If UUT("SN").Text = "" Or UUT("FAILED") Then Continue For

    '        Try
    '            TestReportData = New Hashtable
    '            TestReportData.Add("TM1_PPM_1000", 0)
    '            TestReportData.Add("TM8_PPM_1000", 0)
    '            TestReportData.Add("TM1_PPM_6000", 0)
    '            TestReportData.Add("TM8_PPM_6000", 0)
    '            TestReportData.Add("TM1_PPM_10000", 0)
    '            TestReportData.Add("TM8_PPM_10000", 0)
    '            TestReportData.Add("CAL_DATE", "66/66/66")
    '            Get_PPMS_Success = True

    '            TR = New TestReport
    '            If Not TR.CreateTestReport(UUT, TestReportData) Then
    '                Form1.AppendText(TR.ErrorMsg, UUT:=UUT)
    '                UUT("FAILED") = True
    '                UUT("TAB").ImageIndex = StatusColor.FAILED
    '            Else
    '                Form1.AppendText("Test Summary Report.doc successfully generated", UUT:=UUT)
    '                somePassed = True
    '            End If
    '        Catch ex As Exception
    '            Form1.AppendText("CreateTestReport() caught " + ex.ToString, UUT:=UUT)
    '            UUT("FAILED") = True
    '            UUT("TAB").ImageIndex = StatusColor.FAILED
    '        End Try
    '    Next
    '    Form1.AppendText("CreateTestReports() is complete.")

    '    Return somePassed
    'End Function

    Private Shared Function ExtractTemplate(ByVal doc As String) As TestReport
        Dim template As TestReport
        template = New TestReport
        For i As Integer = 1 To 3
            Try
                If Not template.ExtractTemplate(doc) Then
                    Form1.AppendText(template.ErrorMsg)
                    MsgBox("Failed to extract the report template.")
                    CREATE_TEST_REPORT_DOC = False
                    Thread.Sleep(250)
                Else
                    Form1.AppendText("Test Report Template successfully extracted")
                    CREATE_TEST_REPORT_DOC = True
                    Exit For
                End If
            Catch ex As Exception
                CREATE_TEST_REPORT_DOC = False
                MsgBox("CreateTestReports().ExtractTemplate() Caught " + ex.Message)
            End Try
        Next

        Return template
    End Function

    Private Shared Function CreateTestReports() As Boolean
        Dim TestReportData As Hashtable
        Dim testreport_filepath As String
        Dim TimeStamp As String
        Dim TestReport_File As FileStream
        Dim TestReport_FileWriter As StreamWriter
        Dim Get_PPMS_Success As Boolean
        Dim TargetPPM As Integer
        Dim TM8_PPM, TM1_PPM As Integer
        Dim report As TestReport = Nothing
        Dim theReport As TestReport = Nothing
        Dim emReport As TestReport = Nothing
        Dim templateDoc As String
        Dim somePassed As Boolean = False

        Form1.AppendText("CreateTestReports() is being invoked...")

        For Each UUT In UUTs
            If UUT("SN").Text = "" Or UUT("FAILED") Then Continue For

            Try
                ' Get the correct template for the type of the UUT
                If UUT("embedded moisture") = "yes" Then
                    templateDoc = "825-0074-00-EM TM1 Test Summary Report RevC.doc"
                    If IsNothing(emReport) = True Then
                        emReport = ExtractTemplate(templateDoc)
                    End If
                    theReport = emReport
                Else
                    templateDoc = "825-0074-00 TM1 Test Summary Report RevC.doc"
                    If IsNothing(report) = True Then
                        report = ExtractTemplate(templateDoc)
                    End If
                    theReport = report
                End If

                If Not CREATE_TEST_REPORT_DOC Or IsNothing(theReport) = True Then
                    Form1.AppendText("Failed to extract test report. Test skipped!")
                    Continue For
                End If

                TestReportData = New Hashtable
                TestReportData.Add("TM1_PPM_1000", 0)
                TestReportData.Add("TM8_PPM_1000", 0)
                TestReportData.Add("TM1_PPM_6000", 0)
                TestReportData.Add("TM8_PPM_6000", 0)
                TestReportData.Add("TM1_PPM_10000", 0)
                TestReportData.Add("TM8_PPM_10000", 0)
                TestReportData.Add("CAL_DATE", "66/66/66")
                Get_PPMS_Success = True
                For Each TargetPPM In {1000, 6000, 10000}
                    If Not CommonLib.GetAverageH2Ppms(UUT("SN").Text, TargetPPM, TM1_PPM, TM8_PPM) Then
                        Form1.AppendText(CommonLib.ErrorMsg, UUT:=UUT)
                        Get_PPMS_Success = False
                        Exit For
                    End If
                    Form1.AppendText("target = " + TargetPPM.ToString + ", TM1 = " + TM1_PPM.ToString + ", TM8 = " + TM8_PPM.ToString, UUT:=UUT)
                    TestReportData("TM1_PPM_" + TargetPPM.ToString) = TM1_PPM
                    TestReportData("TM8_PPM_" + TargetPPM.ToString) = TM8_PPM
                Next

                ' If failed, mark the unit as fail and skip it
                If Not Get_PPMS_Success Then
                    UUT("FAILED") = True
                    UUT("TAB").ImageIndex = StatusColor.FAILED
                    Continue For
                End If

                ' Add Cal date
                ' Temporary added for this run only 11/3/15
                'UUT("CAL_DATE") = "02/12/2015"
                'UUT("TM1 Firmware Version") = "1.3.5559"
                'UUT("Sensor Firmware Version") = "3.955B"

                Form1.AppendText("cal date:  " + UUT("CAL_DATE"), UUT:=UUT)
                TestReportData("CAL_DATE") = UUT("CAL_DATE")
                TestReportData("TM1 Firmware Version") = UUT("TM1 Firmware Version")
                TestReportData("Sensor Firmware Version") = UUT("Sensor Firmware Version")

                If UUT("SN").Text = "" Or UUT("FAILED") Then Continue For
                TimeStamp = Format(Date.UtcNow, "yyyyMMddHHmmss")
                testreport_filepath = ReportDir + "FINAL_TEST" + "\" + UUT("SN").Text + "\" + UUT("SN").Text + "_TestReport." + TimeStamp + ".txt"

                TestReport_File = New FileStream(testreport_filepath, FileMode.Create, FileAccess.Write)
                TestReport_FileWriter = New StreamWriter(TestReport_File)
                TestReport_FileWriter.AutoFlush = True

                TestReport_FileWriter.WriteLine("SN = " + UUT("SN").Text)
                For Each TargetPPM In {1000, 6000, 10000}
                    TestReport_FileWriter.WriteLine("TARGET PPM = " + TargetPPM.ToString +
                                    ", TM8_PPM = " + TestReportData("TM8_PPM_" + TargetPPM.ToString).ToString +
                                    ", TM1_PPM = " + TestReportData("TM1_PPM_" + TargetPPM.ToString).ToString)
                Next
                TestReport_FileWriter.WriteLine("CAL_DATE = " + TestReportData("CAL_DATE"))
                TestReport_FileWriter.Close()
                TestReport_File.Close()
                somePassed = True

                If Not CREATE_TEST_REPORT_DOC Then Continue For
                Form1.AppendText(UUT("SN").Text + ":  Generating Test Summary Report.doc", UUT:=UUT)
                'TR = New TestReport
                If Not theReport.CreateTestReport(UUT, TestReportData, templateDoc) Then
                    Form1.AppendText(theReport.ErrorMsg, UUT:=UUT)
                    UUT("FAILED") = True
                    UUT("TAB").ImageIndex = StatusColor.FAILED
                Else
                    Form1.AppendText("Test Summary Report.doc successfully generated", UUT:=UUT)
                End If
            Catch ex As Exception
                Form1.AppendText("CreateTestReport() caught " + ex.ToString, UUT:=UUT)
                UUT("FAILED") = True
                UUT("TAB").ImageIndex = StatusColor.FAILED
            End Try

            Application.DoEvents()
        Next
        Form1.AppendText("CreateTestReports() is complete.")

        Return somePassed
    End Function

    Private Shared Function DumpDataOnPassUnits() As Boolean
        Dim somePassed As Boolean = False

        Form1.AppendText("DumpDataOnGoodUnits() is being invoked...")
        For Each UUT In UUTs
            ' Skipp failing units
            If UUT("SN").Text = "" Or UUT("FAILED") Then Continue For

            Application.DoEvents()
            Try
                ' dump h2scan logs
                Form1.AppendText(UUT("SN").Text + ":  Starting DumpAndClear()", UUT:=UUT)
                Log.WriteLine(String.Format("DumpAndClear {0}", UUT("SN").Text))
                If Not DumpAndClear(UUT) Then
                    Form1.AppendText(UUT("SN").Text + ": DumpAndClear() failed.", UUT:=UUT)
                    UUT("FAILED") = True
                    UUT("TAB").ImageIndex = StatusColor.FAILED
                Else
                    Form1.AppendText(UUT("SN").Text + ": DumpAndClear() finished.", UUT:=UUT)
                    somePassed = True
                End If
                CommonLib.Delay(1)
            Catch ex As Exception
                Form1.AppendText(UUT("SN").Text + ": DumpAndClear() caught " + ex.ToString, UUT:=UUT)
                UUT("FAILED") = True
                UUT("TAB").ImageIndex = StatusColor.FAILED
            End Try
        Next
        Form1.AppendText("DumpDataOnGoodUnits() is complete.")

        Return somePassed
    End Function

    Private Shared Function DoFinalConfigOnPassUnits() As Boolean
        Dim somePassed As Boolean = False

        Form1.AppendText("DoFinalConfigOnPassUnits() is being invoked...")
        For Each UUT In UUTs
            ' Skipp failing units
            If UUT("SN").Text = "" Or UUT("FAILED") Then Continue For

            Application.DoEvents()

            Try
                If Not DoFinalConfig(UUT) Then
                    Form1.AppendText(UUT("SN").Text + ": DoFinalConfig() failed.", UUT:=UUT)
                    UUT("FAILED") = True
                    UUT("TAB").ImageIndex = StatusColor.FAILED
                Else
                    Form1.AppendText(UUT("SN").Text + ": DoFinalConfig() passed.", UUT:=UUT)
                    somePassed = True
                End If
            Catch ex As Exception
                Form1.AppendText(UUT("SN").Text + ": DoFinalConfigOnPassUnits() caught " + ex.ToString, UUT:=UUT)
                UUT("FAILED") = True
                UUT("TAB").ImageIndex = StatusColor.FAILED
            End Try
        Next
        Form1.AppendText("DoFinalConfigOnPassUnits() is complete.")

        Return somePassed
    End Function

    Private Shared Function DoFinalConfig(ByRef UUT As Hashtable) As Boolean
        Dim h2scan As New H2SCAN_debug
        Dim ProductInfo As Hashtable
        Dim T As New TM1
        Dim C As New config
        Dim SF As New SerialFunctions
        Dim Response As String
        Dim results As ReturnResults
        Dim VersionInfo As Hashtable

        ' login to UUT's
        results = SF.Connect(UUT)
        If Not results.PassFail Then
            Form1.AppendText(UUT("SN").Text + ":  Failed to connect.", UUT:=UUT)
            Return False
        End If
        Form1.AppendText(UUT("SN").Text + ":  Successfully connected.", UUT:=UUT)

        ' Perform a factory reset.
        CommonLib.Delay(3)
        Form1.AppendText(UUT("SN").Text + " resetting config to defaults", UUT:=UUT)
        If Not C.ConfigFactoryReset(UUT, True) Then
            Form1.AppendText(C.Results, UUT:=UUT, LogToResults:=False)
            Form1.AppendText(C.ErrorMsg, UUT:=UUT, LogToResults:=False)
            Return False
        End If
        Form1.AppendText(C.Results, UUT:=UUT, LogToResults:=False)

        If UUT("hardware version") <> Products(Product)("hardware version 0") Then
            If Not T.SetHwRev2Config(UUT) Then
                Form1.AppendText("Failed SetHwRev2Config()", UUT:=UUT, LogToResults:=False)
                Return False
            End If
        End If

        ' clear TM1 logs/records
        CommonLib.Delay(10)
        Form1.AppendText("Clearing " + UUT("SN").Text + " logs and records", UUT:=UUT)
        If Not SF.Cmd(UUT, Response, "fr -A", 30, Quiet:=True) Then
            Form1.AppendText(Response, UUT:=UUT, LogToResults:=False)
            Form1.AppendText("failed sending cmd 'fr -A'", UUT:=UUT, LogToResults:=False)
            Form1.AppendText(SF.ErrorMsg, UUT:=UUT, LogToResults:=False)
            Return False
        End If
        Form1.AppendText(Response, UUT:=UUT, LogToResults:=False)

        CommonLib.Delay(3)
        Form1.AppendText("Reconnecting to " + UUT("SN").Text, UUT:=UUT, LogToResults:=False)
        results = SF.Connect(UUT, True)
        Form1.AppendText(SF.RtnResults, UUT:=UUT, LogToResults:=False)
        If Not results.PassFail Then
            Form1.AppendText("Problem reconnecting to after config reset and reboot", UUT:=UUT, LogToResults:=False)
            Form1.AppendText(SF.ErrorMsg, UUT:=UUT, LogToResults:=False)
            Return False
        End If

        CommonLib.Delay(3)
        Form1.AppendText(UUT("SN").Text + " verifying recs cleared", UUT:=UUT)
        If Not SF.Cmd(UUT, Response, "rec", 5, Quiet:=True) Then
            Form1.AppendText(Response, UUT:=UUT, LogToResults:=False)
            Form1.AppendText("failed sending cmd 'rec'", UUT:=UUT, LogToResults:=False)
            Form1.AppendText(SF.ErrorMsg, UUT:=UUT, LogToResults:=False)
            Return False
        End If
        Form1.AppendText(Response, UUT:=UUT, LogToResults:=False)

        If Not Response.Contains("total_recs = 0") Then
            Form1.AppendText("recs not cleared", UUT:=UUT, LogToResults:=False)
            Return False
        End If

        ' Verify serial number is still set
        If Not T.GetVersionInfo(UUT, VersionInfo, True) Then
            Form1.AppendText(T.Results, UUT:=UUT, LogToResults:=False)
            Form1.AppendText(UUT("SN").Text + " problem getting version info", UUT:=UUT, LogToResults:=False)
            Form1.AppendText(T.ErrorMsg, UUT:=UUT, LogToResults:=False)
            Return False
        End If
        Form1.AppendText(T.Results, UUT:=UUT, LogToResults:=False)

        If Not VersionInfo("serial number") = UUT("SN").Text Then
            Form1.AppendText("Expected 'serial number' = " + UUT("SN").Text, UUT:=UUT, LogToResults:=False)
            Return False
        End If

        ' It is considered pass as this point
        Return True
    End Function

    Private Shared Function DumpAndClear(ByRef UUT As Hashtable) As Boolean
        Dim h2scan As New H2SCAN_debug
        Dim ProductInfo As Hashtable
        Dim T As New TM1
        Dim C As New config
        Dim SF As New SerialFunctions
        Dim Response As String
        Dim results As ReturnResults
        Dim VersionInfo As Hashtable

        ' login to UUT's
        results = SF.Connect(UUT)
        If Not results.PassFail Then
            Form1.AppendText(UUT("SN").Text + ":  Failed to connect.", UUT:=UUT)
            Return False
        End If
        CommonLib.Delay(10)

        Form1.AppendText(UUT("SN").Text + ":  Successfully connected.", UUT:=UUT)

        Form1.AppendText("Opening H2SCAN CLI", UUT:=UUT, LogToResults:=False)

        ' Try up to three times to log on the sensor
        For i As Integer = 0 To 3
            If Not h2scan.Open(UUT, True) Then
                Form1.AppendText(h2scan.Results, UUT:=UUT, LogToResults:=False)
                Form1.AppendText(h2scan.ErrorMsg, UUT:=UUT, LogToResults:=False)
                If i < 3 Then
                    Thread.Sleep(50)
                ElseIf i = 3 Then
                    Return False
                End If
            Else
                Exit For
            End If
        Next
        Form1.AppendText(h2scan.Results, UUT:=UUT, LogToResults:=False)

        Form1.AppendText("Getting H2SCAN cal date", UUT:=UUT)
        If Not h2scan.GetProductInfo(UUT, ProductInfo, True) Then
            Form1.AppendText(h2scan.Results, UUT:=UUT, LogToResults:=False)
            Form1.AppendText("Problem GetProductInfo() on " + UUT("SN").Text, UUT:=UUT, LogToResults:=False)
            Form1.AppendText(h2scan.ErrorMsg, UUT:=UUT, LogToResults:=False)
            h2scan.Close(UUT, True)
            Form1.AppendText(h2scan.Results, UUT:=UUT, LogToResults:=False)
            Return False
        End If
        Form1.AppendText(h2scan.Results + System.Environment.NewLine, UUT:=UUT, LogToResults:=False)
        UUT("CAL_DATE") = ProductInfo("TouchUp")
        UUT("Sensor Firmware Version") = ProductInfo("Firmware Rev")
        Form1.AppendText("TouchUp cal date = " + UUT("CAL_DATE"), UUT:=UUT)

        CommonLib.Delay(3)
        Form1.AppendText("Dumping " + UUT("SN").Text + " records", UUT:=UUT)
        If Not h2scan.DumpRecords(UUT, ProductInfo, True) Then
            Form1.AppendText(h2scan.Results, UUT:=UUT, LogToResults:=False)
            Form1.AppendText(h2scan.ErrorMsg, UUT:=UUT, LogToResults:=False)
            h2scan.Close(UUT, True)
            Form1.AppendText(h2scan.Results, UUT:=UUT, LogToResults:=False)
            Return False
        End If
        Form1.AppendText(h2scan.Results, UUT:=UUT, LogToResults:=False)

        CommonLib.Delay(3)
        Form1.AppendText("Clearing " + UUT("SN").Text + " records", UUT:=UUT)
        If Not h2scan.ClearRecords(UUT, True) Then
            Form1.AppendText(h2scan.Results, UUT:=UUT, LogToResults:=False)
            Form1.AppendText(h2scan.ErrorMsg, UUT:=UUT, LogToResults:=False)
            h2scan.Close(UUT, True)
            Form1.AppendText(h2scan.Results, UUT:=UUT, LogToResults:=False)
            Return False
        End If
        Form1.AppendText(h2scan.Results, UUT:=UUT, LogToResults:=False)

        CommonLib.Delay(3)
        Form1.AppendText("Clearing " + UUT("SN").Text + " short term memory", UUT:=UUT)
        If Not h2scan.ClearMemory(UUT, True) Then
            Form1.AppendText(h2scan.Results, UUT:=UUT, LogToResults:=False)
            Form1.AppendText(h2scan.ErrorMsg, UUT:=UUT, LogToResults:=False)
            h2scan.Close(UUT, True)
            Form1.AppendText(h2scan.Results, UUT:=UUT, LogToResults:=False)
            Return False
        End If
        Form1.AppendText(h2scan.Results, UUT:=UUT, LogToResults:=False)

        ' Done at the sensor mode so close it
        CommonLib.Delay(3)
        If Not h2scan.Close(UUT, True) Then
            Form1.AppendText(h2scan.Results, UUT:=UUT, LogToResults:=False)
            Form1.AppendText(h2scan.ErrorMsg, UUT:=UUT, LogToResults:=False)
            h2scan.Close(UUT, True)
            Form1.AppendText(h2scan.Results, UUT:=UUT, LogToResults:=False)
            Return False
        End If
        Form1.AppendText(h2scan.Results, UUT:=UUT, LogToResults:=False)

        ' Dump TM1 logs/records
        CommonLib.Delay(10)
        Form1.AppendText("Dumping " + UUT("SN").Text + " config", UUT:=UUT)
        If Not T.DumpConfig(UUT, True) Then
            Form1.AppendText(T.Results, UUT:=UUT, LogToResults:=False)
            Form1.AppendText(T.ErrorMsg, UUT:=UUT, LogToResults:=False)
            Return False
        End If
        Form1.AppendText(T.Results + vbCr, UUT:=UUT, LogToResults:=False)

        CommonLib.Delay(3)
        Form1.AppendText("Dumping " + UUT("SN").Text + " rec's", UUT:=UUT)
        If Not T.DumpRecs(UUT, True) Then
            Form1.AppendText(T.Results, UUT:=UUT, LogToResults:=False)
            Form1.AppendText(T.ErrorMsg, UUT:=UUT, LogToResults:=False)
            Return False
        End If
        Form1.AppendText(T.Results, UUT:=UUT, LogToResults:=False)

        CommonLib.Delay(3)
        Form1.AppendText("Dumping " + UUT("SN").Text + " events", UUT:=UUT)
        If Not T.DumpEvents(UUT, True) Then
            Form1.AppendText(T.Results, UUT:=UUT, LogToResults:=False)
            Form1.AppendText(T.ErrorMsg, UUT:=UUT, LogToResults:=False)
            Return False
        End If
        Form1.AppendText(T.Results, UUT:=UUT, LogToResults:=False)

        ' Set sensor.mode to FIELD mdoe
        'CommonLib.Delay(3)
        'Form1.AppendText("Setting " + UUT("SN").Text + " operating mode to FIELD", UUT:=UUT)
        'If Not h2scan.SetSensorMode(UUT, H2SCAN_debug.H2SCAN_OP_MODE.FIELD) Then
        '    Form1.AppendText(h2scan.Results, UUT:=UUT, LogToResults:=False)
        '    Form1.AppendText("Problem setting " + UUT("SN").Text + " to field operating mode", UUT:=UUT, LogToResults:=False)
        '    Form1.AppendText(h2scan.ErrorMsg, UUT:=UUT, LogToResults:=False)
        '    Return False
        'End If

        Form1.AppendText(h2scan.Results, UUT:=UUT, LogToResults:=False)

        ' It is considered pass as this point
        Return True

    End Function

    Private Shared Function DumpDataOnFailUnits() As Boolean
        Dim h2scan As New H2SCAN_debug
        Dim T As New TM1
        Dim C As New config
        Dim SF As New SerialFunctions
        Dim results As ReturnResults
        Dim ProductInfo As Hashtable

        Form1.AppendText("DumpDataOnFailUnits() is being invoked...")

        For Each UUT In UUTs
            Try
                ' Only login to the failed units
                If UUT("FAILED") Then
                    CommonLib.Delay(1)
                    results = SF.Connect(UUT)
                    If Not results.PassFail Then
                        Continue For
                    End If

                    ' Try to log to H2Scan
                    Form1.AppendText("Opening H2SCAN CLI", UUT:=UUT, LogToResults:=False)
                    If Not h2scan.Open(UUT, True) Then
                        Form1.AppendText(h2scan.Results, UUT:=UUT, LogToResults:=False)
                        Form1.AppendText(h2scan.ErrorMsg, UUT:=UUT, LogToResults:=False)
                        Continue For
                    End If

                    ' Try to dump the H2Scan records
                    Form1.AppendText("Dumping " + UUT("SN").Text + " records", UUT:=UUT)
                    If Not h2scan.DumpRecords(UUT, ProductInfo, True) Then
                        Form1.AppendText(h2scan.Results, UUT:=UUT, LogToResults:=False)
                        Form1.AppendText(h2scan.ErrorMsg, UUT:=UUT, LogToResults:=False)
                        h2scan.Close(UUT, True)
                        Form1.AppendText(h2scan.Results, UUT:=UUT, LogToResults:=False)
                        Continue For
                    End If

                    ' Try to close the H2Scan
                    If Not h2scan.Close(UUT, True) Then
                        Form1.AppendText(h2scan.Results, UUT:=UUT, LogToResults:=False)
                        Form1.AppendText(h2scan.ErrorMsg, UUT:=UUT, LogToResults:=False)
                        h2scan.Close(UUT, True)
                        Form1.AppendText(h2scan.Results, UUT:=UUT, LogToResults:=False)
                        Continue For
                    End If

                    ' Try to dump the TM1 records
                    Form1.AppendText("Dumping " + UUT("SN").Text + " rec's", UUT:=UUT)
                    If Not T.DumpRecs(UUT, True) Then
                        Form1.AppendText(T.Results, UUT:=UUT, LogToResults:=False)
                        Form1.AppendText(T.ErrorMsg, UUT:=UUT, LogToResults:=False)
                        Continue For
                    End If

                End If
            Catch ex As Exception
                Form1.AppendText(UUT("SN").Text + ": Caught " + ex.Message, UUT:=UUT, LogAllUuts:=False)
                Continue For
            End Try

        Next

        Form1.AppendText("DumpDataOnFailUnits() is complete.")

        Return True
    End Function

    '' This is used for creating reports from a final test abort run
    'Public Shared Function Test_FINAL_CreateReports() As Boolean
    '    Dim Success As Boolean
    '    Dim results As ReturnResults
    '    Dim SF As New SerialFunctions
    '    Dim h2scan As New H2SCAN_debug
    '    Dim T As New TM1
    '    Dim C As New config
    '    Dim AllFailed As Boolean = True
    '    Dim DR As DialogResult
    '    Dim UUT_Dir As String
    '    Dim TestReportData As Hashtable
    '    Dim testreport_filepath As String
    '    Dim TimeStamp As String
    '    Dim TestReport_File As FileStream
    '    Dim TestReport_FileWriter As StreamWriter
    '    Dim Get_PPMS_Success As Boolean
    '    Dim TargetPPM As Integer
    '    Dim TM8_PPM, TM1_PPM As Integer
    '    Dim TR As TestReport

    '    ' Generate Test Report txt
    '    If CREATE_TEST_REPORT_DOC Then
    '        TR = New TestReport
    '        If Not TR.ExtractTemplate Then
    '            Form1.AppendText(TR.ErrorMsg)
    '            MsgBox("Auto Test Report Generation Aborted due to error")
    '            CREATE_TEST_REPORT_DOC = False
    '        Else
    '            Form1.AppendText("Test Report Template successfully extracted")
    '        End If
    '    End If

    '    For Each UUT In UUTs
    '        If UUT("SN").Text = "" Or UUT("FAILED") Then Continue For

    '        UUT("CAL_DATE") = "02/09/2015"
    '        UUT("TM1 Firmware Version") = "1.1.5321"
    '        UUT("Sensor Firmware Version") = "3.36B"

    '        TestReportData = New Hashtable
    '        TestReportData.Add("TM1_PPM_1000", 0)
    '        TestReportData.Add("TM8_PPM_1000", 0)
    '        TestReportData.Add("TM1_PPM_6000", 0)
    '        TestReportData.Add("TM8_PPM_6000", 0)
    '        TestReportData.Add("TM1_PPM_10000", 0)
    '        TestReportData.Add("TM8_PPM_10000", 0)
    '        TestReportData.Add("CAL_DATE", "66/66/66")
    '        Get_PPMS_Success = True
    '        For Each TargetPPM In {1000, 6000, 10000}
    '            Try
    '                If Not CommonLib.GetAverageH2Ppms(UUT("SN").Text, TargetPPM, TM1_PPM, TM8_PPM) Then
    '                    Form1.AppendText(CommonLib.ErrorMsg, UUT:=UUT)
    '                    Get_PPMS_Success = False
    '                End If
    '                If Not Get_PPMS_Success Then Exit For
    '                Form1.AppendText("target = " + TargetPPM.ToString + ", TM1 = " + TM1_PPM.ToString + ", TM8 = " + TM8_PPM.ToString, UUT:=UUT)
    '                TestReportData("TM1_PPM_" + TargetPPM.ToString) = TM1_PPM
    '                TestReportData("TM8_PPM_" + TargetPPM.ToString) = TM8_PPM
    '            Catch ex As Exception
    '                Form1.AppendText(UUT("SN").Text + ": DumpAndClear() caught " + ex.ToString, UUT:=UUT)
    '                Get_PPMS_Success = False
    '                Exit For
    '            End Try
    '        Next
    '        If Not Get_PPMS_Success Then
    '            UUT("FAILED") = True
    '            UUT("TAB").ImageIndex = StatusColor.FAILED
    '            Success = False
    '            Continue For
    '        End If

    '        ' Add Cal date
    '        Form1.AppendText("cal date:  " + UUT("CAL_DATE"), UUT:=UUT)
    '        TestReportData("CAL_DATE") = UUT("CAL_DATE")
    '        TestReportData("TM1 Firmware Version") = UUT("TM1 Firmware Version")
    '        TestReportData("Sensor Firmware Version") = UUT("Sensor Firmware Version")

    '        If UUT("SN").Text = "" Or UUT("FAILED") Then Continue For
    '        TimeStamp = Format(Date.UtcNow, "yyyyMMddHHmmss")
    '        testreport_filepath = ReportDir + "FINAL_TEST" + "\" + UUT("SN").Text + "\" + UUT("SN").Text + "_TestReport." + TimeStamp + ".txt"

    '        Try
    '            TestReport_File = New FileStream(testreport_filepath, FileMode.Create, FileAccess.Write)
    '            TestReport_FileWriter = New StreamWriter(TestReport_File)
    '            TestReport_FileWriter.AutoFlush = True
    '        Catch ex As Exception
    '            Form1.AppendText("Problem opening " + testreport_filepath + " for writing", UUT:=UUT)
    '            Form1.AppendText(ex.ToString, UUT:=UUT)
    '            UUT("FAILED") = True
    '            UUT("TAB").ImageIndex = StatusColor.FAILED
    '            Success = False
    '            Continue For
    '        End Try

    '        Try
    '            TestReport_FileWriter.WriteLine("SN = " + UUT("SN").Text)
    '            For Each TargetPPM In {1000, 6000, 10000}
    '                TestReport_FileWriter.WriteLine("TARGET PPM = " + TargetPPM.ToString +
    '                                ", TM8_PPM = " + TestReportData("TM8_PPM_" + TargetPPM.ToString).ToString +
    '                                ", TM1_PPM = " + TestReportData("TM1_PPM_" + TargetPPM.ToString).ToString)
    '            Next
    '            TestReport_FileWriter.WriteLine("CAL_DATE = " + TestReportData("CAL_DATE"))
    '            TestReport_FileWriter.Close()
    '            TestReport_File.Close()

    '            If Not CREATE_TEST_REPORT_DOC Then Continue For
    '            Form1.AppendText(UUT("SN").Text + ":  Generating Test Summary Report.doc", UUT:=UUT)
    '            TR = New TestReport
    '            If Not TR.CreateTestReport(UUT, TestReportData) Then
    '                Form1.AppendText(TR.ErrorMsg, UUT:=UUT)
    '                UUT("FAILED") = True
    '                UUT("TAB").ImageIndex = StatusColor.FAILED
    '                Success = False
    '            Else
    '                Form1.AppendText("Test Summary Report.doc successfully generated", UUT:=UUT)
    '            End If
    '        Catch ex As Exception
    '            Form1.AppendText("DumpAndClear() caught " + ex.ToString, UUT:=UUT)
    '            UUT("FAILED") = True
    '            UUT("TAB").ImageIndex = StatusColor.FAILED
    '            Success = False
    '        End Try
    '    Next

    '    ' Dump the data records on the failed units
    '    'DumpDataOnFailUnits()

    '    If AllFailed Then
    '        Return False
    '    Else
    '        Return True
    '    End If
    '    'Return Success
    'End Function

End Class

Public Class ClearLogsAndConfig
    Friend UUT As Object
    Friend _Messages As Queue(Of String) = New Queue(Of String)
    Friend _RetVal As Boolean = False
    Friend _stop_test As Boolean = False
    'Public Event UutDataAvailable(ByVal UUT As Hashtable)

    Public Property RetVal As Boolean
        Get
            Return _RetVal
        End Get
        Set(value As Boolean)

        End Set
    End Property

    Public Property stop_test As Boolean
        Get

        End Get
        Set(value As Boolean)
            _stop_test = value
        End Set
    End Property


    'Sub New()
    '    AddHandler UutDataAvailable, AddressOf Form1.UutDataAvailable_Handler
    'End Sub

    Public Function MessagesDequeue() As String
        Return _Messages.Dequeue
    End Function

    Public Function MessageCnt() As Integer
        Return _Messages.Count
    End Function

    'Sub Raise(ByVal [event] As [Delegate], ByRef UUT As Object)
    '    'If the event has no handlers just exit the method call.
    '    If [event] Is Nothing Then Return

    '    'Enumerates through the list of handlers.
    '    For Each D As [Delegate] In [event].GetInvocationList()
    '        'Casts the handler's parent instance to ISynchronizeInvoke.
    '        Dim T As ISynchronizeInvoke = DirectCast(D.Target, ISynchronizeInvoke)

    '        'If an invoke is required (working on a seperate thread) then invoke it
    '        'on the parent thread, otherwise we can invoke it directly.
    '        If T.InvokeRequired Then T.Invoke(D, UUT) Else D.DynamicInvoke(UUT)
    '    Next
    'End Sub

    Sub DumpAndClear()
        Dim h2scan As New H2SCAN_debug
        Dim ProductInfo As Hashtable
        Dim T As New TM1
        Dim C As New config
        Dim SF As New SerialFunctions
        Dim Response As String
        Dim results As ReturnResults
        Dim VersionInfo As Hashtable

        'For i = 0 To 10
        '    CommonLib.Delay(2)
        '    _Messages.Enqueue("Message " + i.ToString)
        'Next
        'CommonLib.Delay(10)
        '_Messages.Enqueue("DONE")
        'If UUT("SN").Text = "TM1010112260001" Then
        '    _RetVal = False
        'Else
        '    _RetVal = True
        'End If
        ''_RetVal = True
        'Exit Sub

        ' dump h2scan logs
        _Messages.Enqueue("Dumping H2SCAN log files")

        _Messages.Enqueue("Opening H2SCAN CLI")
        If Not h2scan.Open(UUT, True) Then
            _Messages.Enqueue(h2scan.Results)
            _Messages.Enqueue(h2scan.ErrorMsg)
            _RetVal = False
            Exit Sub
        End If
        _Messages.Enqueue(h2scan.Results)
        If _stop_test Then Exit Sub

        _Messages.Enqueue("Getting H2SCAN cal date")
        If Not h2scan.GetProductInfo(UUT, ProductInfo, True) Then
            _Messages.Enqueue(h2scan.Results)
            _Messages.Enqueue("Problem setting " + UUT("SN").Text + " to field operating mode")
            _Messages.Enqueue(h2scan.ErrorMsg)
            _RetVal = False
            h2scan.Close(UUT, True)
            _Messages.Enqueue(h2scan.Results)
            Exit Sub
        End If
        _Messages.Enqueue(h2scan.Results + System.Environment.NewLine)
        UUT("CAL_DATE") = ProductInfo("TouchUp")
        UUT("Sensor Firmware Version") = ProductInfo("Firmware Rev")
        _Messages.Enqueue("TouchUp cal date = " + UUT("CAL_DATE"))
        If _stop_test Then Exit Sub

        _Messages.Enqueue("Setting " + UUT("SN").Text + " operating mode to FIELD")
        If Not h2scan.SetOperatingMode(UUT, H2SCAN_debug.H2SCAN_OP_MODE.FIELD, True) Then
            _Messages.Enqueue(h2scan.Results)
            _Messages.Enqueue("Problem setting " + UUT("SN").Text + " to field operating mode")
            _Messages.Enqueue(h2scan.ErrorMsg)
            _RetVal = False
            h2scan.Close(UUT, True)
            _Messages.Enqueue(h2scan.Results)
            Exit Sub
        End If
        _Messages.Enqueue(h2scan.Results)
        If _stop_test Then Exit Sub

        _Messages.Enqueue("Dumping " + UUT("SN").Text + " records")
        If Not h2scan.DumpRecords(UUT, ProductInfo, True) Then
            _Messages.Enqueue(h2scan.Results)
            _Messages.Enqueue(h2scan.ErrorMsg)
            _RetVal = False
            h2scan.Close(UUT, True)
            _Messages.Enqueue(h2scan.Results)
            Exit Sub
        End If
        _Messages.Enqueue(h2scan.Results)
        If _stop_test Then Exit Sub

        _Messages.Enqueue("Clearing " + UUT("SN").Text + " records")
        If Not h2scan.ClearRecords(UUT, True) Then
            _Messages.Enqueue(h2scan.Results)
            _Messages.Enqueue(h2scan.ErrorMsg)
            _RetVal = False
            h2scan.Close(UUT, True)
            _Messages.Enqueue(h2scan.Results)
            Exit Sub
        End If
        _Messages.Enqueue(h2scan.Results)
        If _stop_test Then Exit Sub

        _Messages.Enqueue("Clearing " + UUT("SN").Text + " short term memory")
        If Not h2scan.ClearMemory(UUT, True) Then
            _Messages.Enqueue(h2scan.Results)
            _Messages.Enqueue(h2scan.ErrorMsg)
            _RetVal = False
            h2scan.Close(UUT, True)
            _Messages.Enqueue(h2scan.Results)
            Exit Sub
        End If
        _Messages.Enqueue(h2scan.Results)
        If _stop_test Then Exit Sub

        If Not h2scan.Close(UUT, True) Then
            _Messages.Enqueue(h2scan.Results)
            _Messages.Enqueue(h2scan.ErrorMsg)
            _RetVal = False
            h2scan.Close(UUT, True)
            _Messages.Enqueue(h2scan.Results)
            Exit Sub
        End If
        _Messages.Enqueue(h2scan.Results)
        If _stop_test Then Exit Sub

        ' Dump TM1 logs/records
        _Messages.Enqueue("Dumping " + UUT("SN").Text + " config")
        If Not T.DumpConfig(UUT, True) Then
            _Messages.Enqueue(T.Results)
            _Messages.Enqueue(T.ErrorMsg)
            _RetVal = False
            Exit Sub
        End If
        _Messages.Enqueue(T.Results + vbCr)
        If _stop_test Then Exit Sub

        _Messages.Enqueue("Dumping " + UUT("SN").Text + " rec's")
        If Not T.DumpRecs(UUT, True) Then
            _Messages.Enqueue(T.Results)
            _Messages.Enqueue(T.ErrorMsg)
            _RetVal = False
            Exit Sub
        End If
        _Messages.Enqueue(T.Results)
        If _stop_test Then Exit Sub

        _Messages.Enqueue("Dumping " + UUT("SN").Text + " events")
        If Not T.DumpEvents(UUT, True) Then
            _Messages.Enqueue(T.Results)
            _Messages.Enqueue(T.ErrorMsg)
            _RetVal = False
            Exit Sub
        End If
        _Messages.Enqueue(T.Results)
        If _stop_test Then Exit Sub

        ' clear TM1 logs/records
        _Messages.Enqueue(UUT("SN").Text + " resetting config to defaults")
        If Not C.ConfigFactoryReset(UUT, True) Then
            _Messages.Enqueue(C.Results)
            _Messages.Enqueue(C.ErrorMsg)
            _RetVal = False
            Exit Sub
        End If
        _Messages.Enqueue(C.Results)
        If _stop_test Then Exit Sub

        ' Perform a factory reset.
        CommonLib.Delay(30)
        _Messages.Enqueue("Clearing " + UUT("SN").Text + " logs and records")
        If Not SF.Cmd(UUT, Response, "fr -A", 30, Quiet:=True) Then
            _Messages.Enqueue(Response)
            _Messages.Enqueue("failed sending cmd 'fr -A'")
            _Messages.Enqueue(SF.ErrorMsg)
            _RetVal = False
            Exit Sub
        End If
        _Messages.Enqueue(Response)
        If _stop_test Then Exit Sub

        _Messages.Enqueue("Reconnecting to " + UUT("SN").Text)
        results = SF.Connect(UUT, True)
        _Messages.Enqueue(SF.RtnResults)
        If Not results.PassFail Then
            _Messages.Enqueue("Problem reconnecting to after config reset and reboot")
            _Messages.Enqueue(SF.ErrorMsg)
            _RetVal = False
            Exit Sub
        End If
        If _stop_test Then Exit Sub

        CommonLib.Delay(30)
        _Messages.Enqueue(UUT("SN").Text + " verifying recs cleared")
        If Not SF.Cmd(UUT, Response, "rec", 5, Quiet:=True) Then
            _Messages.Enqueue(Response)
            _Messages.Enqueue("failed sending cmd 'rec'")
            _Messages.Enqueue(SF.ErrorMsg)
            _RetVal = False
            Exit Sub
        End If
        _Messages.Enqueue(Response)
        If _stop_test Then Exit Sub

        If Not Response.Contains("total_recs = 0") Then
            _Messages.Enqueue("recs not cleared")
            _RetVal = False
            Exit Sub
        End If

        ' Verify serial number is still set
        If Not T.GetVersionInfo(UUT, VersionInfo, True) Then
            _Messages.Enqueue(T.Results)
            _Messages.Enqueue(UUT("SN").Text + " problem getting version info")
            _Messages.Enqueue(T.ErrorMsg)
            _RetVal = False
            Exit Sub
        End If
        _Messages.Enqueue(T.Results)
        If _stop_test Then Exit Sub

        If Not VersionInfo("serial number") = UUT("SN").Text Then
            _Messages.Enqueue("Expected 'serial number' = " + UUT("SN").Text)
            _RetVal = False
            Exit Sub
        End If

        _RetVal = True
    End Sub
End Class