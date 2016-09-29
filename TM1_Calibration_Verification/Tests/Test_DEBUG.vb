Imports System.IO

Partial Class Tests
    Public Shared Function Test_DEBUG() As Boolean
        Form1.AppendText("Test 1 enabled = " + Form1.Test_1.Enabled.ToString)
        CommonLib.Delay(60 * 15)
        Return True

        Dim TM8_PPM, TM1_PPM As Integer
        Dim h2scan As New H2SCAN_debug
        Dim Success As Boolean
        Dim ProductInfo As Hashtable
        Dim TargetPPM As Integer
        Dim TestReports As New Hashtable
        Dim TestReportData As Hashtable
        Dim testreport_filepath As String
        Dim TimeStamp As String
        Dim TestReport_File As FileStream
        Dim TestReport_FileWriter As StreamWriter
        Dim Get_PPMS_Success As Boolean

        Success = True
        For Each UUT In UUTs
            If UUT("SN").Text = "" Or UUT("FAILED") Then Continue For

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
                End If
                If Not Get_PPMS_Success Then Exit For
                Form1.AppendText("target = " + TargetPPM.ToString + ", TM1 = " + TM1_PPM.ToString + ", TM8 = " + TM8_PPM.ToString, UUT:=UUT)
                TestReportData("TM1_PPM_" + TargetPPM.ToString) = TM1_PPM
                TestReportData("TM8_PPM_" + TargetPPM.ToString) = TM8_PPM
            Next
            If Not Get_PPMS_Success Then
                UUT("FAILED") = True
                UUT("TAB").ImageIndex = StatusColor.FAILED
                Success = False
                Continue For
            End If

            ' Get Cal date
            If Not h2scan.Open(UUT) Then
                Form1.AppendText(h2scan.ErrorMsg, UUT:=UUT)
                UUT("FAILED") = True
                UUT("TAB").ImageIndex = StatusColor.FAILED
                Success = False
                Continue For
            End If
            If Not h2scan.GetProductInfo(UUT, ProductInfo) Then
                Form1.AppendText("Problem setting " + UUT("SN").Text + " to field operating mode")
                Form1.AppendText(h2scan.ErrorMsg, UUT:=UUT)
                UUT("FAILED") = True
                UUT("TAB").ImageIndex = StatusColor.FAILED
                h2scan.Close(UUT)
                Success = False
                Continue For
            End If
            If Not h2scan.Close(UUT) Then
                Form1.AppendText(h2scan.ErrorMsg, UUT:=UUT)
                UUT("FAILED") = True
                UUT("TAB").ImageIndex = StatusColor.FAILED
                Success = False
                Continue For
            End If
            Form1.AppendText("cal date:  " + ProductInfo("TouchUp"), UUT:=UUT)
            TestReportData("CAL_DATE") = ProductInfo("TouchUp")


            If UUT("SN").Text = "" Or UUT("FAILED") Then Continue For
            TimeStamp = Format(Date.UtcNow, "yyyyMMddHHmmss")
            testreport_filepath = ReportDir + "FINAL_TEST" + "\" + UUT("SN").Text + "\" + UUT("SN").Text + "_TestReport." + TimeStamp + ".txt"
            Try
                TestReport_File = New FileStream(testreport_filepath, FileMode.Create, FileAccess.Write)
                TestReport_FileWriter = New StreamWriter(TestReport_File)
            Catch ex As Exception
                Form1.AppendText("Problem opening " + testreport_filepath + " for writing", UUT:=UUT)
                Form1.AppendText(ex.ToString, UUT:=UUT)
                UUT("FAILED") = True
                UUT("TAB").ImageIndex = StatusColor.FAILED
                Success = False
                Continue For
            End Try
            TestReport_FileWriter.WriteLine("SN = " + UUT("SN").Text)
            For Each TargetPPM In {1000, 6000, 10000}
                TestReport_FileWriter.WriteLine("TARGET PPM = " + TargetPPM.ToString +
                                ", TM8_PPM = " + TestReportData("TM8_PPM_" + TargetPPM.ToString).ToString +
                                ", TM1_PPM = " + TestReportData("TM1_PPM_" + TargetPPM.ToString).ToString)
            Next
            TestReport_FileWriter.WriteLine("CAL_DATE = " + TestReportData("CAL_DATE"))
            TestReport_FileWriter.Close()
            TestReport_File.Close()
        Next

        Return True
    End Function
End Class