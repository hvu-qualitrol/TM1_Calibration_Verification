Imports System.IO
Imports System.IO.File
Imports System.Text.RegularExpressions
Imports System.Threading

Public Class TestReport
    Private _ErrorMsg As String = ""
    Property ErrorMsg() As String
        Get
            Return _ErrorMsg
        End Get
        Set(value As String)

        End Set
    End Property

    'Dim TemplateDoc As String = "825-0074-00 TM1 Test Summary Report RevC.doc"
    'WithEvents TestReportDoc As New WordWrapper

    Public Function ExtractTemplate(ByVal TemplateDoc As String) As Boolean
        Dim res() As String
        Dim template_file As String
        Dim tFile As FileStream
        Dim swriter As Stream
        Dim Success As Boolean = False

        res = Me.GetType.Assembly.GetManifestResourceNames()
        For i = 0 To UBound(res)
            If Not res(i).Contains(TemplateDoc) Then Continue For
            template_file = Regex.Split(res(i), "\w+\.(.*)")(1)
            Try
                tFile = New FileStream("C:\Temp\" + TemplateDoc, FileMode.Create)
            Catch ex As Exception
                _ErrorMsg = "Problem extracting creating C:\Temp\" + TemplateDoc + System.Environment.NewLine
                _ErrorMsg += ex.ToString
                Return False
            End Try

            Try
                swriter = Me.GetType.Assembly.GetManifestResourceStream(res(i))
                For x = 1 To swriter.Length
                    tFile.WriteByte(swriter.ReadByte)
                Next
                swriter.Close()
                tFile.Close()
            Catch ex As Exception
                _ErrorMsg = "Problem writing file C:\Temp\" + TemplateDoc + System.Environment.NewLine
                _ErrorMsg += ex.Message
                Return False
            End Try

            Success = True
        Next
        If Not Success Then
            _ErrorMsg = "Embedded Template not found"
        End If
        Return Success
    End Function

    Public Function CreateTestReport(ByVal UUT As Hashtable, ByVal TestReportData As Hashtable, ByVal TemplateDoc As String) As Boolean
        Try
            Dim success As Boolean = False
            Dim configurationAppSettings As System.Configuration.AppSettingsReader = New System.Configuration.AppSettingsReader
            Dim ReportDir As String = configurationAppSettings.GetValue("DrivePath", GetType(String))
            Dim SN As String = UUT("SN").Text
            Dim TestReportDoc As WordWrapper = Nothing

            ' Try up to three attemps to create a new report
            For attempt As Integer = 1 To 3
                TestReportDoc = New WordWrapper
                If TestReportDoc.CreateWordDoc("C:\Temp\" & TemplateDoc) Then
                    success = True
                    Exit For
                Else
                    Thread.Sleep(50)
                End If
            Next
            If Not success Then Return False

            ' Replace the fields with the cal data
            TestReportDoc.ReplaceItemWith("TM10101YYWWNNNN", SN)
            TestReportDoc.ReplaceItemWith("x.y.zzzz", TestReportData("TM1 Firmware Version"))
            TestReportDoc.ReplaceItemWith("x.yy", TestReportData("Sensor Firmware Version"))
            TestReportDoc.ReplaceItemWith("CAL_DATE", TestReportData("CAL_DATE"))
            TestReportDoc.ReplaceItemWith("<TM8_PPM_1000>", TestReportData("TM8_PPM_1000"))
            TestReportDoc.ReplaceItemWith("<TM8_PPM_6000>", TestReportData("TM8_PPM_6000"))
            TestReportDoc.ReplaceItemWith("<TM8_PPM_10000>", TestReportData("TM8_PPM_10000"))
            TestReportDoc.ReplaceItemWith("<TM1_PPM_1000>", TestReportData("TM1_PPM_1000"))
            TestReportDoc.ReplaceItemWith("<TM1_PPM_6000>", TestReportData("TM1_PPM_6000"))
            TestReportDoc.ReplaceItemWith("<TM1_PPM_10000>", TestReportData("TM1_PPM_10000"))
            TestReportDoc.Finished(ReportDir & "FINAL_TEST\" & SN & "\" & SN & " Test Summary Report.doc")
        Catch ex As Exception
            _ErrorMsg += ex.Message
            Return False
        End Try

        Return True
    End Function
End Class
