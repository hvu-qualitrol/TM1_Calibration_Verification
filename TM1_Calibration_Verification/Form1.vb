Imports System.Text.RegularExpressions
Imports TM1_Calibration_Verification.Tests
Imports System.IO
Imports System.IO.Ports
Imports Microsoft.Office.Interop

Public Class Form1
    Private Test_Sequence As New ArrayList
    Private LogFile As FileStream
    Private LogFileWriter As StreamWriter
    Private ReportFile(20) As FileStream
    Private ReportFileWriter(20) As StreamWriter
    Private LogToDatabase As Boolean = False
    Private DefaultButtonForeColor
    Private DefaultButtonBackColor

    'Private UUTs As New ArrayList
    'Private SN(19) As TextBox

    Sub New()
        Dim configurationAppSettings As System.Configuration.AppSettingsReader = New System.Configuration.AppSettingsReader

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        ReportDir = configurationAppSettings.GetValue("DrivePath", GetType(String))
        FinalReportDir = configurationAppSettings.GetValue("FinalDrivePath", GetType(String))
        Me.Text = "TM1 Calibration and Verification - Version " + Application.ProductVersion.ToString
        Log.FileName = "TM1_Calibration_Verification " + Format(Date.UtcNow, "yyyy-MM-dd-HHmmss")
        Log.WriteLine("Starting Application")
    End Sub

    Private Sub Form1_Load( _
        ByVal sender As System.Object, _
        ByVal e As System.EventArgs) Handles MyBase.Load

        For Each Str As String In System.IO.Ports.SerialPort.GetPortNames()
            GDA_ComportSelect.Items.Add(Str)
        Next

        SetControlArray()
        LoadProductInfo()
        InitializeTestSequence()
        InitializeStatusIcons()
        U1_Time_Select.Text = Timeout_U1.ToString
        U2_Time_Select.Text = Timeout_U2.ToString
    End Sub

    Private Sub LoadProductInfo()
        Dim ProductInfo As Hashtable

        ProductInfo = New Hashtable
        ' ProductInfo.Add("FW_VERSION", "0.9.4862")
        ' ProductInfo.Add("FW_VERSION", "0.9.4884")
        'ProductInfo.Add("FW_VERSION", "1.0.4898")
        'ProductInfo.Add("FW_VERSION", "1.1.5321")
        'ProductInfo.Add("FW_VERSION", "1.1.5056")
        'ProductInfo.Add("FW_VERSION", "1.1.5060")

        ' This is the current production version
        'ProductInfo.Add("sensor firmware version", "3.36B")
        'ProductInfo.Add("sensor firmware checksum", "C2B9")

        ' This is the experimental version under test
        'ProductInfo.Add("sensor firmware version", "3.903B")
        'ProductInfo.Add("sensor firmware checksum", "7286")

        'ProductInfo.Add("hardware version", "0x0")
        'ProductInfo.Add("assembly version", "0x0")

        ' Get version info from the app.config
        Dim configValvue As String
        Dim configurationAppSettings As System.Configuration.AppSettingsReader = New System.Configuration.AppSettingsReader
        Try
            configValvue = Convert.ToString(configurationAppSettings.GetValue("TM1FirmwareVersion", GetType(String)))
            ProductInfo.Add("FW_VERSION", configValvue)
            configValvue = Convert.ToString(configurationAppSettings.GetValue("H2ScanFirmwareVersion", GetType(String)))
            ProductInfo.Add("sensor firmware version", configValvue)
            configValvue = Convert.ToString(configurationAppSettings.GetValue("H2ScanFirmwareCheckSum", GetType(String)))
            ProductInfo.Add("sensor firmware checksum", configValvue)
            configValvue = Convert.ToString(configurationAppSettings.GetValue("HardwareVersion0", GetType(String)))
            ProductInfo.Add("hardware version 0", configValvue)
            configValvue = Convert.ToString(configurationAppSettings.GetValue("AssemblyVersion0", GetType(String)))
            ProductInfo.Add("assembly version 0", configValvue)
            configValvue = Convert.ToString(configurationAppSettings.GetValue("HardwareVersion1", GetType(String)))
            ProductInfo.Add("hardware version 1", configValvue)
            configValvue = Convert.ToString(configurationAppSettings.GetValue("AssemblyVersion1", GetType(String)))
            ProductInfo.Add("assembly version 1", configValvue)
            configValvue = Convert.ToString(configurationAppSettings.GetValue("HardwareVersion2", GetType(String)))
            ProductInfo.Add("hardware version 2", configValvue)
            configValvue = Convert.ToString(configurationAppSettings.GetValue("AssemblyVersion2", GetType(String)))
            ProductInfo.Add("assembly version 2", configValvue)
        Catch ex As Exception
            MessageBox.Show("LoadProductInfo(): Caught exception " + ex.Message)
        End Try
        Products.Add("STANDARD", ProductInfo)
    End Sub

    Sub SetControlArray()
        Dim UUT_Struct As Hashtable

        UUT_Struct = New Hashtable
        'UUT_Struct.Add("POSITION", "LEFT 1")
        UUT_Struct.Add("SN", SN_entry_left_1)
        UUT_Struct.Add("COM", ComPort_1)
        UUT_Struct.Add("SP", Nothing)
        UUT_Struct.Add("RESULT", UUT_Result_1)
        UUT_Struct.Add("FAILED", False)
        UUT_Struct.Add("GV", UUT_Data_1)
        UUT_Struct.Add("DT", Nothing)
        UUT_Struct.Add("LI", Nothing)
        UUT_Struct.Add("TAB", TabPage1)
        UUT_Struct.Add("REPORT", Nothing)
        UUT_Struct.Add("log_filename", "")
        UUT_Struct.Add("CAL_DATE", "")
        UUT_Struct.Add("PROCESS", Nothing)
        UUT_Struct.Add("CLASS_REF", Nothing)
        UUT_Struct.Add("VERIFY_1000_FAILED", False)
        UUT_Struct.Add("VERIFY_6000_FAILED", False)
        UUT_Struct.Add("VERIFY_10000_FAILED", False)
        UUT_Struct.Add("passedVerify", False)
        UUT_Struct.Add("doneRefCyclesAt", 0)
        UUTs.Add(UUT_Struct)

        UUT_Struct = New Hashtable
        UUT_Struct.Add("SN", SN_entry_left_2)
        UUT_Struct.Add("COM", ComPort_2)
        UUT_Struct.Add("SP", Nothing)
        UUT_Struct.Add("RESULT", UUT_Result_2)
        UUT_Struct.Add("FAILED", False)
        UUT_Struct.Add("GV", UUT_Data_2)
        UUT_Struct.Add("DT", Nothing)
        UUT_Struct.Add("LI", Nothing)
        UUT_Struct.Add("TAB", TabPage2)
        UUT_Struct.Add("REPORT", Nothing)
        UUT_Struct.Add("log_filename", "")
        UUT_Struct.Add("CAL_DATE", "")
        UUT_Struct.Add("PROCESS", Nothing)
        UUT_Struct.Add("CLASS_REF", Nothing)
        UUT_Struct.Add("VERIFY_1000_FAILED", False)
        UUT_Struct.Add("VERIFY_6000_FAILED", False)
        UUT_Struct.Add("VERIFY_10000_FAILED", False)
        UUT_Struct.Add("passedVerify", False)
        UUT_Struct.Add("doneRefCyclesAt", 0)
        UUTs.Add(UUT_Struct)

        UUT_Struct = New Hashtable
        UUT_Struct.Add("SN", SN_entry_left_3)
        UUT_Struct.Add("COM", ComPort_3)
        UUT_Struct.Add("SP", Nothing)
        UUT_Struct.Add("RESULT", UUT_Result_3)
        UUT_Struct.Add("FAILED", False)
        UUT_Struct.Add("GV", UUT_Data_3)
        UUT_Struct.Add("DT", Nothing)
        UUT_Struct.Add("LI", Nothing)
        UUT_Struct.Add("TAB", TabPage3)
        UUT_Struct.Add("REPORT", Nothing)
        UUT_Struct.Add("log_filename", "")
        UUT_Struct.Add("CAL_DATE", "")
        UUT_Struct.Add("PROCESS", Nothing)
        UUT_Struct.Add("CLASS_REF", Nothing)
        UUT_Struct.Add("VERIFY_1000_FAILED", False)
        UUT_Struct.Add("VERIFY_6000_FAILED", False)
        UUT_Struct.Add("VERIFY_10000_FAILED", False)
        UUT_Struct.Add("passedVerify", False)
        UUT_Struct.Add("doneRefCyclesAt", 0)
        UUTs.Add(UUT_Struct)

        UUT_Struct = New Hashtable
        UUT_Struct.Add("SN", SN_entry_left_4)
        UUT_Struct.Add("COM", ComPort_4)
        UUT_Struct.Add("SP", Nothing)
        UUT_Struct.Add("RESULT", UUT_Result_4)
        UUT_Struct.Add("FAILED", False)
        UUT_Struct.Add("GV", UUT_Data_4)
        UUT_Struct.Add("DT", Nothing)
        UUT_Struct.Add("LI", Nothing)
        UUT_Struct.Add("TAB", TabPage4)
        UUT_Struct.Add("REPORT", Nothing)
        UUT_Struct.Add("log_filename", "")
        UUT_Struct.Add("CAL_DATE", "")
        UUT_Struct.Add("PROCESS", Nothing)
        UUT_Struct.Add("CLASS_REF", Nothing)
        UUT_Struct.Add("VERIFY_1000_FAILED", False)
        UUT_Struct.Add("VERIFY_6000_FAILED", False)
        UUT_Struct.Add("VERIFY_10000_FAILED", False)
        UUT_Struct.Add("passedVerify", False)
        UUT_Struct.Add("doneRefCyclesAt", 0)
        UUTs.Add(UUT_Struct)

        UUT_Struct = New Hashtable
        UUT_Struct.Add("SN", SN_entry_left_5)
        UUT_Struct.Add("COM", ComPort_5)
        UUT_Struct.Add("SP", Nothing)
        UUT_Struct.Add("RESULT", UUT_Result_5)
        UUT_Struct.Add("FAILED", False)
        UUT_Struct.Add("GV", UUT_Data_5)
        UUT_Struct.Add("DT", Nothing)
        UUT_Struct.Add("LI", Nothing)
        UUT_Struct.Add("TAB", TabPage5)
        UUT_Struct.Add("REPORT", Nothing)
        UUT_Struct.Add("log_filename", "")
        UUT_Struct.Add("CAL_DATE", "")
        UUT_Struct.Add("PROCESS", Nothing)
        UUT_Struct.Add("CLASS_REF", Nothing)
        UUT_Struct.Add("VERIFY_1000_FAILED", False)
        UUT_Struct.Add("VERIFY_6000_FAILED", False)
        UUT_Struct.Add("VERIFY_10000_FAILED", False)
        UUT_Struct.Add("passedVerify", False)
        UUT_Struct.Add("doneRefCyclesAt", 0)
        UUTs.Add(UUT_Struct)

        UUT_Struct = New Hashtable
        UUT_Struct.Add("SN", SN_entry_left_6)
        UUT_Struct.Add("COM", ComPort_6)
        UUT_Struct.Add("SP", Nothing)
        UUT_Struct.Add("RESULT", UUT_Result_6)
        UUT_Struct.Add("FAILED", False)
        UUT_Struct.Add("GV", UUT_Data_6)
        UUT_Struct.Add("DT", Nothing)
        UUT_Struct.Add("LI", Nothing)
        UUT_Struct.Add("TAB", TabPage6)
        UUT_Struct.Add("REPORT", Nothing)
        UUT_Struct.Add("log_filename", "")
        UUT_Struct.Add("CAL_DATE", "")
        UUT_Struct.Add("PROCESS", Nothing)
        UUT_Struct.Add("CLASS_REF", Nothing)
        UUT_Struct.Add("VERIFY_1000_FAILED", False)
        UUT_Struct.Add("VERIFY_6000_FAILED", False)
        UUT_Struct.Add("VERIFY_10000_FAILED", False)
        UUT_Struct.Add("passedVerify", False)
        UUT_Struct.Add("doneRefCyclesAt", 0)
        UUTs.Add(UUT_Struct)

        UUT_Struct = New Hashtable
        UUT_Struct.Add("SN", SN_entry_left_7)
        UUT_Struct.Add("COM", ComPort_7)
        UUT_Struct.Add("SP", Nothing)
        UUT_Struct.Add("RESULT", UUT_Result_7)
        UUT_Struct.Add("FAILED", False)
        UUT_Struct.Add("GV", UUT_Data_7)
        UUT_Struct.Add("DT", Nothing)
        UUT_Struct.Add("LI", Nothing)
        UUT_Struct.Add("TAB", TabPage7)
        UUT_Struct.Add("REPORT", Nothing)
        UUT_Struct.Add("log_filename", "")
        UUT_Struct.Add("CAL_DATE", "")
        UUT_Struct.Add("PROCESS", Nothing)
        UUT_Struct.Add("CLASS_REF", Nothing)
        UUT_Struct.Add("VERIFY_1000_FAILED", False)
        UUT_Struct.Add("VERIFY_6000_FAILED", False)
        UUT_Struct.Add("VERIFY_10000_FAILED", False)
        UUT_Struct.Add("passedVerify", False)
        UUT_Struct.Add("doneRefCyclesAt", 0)
        UUTs.Add(UUT_Struct)

        UUT_Struct = New Hashtable
        UUT_Struct.Add("SN", SN_entry_left_8)
        UUT_Struct.Add("COM", ComPort_8)
        UUT_Struct.Add("SP", Nothing)
        UUT_Struct.Add("RESULT", UUT_Result_8)
        UUT_Struct.Add("FAILED", False)
        UUT_Struct.Add("GV", UUT_Data_8)
        UUT_Struct.Add("DT", Nothing)
        UUT_Struct.Add("LI", Nothing)
        UUT_Struct.Add("TAB", TabPage8)
        UUT_Struct.Add("REPORT", Nothing)
        UUT_Struct.Add("log_filename", "")
        UUT_Struct.Add("CAL_DATE", "")
        UUT_Struct.Add("PROCESS", Nothing)
        UUT_Struct.Add("CLASS_REF", Nothing)
        UUT_Struct.Add("VERIFY_1000_FAILED", False)
        UUT_Struct.Add("VERIFY_6000_FAILED", False)
        UUT_Struct.Add("VERIFY_10000_FAILED", False)
        UUT_Struct.Add("passedVerify", False)
        UUT_Struct.Add("doneRefCyclesAt", 0)
        UUTs.Add(UUT_Struct)

        UUT_Struct = New Hashtable
        UUT_Struct.Add("SN", SN_entry_left_9)
        UUT_Struct.Add("COM", ComPort_9)
        UUT_Struct.Add("SP", Nothing)
        UUT_Struct.Add("RESULT", UUT_Result_9)
        UUT_Struct.Add("FAILED", False)
        UUT_Struct.Add("GV", UUT_Data_9)
        UUT_Struct.Add("DT", Nothing)
        UUT_Struct.Add("LI", Nothing)
        UUT_Struct.Add("TAB", TabPage9)
        UUT_Struct.Add("REPORT", Nothing)
        UUT_Struct.Add("log_filename", "")
        UUT_Struct.Add("CAL_DATE", "")
        UUT_Struct.Add("PROCESS", Nothing)
        UUT_Struct.Add("CLASS_REF", Nothing)
        UUT_Struct.Add("VERIFY_1000_FAILED", False)
        UUT_Struct.Add("VERIFY_6000_FAILED", False)
        UUT_Struct.Add("VERIFY_10000_FAILED", False)
        UUT_Struct.Add("passedVerify", False)
        UUT_Struct.Add("doneRefCyclesAt", 0)
        UUTs.Add(UUT_Struct)

        UUT_Struct = New Hashtable
        UUT_Struct.Add("SN", SN_entry_right_1)
        UUT_Struct.Add("COM", ComPort_11)
        UUT_Struct.Add("SP", Nothing)
        UUT_Struct.Add("RESULT", UUT_Result_11)
        UUT_Struct.Add("FAILED", False)
        UUT_Struct.Add("GV", UUT_Data_11)
        UUT_Struct.Add("DT", Nothing)
        UUT_Struct.Add("LI", Nothing)
        UUT_Struct.Add("TAB", TabPage11)
        UUT_Struct.Add("REPORT", Nothing)
        UUT_Struct.Add("log_filename", "")
        UUT_Struct.Add("CAL_DATE", "")
        UUT_Struct.Add("PROCESS", Nothing)
        UUT_Struct.Add("CLASS_REF", Nothing)
        UUT_Struct.Add("VERIFY_1000_FAILED", False)
        UUT_Struct.Add("VERIFY_6000_FAILED", False)
        UUT_Struct.Add("VERIFY_10000_FAILED", False)
        UUT_Struct.Add("passedVerify", False)
        UUT_Struct.Add("doneRefCyclesAt", 0)
        UUTs.Add(UUT_Struct)

        UUT_Struct = New Hashtable
        UUT_Struct.Add("SN", SN_entry_right_2)
        UUT_Struct.Add("COM", ComPort_12)
        UUT_Struct.Add("SP", Nothing)
        UUT_Struct.Add("RESULT", UUT_Result_12)
        UUT_Struct.Add("FAILED", False)
        UUT_Struct.Add("GV", UUT_Data_12)
        UUT_Struct.Add("DT", Nothing)
        UUT_Struct.Add("LI", Nothing)
        UUT_Struct.Add("TAB", TabPage12)
        UUT_Struct.Add("REPORT", Nothing)
        UUT_Struct.Add("log_filename", "")
        UUT_Struct.Add("CAL_DATE", "")
        UUT_Struct.Add("PROCESS", Nothing)
        UUT_Struct.Add("CLASS_REF", Nothing)
        UUT_Struct.Add("VERIFY_1000_FAILED", False)
        UUT_Struct.Add("VERIFY_6000_FAILED", False)
        UUT_Struct.Add("VERIFY_10000_FAILED", False)
        UUT_Struct.Add("passedVerify", False)
        UUT_Struct.Add("doneRefCyclesAt", 0)
        UUTs.Add(UUT_Struct)

        UUT_Struct = New Hashtable
        UUT_Struct.Add("SN", SN_entry_right_3)
        UUT_Struct.Add("COM", ComPort_13)
        UUT_Struct.Add("SP", Nothing)
        UUT_Struct.Add("RESULT", UUT_Result_13)
        UUT_Struct.Add("FAILED", False)
        UUT_Struct.Add("GV", UUT_Data_13)
        UUT_Struct.Add("DT", Nothing)
        UUT_Struct.Add("LI", Nothing)
        UUT_Struct.Add("TAB", TabPage13)
        UUT_Struct.Add("REPORT", Nothing)
        UUT_Struct.Add("log_filename", "")
        UUT_Struct.Add("CAL_DATE", "")
        UUT_Struct.Add("PROCESS", Nothing)
        UUT_Struct.Add("CLASS_REF", Nothing)
        UUT_Struct.Add("VERIFY_1000_FAILED", False)
        UUT_Struct.Add("VERIFY_6000_FAILED", False)
        UUT_Struct.Add("VERIFY_10000_FAILED", False)
        UUT_Struct.Add("passedVerify", False)
        UUT_Struct.Add("doneRefCyclesAt", 0)
        UUTs.Add(UUT_Struct)

        UUT_Struct = New Hashtable
        UUT_Struct.Add("SN", SN_entry_right_4)
        UUT_Struct.Add("COM", ComPort_14)
        UUT_Struct.Add("SP", Nothing)
        UUT_Struct.Add("RESULT", UUT_Result_14)
        UUT_Struct.Add("FAILED", False)
        UUT_Struct.Add("GV", UUT_Data_14)
        UUT_Struct.Add("DT", Nothing)
        UUT_Struct.Add("LI", Nothing)
        UUT_Struct.Add("TAB", TabPage14)
        UUT_Struct.Add("REPORT", Nothing)
        UUT_Struct.Add("log_filename", "")
        UUT_Struct.Add("CAL_DATE", "")
        UUT_Struct.Add("PROCESS", Nothing)
        UUT_Struct.Add("CLASS_REF", Nothing)
        UUT_Struct.Add("VERIFY_1000_FAILED", False)
        UUT_Struct.Add("VERIFY_6000_FAILED", False)
        UUT_Struct.Add("VERIFY_10000_FAILED", False)
        UUT_Struct.Add("passedVerify", False)
        UUT_Struct.Add("doneRefCyclesAt", 0)
        UUTs.Add(UUT_Struct)

        UUT_Struct = New Hashtable
        UUT_Struct.Add("SN", SN_entry_right_5)
        UUT_Struct.Add("COM", ComPort_15)
        UUT_Struct.Add("SP", Nothing)
        UUT_Struct.Add("RESULT", UUT_Result_15)
        UUT_Struct.Add("FAILED", False)
        UUT_Struct.Add("GV", UUT_Data_15)
        UUT_Struct.Add("DT", Nothing)
        UUT_Struct.Add("LI", Nothing)
        UUT_Struct.Add("TAB", TabPage15)
        UUT_Struct.Add("REPORT", Nothing)
        UUT_Struct.Add("log_filename", "")
        UUT_Struct.Add("CAL_DATE", "")
        UUT_Struct.Add("PROCESS", Nothing)
        UUT_Struct.Add("CLASS_REF", Nothing)
        UUT_Struct.Add("VERIFY_1000_FAILED", False)
        UUT_Struct.Add("VERIFY_6000_FAILED", False)
        UUT_Struct.Add("VERIFY_10000_FAILED", False)
        UUT_Struct.Add("passedVerify", False)
        UUT_Struct.Add("doneRefCyclesAt", 0)
        UUTs.Add(UUT_Struct)

        UUT_Struct = New Hashtable
        UUT_Struct.Add("SN", SN_entry_right_6)
        UUT_Struct.Add("COM", ComPort_16)
        UUT_Struct.Add("SP", Nothing)
        UUT_Struct.Add("RESULT", UUT_Result_16)
        UUT_Struct.Add("FAILED", False)
        UUT_Struct.Add("GV", UUT_Data_16)
        UUT_Struct.Add("DT", Nothing)
        UUT_Struct.Add("LI", Nothing)
        UUT_Struct.Add("TAB", TabPage16)
        UUT_Struct.Add("REPORT", Nothing)
        UUT_Struct.Add("log_filename", "")
        UUT_Struct.Add("CAL_DATE", "")
        UUT_Struct.Add("PROCESS", Nothing)
        UUT_Struct.Add("CLASS_REF", Nothing)
        UUT_Struct.Add("VERIFY_1000_FAILED", False)
        UUT_Struct.Add("VERIFY_6000_FAILED", False)
        UUT_Struct.Add("VERIFY_10000_FAILED", False)
        UUT_Struct.Add("passedVerify", False)
        UUT_Struct.Add("doneRefCyclesAt", 0)
        UUTs.Add(UUT_Struct)

        UUT_Struct = New Hashtable
        UUT_Struct.Add("SN", SN_entry_right_7)
        UUT_Struct.Add("COM", ComPort_17)
        UUT_Struct.Add("SP", Nothing)
        UUT_Struct.Add("RESULT", UUT_Result_17)
        UUT_Struct.Add("FAILED", False)
        UUT_Struct.Add("GV", UUT_Data_17)
        UUT_Struct.Add("DT", Nothing)
        UUT_Struct.Add("LI", Nothing)
        UUT_Struct.Add("TAB", TabPage17)
        UUT_Struct.Add("REPORT", Nothing)
        UUT_Struct.Add("log_filename", "")
        UUT_Struct.Add("CAL_DATE", "")
        UUT_Struct.Add("PROCESS", Nothing)
        UUT_Struct.Add("CLASS_REF", Nothing)
        UUT_Struct.Add("VERIFY_1000_FAILED", False)
        UUT_Struct.Add("VERIFY_6000_FAILED", False)
        UUT_Struct.Add("VERIFY_10000_FAILED", False)
        UUT_Struct.Add("passedVerify", False)
        UUT_Struct.Add("doneRefCyclesAt", 0)
        UUTs.Add(UUT_Struct)

        UUT_Struct = New Hashtable
        UUT_Struct.Add("SN", SN_entry_right_8)
        UUT_Struct.Add("COM", ComPort_18)
        UUT_Struct.Add("SP", Nothing)
        UUT_Struct.Add("RESULT", UUT_Result_18)
        UUT_Struct.Add("FAILED", False)
        UUT_Struct.Add("GV", UUT_Data_18)
        UUT_Struct.Add("DT", Nothing)
        UUT_Struct.Add("LI", Nothing)
        UUT_Struct.Add("TAB", TabPage18)
        UUT_Struct.Add("REPORT", Nothing)
        UUT_Struct.Add("log_filename", "")
        UUT_Struct.Add("CAL_DATE", "")
        UUT_Struct.Add("PROCESS", Nothing)
        UUT_Struct.Add("CLASS_REF", Nothing)
        UUT_Struct.Add("VERIFY_1000_FAILED", False)
        UUT_Struct.Add("VERIFY_6000_FAILED", False)
        UUT_Struct.Add("VERIFY_10000_FAILED", False)
        UUT_Struct.Add("passedVerify", False)
        UUT_Struct.Add("doneRefCyclesAt", 0)
        UUTs.Add(UUT_Struct)

        UUT_Struct = New Hashtable
        UUT_Struct.Add("SN", SN_entry_right_9)
        UUT_Struct.Add("COM", ComPort_19)
        UUT_Struct.Add("SP", Nothing)
        UUT_Struct.Add("RESULT", UUT_Result_19)
        UUT_Struct.Add("FAILED", False)
        UUT_Struct.Add("GV", UUT_Data_19)
        UUT_Struct.Add("DT", Nothing)
        UUT_Struct.Add("LI", Nothing)
        UUT_Struct.Add("TAB", TabPage19)
        UUT_Struct.Add("REPORT", Nothing)
        UUT_Struct.Add("log_filename", "")
        UUT_Struct.Add("CAL_DATE", "")
        UUT_Struct.Add("PROCESS", Nothing)
        UUT_Struct.Add("CLASS_REF", Nothing)
        UUT_Struct.Add("VERIFY_1000_FAILED", False)
        UUT_Struct.Add("VERIFY_6000_FAILED", False)
        UUT_Struct.Add("VERIFY_10000_FAILED", False)
        UUT_Struct.Add("passedVerify", False)
        UUT_Struct.Add("doneRefCyclesAt", 0)
        UUTs.Add(UUT_Struct)

    End Sub

    Sub InitializeTestSequence_Partial()
        'Sub InitializeTestSequence_Partial()
        Dim TS As Test_Item

        For i = 0 To Test_Sequence.Count - 1
            Test_Sequence(i).Button.Visible = False
        Next
        Test_Sequence.Clear()

        TS = New Test_Item
        Test_Sequence.Add(TS)
        Test_Sequence(0).Button = Test_1
        Test_Sequence(0).Button.Text = "INITIAL"
        Test_Sequence(0).Handler = New MyDelFun(AddressOf Test_INITIAL)
        Test_Sequence(0).Enabled = True
        Test_Sequence(0).Timeout = 0.1

        TS = New Test_Item
        Test_Sequence.Add(TS)
        Test_Sequence(1).Button = Test_2
        Test_Sequence(1).Button.Text = "FINAL"
        Test_Sequence(1).Handler = New MyDelFun(AddressOf Test_FINAL)
        Test_Sequence(1).Enabled = True
        Test_Sequence(1).Timeout = 0.3

        For i = 0 To Test_Sequence.Count - 1
            Test_Sequence(i).Button.Visible = True
            Test_Sequence(i).Button.BackColor = Color.PaleTurquoise
            Test_Sequence(i).Button.ForeColor = Color.Black
        Next

        For i = 0 To Test_Sequence.Count - 1
            Test_Sequence(i).Button.Visible = True
            Test_Sequence(i).Button.BackColor = Color.PaleTurquoise
            Test_Sequence(i).Button.ForeColor = Color.Black
        Next
    End Sub

    Sub InitializeTestSequence()
        'Sub InitializeTestSequence_Current()
        Dim TS As Test_Item

        For i = 0 To Test_Sequence.Count - 1
            Test_Sequence(i).Button.Visible = False
        Next
        Test_Sequence.Clear()

        TS = New Test_Item
        Test_Sequence.Add(TS)
        Test_Sequence(0).Button = Test_1
        Test_Sequence(0).Button.Text = "INITIAL"
        Test_Sequence(0).Handler = New MyDelFun(AddressOf Test_INITIAL)
        Test_Sequence(0).Enabled = True
        Test_Sequence(0).Timeout = 0.1

        TS = New Test_Item
        Test_Sequence.Add(TS)
        Test_Sequence(1).Button = Test_2
        Test_Sequence(1).Button.Text = "GDA_1000"
        'Test_Sequence(1).Handler = New MyDelFun(AddressOf SkipTest)
        Test_Sequence(1).Handler = New MyDelFun(AddressOf Test_GDA_1000)
        Test_Sequence(1).Enabled = True
        Test_Sequence(1).Timeout = 1.5

        TS = New Test_Item
        Test_Sequence.Add(TS)
        Test_Sequence(2).Button = Test_3
        Test_Sequence(2).Button.Text = "CLEAR_CAL"
        'Test_Sequence(2).Handler = New MyDelFun(AddressOf SkipTest)
        Test_Sequence(2).Handler = New MyDelFun(AddressOf Test_CLEAR_CAL)
        Test_Sequence(2).Enabled = True
        Test_Sequence(2).Timeout = 0.1

        TS = New Test_Item
        Test_Sequence.Add(TS)
        Test_Sequence(3).Button = Test_4
        Test_Sequence(3).Button.Text = "U0"
        'Test_Sequence(3).Handler = New MyDelFun(AddressOf SkipTest)
        Test_Sequence(3).Handler = New MyDelFun(AddressOf Test_U0)
        Test_Sequence(3).Enabled = True
        Test_Sequence(3).Timeout = 0.2

        'TS = New Test_Item
        'Test_Sequence.Add(TS)
        'Test_Sequence(3).Button = Test_4
        'Test_Sequence(3).Button.Text = "GDA_1000"
        ''Test_Sequence(3).Handler = New MyDelFun(AddressOf SkipTest)
        'Test_Sequence(3).Handler = New MyDelFun(AddressOf Test_GDA_1000)
        'Test_Sequence(3).Enabled = True
        'Test_Sequence(3).Timeout = 1.5

        TS = New Test_Item
        Test_Sequence.Add(TS)
        Test_Sequence(4).Button = Test_5
        Test_Sequence(4).Button.Text = "U1"
        'Test_Sequence(4).Handler = New MyDelFun(AddressOf SkipTest)
        Test_Sequence(4).Handler = New MyDelFun(AddressOf Test_U1)
        Test_Sequence(4).Enabled = True
        Test_Sequence(4).Timeout = Timeout_U1

        TS = New Test_Item
        Test_Sequence.Add(TS)
        Test_Sequence(5).Button = Test_6
        Test_Sequence(5).Button.Text = "GDA_10000"
        'Test_Sequence(5).Handler = New MyDelFun(AddressOf SkipTest)
        Test_Sequence(5).Handler = New MyDelFun(AddressOf Test_GDA_10000)
        Test_Sequence(5).Enabled = True
        Test_Sequence(5).Timeout = 2.0

        TS = New Test_Item
        Test_Sequence.Add(TS)
        Test_Sequence(6).Button = Test_7
        Test_Sequence(6).Button.Text = "U2"
        'Test_Sequence(6).Handler = New MyDelFun(AddressOf SkipTest)
        Test_Sequence(6).Handler = New MyDelFun(AddressOf Test_U2)
        Test_Sequence(6).Enabled = True
        Test_Sequence(6).Timeout = Timeout_U2

        TS = New Test_Item
        Test_Sequence.Add(TS)
        Test_Sequence(7).Button = Test_8
        Test_Sequence(7).Button.Text = "VER_10000"
        'Test_Sequence(7).Handler = New MyDelFun(AddressOf SkipTest)
        Test_Sequence(7).Handler = New MyDelFun(AddressOf Test_VER_10000)
        Test_Sequence(7).Enabled = True
        Test_Sequence(7).Timeout = Timeout_VER_10000

        TS = New Test_Item
        Test_Sequence.Add(TS)
        Test_Sequence(8).Button = Test_9
        Test_Sequence(8).Button.Text = "GDA_6000"
        'Test_Sequence(8).Handler = New MyDelFun(AddressOf SkipTest)
        Test_Sequence(8).Handler = New MyDelFun(AddressOf Test_GDA_6000)
        Test_Sequence(8).Enabled = True
        Test_Sequence(8).Timeout = 2.0

        TS = New Test_Item
        Test_Sequence.Add(TS)
        Test_Sequence(9).Button = Test_10
        Test_Sequence(9).Button.Text = "VER_6000"
        'Test_Sequence(9).Handler = New MyDelFun(AddressOf SkipTest)
        Test_Sequence(9).Handler = New MyDelFun(AddressOf Test_VER_6000)
        Test_Sequence(9).Enabled = True
        Test_Sequence(9).Timeout = Timeout_VER_6000

        TS = New Test_Item
        Test_Sequence.Add(TS)
        Test_Sequence(10).Button = Test_11
        Test_Sequence(10).Button.Text = "GDA_1000_V"
        Test_Sequence(10).Handler = New MyDelFun(AddressOf Test_GDA_1000)
        Test_Sequence(10).Enabled = True
        Test_Sequence(10).Timeout = 2.0

        TS = New Test_Item
        Test_Sequence.Add(TS)
        Test_Sequence(11).Button = Test_12
        Test_Sequence(11).Button.Text = "VER_1000"
        Test_Sequence(11).Handler = New MyDelFun(AddressOf Test_VER_1000)
        Test_Sequence(11).Enabled = True
        Test_Sequence(11).Timeout = Timeout_VER_1000

        TS = New Test_Item
        Test_Sequence.Add(TS)
        Test_Sequence(12).Button = Test_13
        Test_Sequence(12).Button.Text = "FINAL"
        Test_Sequence(12).Handler = New MyDelFun(AddressOf Test_FINAL)
        Test_Sequence(12).Enabled = True
        Test_Sequence(12).Timeout = 0.3

        For i = 0 To Test_Sequence.Count - 1
            Test_Sequence(i).Button.Visible = True
            Test_Sequence(i).Button.BackColor = Color.PaleTurquoise
            Test_Sequence(i).Button.ForeColor = Color.Black
        Next
    End Sub

    'Public Sub InfoInit_Partial()
    Public Sub InfoInit_Partial()
        Dim DebugMode As Boolean = False
        Dim RunSingleTest As Boolean = False
        Dim Test_Index As Integer
        Dim Retry As Boolean
        Dim RetryCnt As Integer
        Dim Test_Status As Boolean
        Dim TestName As String
        Dim TestTime As Integer = 0
        Dim TestStartTime As DateTime = Now
        Dim ReportFilePath As String
        Dim FinalFilePath As String
        Dim TimeStamp As String
        Dim LogFilePath As String
        Dim AllPassed As Boolean
        Dim ReportFileIndex As Integer
        Dim AllFailed As Boolean
        Dim TimeToCompletion As TimeSpan

        Log.WriteLine("Test Started")

        'If Test_Sequence(4).Enabled Then
        '    Test_Sequence(4).Timeout = Timeout_U1
        'End If
        'If Test_Sequence(6).Enabled Then
        '    Test_Sequence(6).Timeout = Timeout_U2
        'End If

        Log.WriteLine("Creating Calibration log file")
        TimeStamp = Format(Date.UtcNow, "yyyyMMddHHmmss")
        LogFilePath = "C:\Temp\TM1_Calibration_Verification_log_file" + TimeStamp + ".txt"
        Try
            LogFile = New FileStream(LogFilePath, FileMode.Create, FileAccess.Write)
            LogFileWriter = New StreamWriter(LogFile)
        Catch ex As Exception
            MsgBox("Problem opening log file " + LogFilePath)
            Exit Sub
        End Try

        Stopped = False
        If ListBox_TestMode.Text = "Debug" Then
            DebugMode = True
        End If

        Log.WriteLine("Clearing Results")
        Results.Clear()
        TestStatus.Text = "RUNNING"
        TestStatus.BackColor = Color.Yellow
        TestStatus.ForeColor = Color.Black


        TestStatus.Text = "RUNNING"
        TestStatus.BackColor = Color.Yellow
        TestStatus.ForeColor = Color.Black

        ' Rack 1
        'Dim sn As String() = {
        '"TM1010113160267",
        '"TM1010113160300",
        '"TM1010113160301",
        '"TM1010113160303",
        '"TM1010113160304",
        '"TM1010113160305",
        '"TM1010113160306",
        '"TM1010113160307",
        '"TM1010113160320",
        '"TM1010113160321",
        '"TM1010113160322",
        '"TM1010113160323",
        '"TM1010113160324",
        '"TM1010113160325",
        '"TM1010113160326",
        '"TM1010113160327",
        '"TM1010113160329"
        '}

        ' Rack 2
        '        Dim sn As String() = {
        '"TM1010113160302",
        '"TM1010113160328",
        '"TM1010113160330",
        '"TM1010113160331",
        '"TM1010113160332",
        '"TM1010113160333",
        '"TM1010113160334",
        '"TM1010113160335",
        '"TM1010113160336",
        '"TM1010113160337",
        '"TM1010113160338",
        '"TM1010113160339",
        '"TM1010113160340",
        '"TM1010113160341",
        '"TM1010113160343",
        '"TM1010113160300"
        '        }

        'Dim sn As String() = {"TM1010113160300", "TM1010113160329"}

        'Dim i As Integer = 0
        'For Each s As String In sn
        '    UUTs(i)("SN").Text = sn(i)
        '    i += 1
        'Next

        ReportFileIndex = 0
        For Each UUT In UUTs
            If Not UUT("SN").Text = "" Then
                Log.WriteLine("Creating report for " + UUT("SN").Text)
                UUT("TAB").ImageIndex = StatusColor.RUNNING
                ReportFilePath = ReportDir + "FINAL_TEST" + "\" + UUT("SN").Text
                Try
                    Directory.CreateDirectory(ReportFilePath)
                Catch ex As Exception
                    AppendText("Problem creating directory " + ReportFilePath)
                    Fail()
                End Try

                FinalFilePath = FinalReportDir + "FINAL_TEST" + "\" + UUT("SN").Text
                Try
                    Directory.CreateDirectory(FinalFilePath)
                Catch ex As Exception
                    AppendText("Problem creating directory " + FinalFilePath)
                    Fail()
                End Try

                Try
                    ReportFile(ReportFileIndex) = New FileStream(ReportFilePath + "\" + TimeStamp + ".txt", FileMode.Create, FileAccess.Write)
                    ReportFileWriter(ReportFileIndex) = New StreamWriter(ReportFile(ReportFileIndex))
                    UUT("REPORT") = ReportFileWriter(ReportFileIndex)
                Catch ex As Exception
                    AppendText("Problem opening log file")
                    AppendText(ex.ToString)
                    Fail()
                End Try
                UUT("log_filename") = TimeStamp + ".txt"
                ReportFileIndex += 1
            Else
                UUT("TAB").ImageIndex = StatusColor.NONE
            End If
        Next
        Application.DoEvents()

        If Stopped Then
            Fail()
            Exit Sub
        End If
        If Not DebugMode Then
            LogToDatabase = True
        End If
    End Sub

    'Public Sub InfoInit_Current()
    Public Sub InfoInit()
        Dim DebugMode As Boolean = False
        Dim RunSingleTest As Boolean = False
        Dim Test_Index As Integer
        Dim Retry As Boolean
        Dim RetryCnt As Integer
        Dim Test_Status As Boolean
        Dim TestName As String
        Dim TestTime As Integer = 0
        Dim TestStartTime As DateTime = Now
        Dim ReportFilePath As String
        Dim FinalFilePath As String
        Dim TimeStamp As String
        Dim LogFilePath As String
        Dim AllPassed As Boolean
        Dim ReportFileIndex As Integer
        Dim AllFailed As Boolean
        Dim TimeToCompletion As TimeSpan

        Log.WriteLine("Test Started")

        If Test_Sequence(4).Enabled Then
            Test_Sequence(4).Timeout = Timeout_U1
        End If
        If Test_Sequence(6).Enabled Then
            Test_Sequence(6).Timeout = Timeout_U2
        End If

        Log.WriteLine("Creating Calibration log file")
        TimeStamp = Format(Date.UtcNow, "yyyyMMddHHmmss")
        LogFilePath = "C:\Temp\TM1_Calibration_Verification_log_file" + TimeStamp + ".txt"
        Try
            LogFile = New FileStream(LogFilePath, FileMode.Create, FileAccess.Write)
            LogFileWriter = New StreamWriter(LogFile)
        Catch ex As Exception
            MsgBox("Problem opening log file " + LogFilePath)
            Exit Sub
        End Try

        Stopped = False
        If ListBox_TestMode.Text = "Debug" Then
            DebugMode = True
        End If

        Log.WriteLine("Clearing Results")
        Results.Clear()
        TestStatus.Text = "RUNNING"
        TestStatus.BackColor = Color.Yellow
        TestStatus.ForeColor = Color.Black


        TestStatus.Text = "RUNNING"
        TestStatus.BackColor = Color.Yellow
        TestStatus.ForeColor = Color.Black

        ReportFileIndex = 0
        For Each UUT In UUTs
            If Not UUT("SN").Text = "" Then
                Log.WriteLine("Creating report for " + UUT("SN").Text)
                UUT("TAB").ImageIndex = StatusColor.RUNNING
                ReportFilePath = ReportDir + "FINAL_TEST" + "\" + UUT("SN").Text
                Try
                    Directory.CreateDirectory(ReportFilePath)
                Catch ex As Exception
                    AppendText("Problem creating directory " + ReportFilePath)
                    Fail()
                End Try

                FinalFilePath = FinalReportDir + "FINAL_TEST" + "\" + UUT("SN").Text
                Try
                    Directory.CreateDirectory(FinalFilePath)
                Catch ex As Exception
                    AppendText("Problem creating directory " + FinalFilePath)
                    Fail()
                End Try

                Try
                    ReportFile(ReportFileIndex) = New FileStream(ReportFilePath + "\" + TimeStamp + ".txt", FileMode.Create, FileAccess.Write)
                    ReportFileWriter(ReportFileIndex) = New StreamWriter(ReportFile(ReportFileIndex))
                    UUT("REPORT") = ReportFileWriter(ReportFileIndex)
                Catch ex As Exception
                    AppendText("Problem opening log file")
                    AppendText(ex.ToString)
                    Fail()
                End Try
                UUT("log_filename") = TimeStamp + ".txt"
                ReportFileIndex += 1
            Else
                UUT("TAB").ImageIndex = StatusColor.NONE
            End If
        Next
        Application.DoEvents()

        If Stopped Then
            Fail()
            Exit Sub
        End If
        If Not DebugMode Then
            LogToDatabase = True
        End If
    End Sub

    Private Sub ResetFailFlags()
        For Each UUT In UUTs
            If Not UUT("SN").Text = "" Then
                UUT("FAILED") = False
                UUT("TM1 Firmware Version") = Products("STANDARD")("FW_VERSION")
                UUT("sensor firmware version") = Products("STANDARD")("sensor firmware version")
            End If
        Next
    End Sub

    Private hasStarted As Boolean = False
    ' Private Sub StartTest_Click()
    Private Sub StartTest_Click(sender As System.Object, e As System.EventArgs) Handles StartTest.Click, Test_1.Click,
        Test_2.Click, Test_3.Click, Test_4.Click, Test_5.Click, Test_6.Click, Test_7.Click, Test_8.Click, Test_9.Click,
        Test_10.Click, Test_11.Click, Test_12.Click, Test_13.Click

        Dim DebugMode As Boolean = False
        Dim RunSingleTest As Boolean = False
        Dim Test_Index As Integer
        Dim Retry As Boolean
        Dim RetryCnt As Integer
        Dim Test_Status As Boolean
        Dim TestName As String
        Dim TestTime As Integer = 0
        Dim TestStartTime As DateTime = Now
        Dim ReportFilePath As String
        Dim FinalFilePath As String
        Dim TimeStamp As String
        Dim LogFilePath As String
        Dim AllPassed As Boolean
        Dim ReportFileIndex As Integer
        Dim AllFailed As Boolean
        Dim TimeToCompletion As TimeSpan

        Log.WriteLine("Test Started")

        ' This reset fail flag to enable test rerun
        If hasStarted Then ResetFailFlags()
        hasStarted = True

        If Test_Sequence(4).Enabled Then
            Test_Sequence(4).Timeout = Timeout_U1
        End If
        If Test_Sequence(6).Enabled Then
            Test_Sequence(6).Timeout = Timeout_U2
        End If

        Log.WriteLine("Creating Calibration log file")
        TimeStamp = Format(Date.UtcNow, "yyyyMMddHHmmss")
        LogFilePath = "C:\Temp\TM1_Calibration_Verification_log_file" + TimeStamp + ".txt"
        Try
            LogFile = New FileStream(LogFilePath, FileMode.Create, FileAccess.Write)
            LogFileWriter = New StreamWriter(LogFile)
        Catch ex As Exception
            MsgBox("Problem opening log file " + LogFilePath)
            Exit Sub
        End Try

        Stopped = False
        If ListBox_TestMode.Text = "Debug" Then
            DebugMode = True
        End If

        If Not sender.Equals(StartTest) Then
            RunSingleTest = True
            If Not DebugMode Then
                Stopped = True
                Exit Sub
            End If
        End If

        Log.WriteLine("Clearing Results")
        Results.Clear()
        TestStatus.Text = "RUNNING"
        TestStatus.BackColor = Color.Yellow
        TestStatus.ForeColor = Color.Black

        If RunSingleTest Then
            For i = 0 To Test_Sequence.Count - 1
                If sender.Equals(Test_Sequence(i).Button) Then
                    Test_Index = i
                End If
            Next
            If Test_Index < 0 Then
                AppendText("Could not find test for " + sender.Text)
                'For i = 0 To Test_Sequence.Count - 1
                '    Test_Sequence(i).Button.Enabled = True
                'Next
                Fail()
                Exit Sub
            End If
        End If

        For i = 0 To Test_Sequence.Count - 1
            If RunSingleTest And Not i = Test_Index Then
                Continue For
            End If
            'If Not Test_Sequence(i).Button.Enabled Then
            If Not Test_Sequence(i).Enabled Then
                Continue For
            End If
            'If Not CheckEntriesForTest(Test_Sequence(i)) Then
            '    Fail()
            '    Exit Sub
            'End If
        Next
        DisableEntries()

        For i = 0 To Test_Sequence.Count - 1
            If Test_Sequence(i).Enabled Then
                Test_Sequence(i).Button.BackColor = Color.PaleTurquoise
                Test_Sequence(i).Button.ForeColor = Color.Black
            End If
            'If ListBox_TestMode.Text = "Debug" Then
            '    If Test_Sequence(i).Button.Enabled Then
            '        Test_Sequence(i).Button.BackColor = Color.LightGray
            '        Test_Sequence(i).Button.ForeColor = Color.Black
            '    End If
            'Else
            '    Test_Sequence(i).Button.Enabled = True
            '    Test_Sequence(i).Button.ForeColor = DefaultForeColor
            '    Test_Sequence(i).Button.BackColor = DefaultBackColor
            'End If
        Next

        TestStatus.Text = "RUNNING"
        TestStatus.BackColor = Color.Yellow
        TestStatus.ForeColor = Color.Black

        ReportFileIndex = 0
        For Each UUT In UUTs
            If Not UUT("SN").Text = "" Then
                Log.WriteLine("Creating report for " + UUT("SN").Text)
                UUT("TAB").ImageIndex = StatusColor.RUNNING
                ReportFilePath = ReportDir + "FINAL_TEST" + "\" + UUT("SN").Text
                Try
                    Directory.CreateDirectory(ReportFilePath)
                Catch ex As Exception
                    AppendText("Problem creating directory " + ReportFilePath)
                    Fail()
                End Try

                FinalFilePath = FinalReportDir + "FINAL_TEST" + "\" + UUT("SN").Text
                Try
                    Directory.CreateDirectory(FinalFilePath)
                Catch ex As Exception
                    AppendText("Problem creating directory " + FinalFilePath)
                    Fail()
                End Try

                Try
                    ReportFile(ReportFileIndex) = New FileStream(ReportFilePath + "\" + TimeStamp + ".txt", FileMode.Create, FileAccess.Write)
                    ReportFileWriter(ReportFileIndex) = New StreamWriter(ReportFile(ReportFileIndex))
                    UUT("REPORT") = ReportFileWriter(ReportFileIndex)
                Catch ex As Exception
                    AppendText("Problem opening log file")
                    AppendText(ex.ToString)
                    Fail()
                End Try
                UUT("log_filename") = TimeStamp + ".txt"
                ReportFileIndex += 1
            Else
                UUT("TAB").ImageIndex = StatusColor.NONE
            End If
        Next
        Application.DoEvents()

        If Stopped Then
            Fail()
            Exit Sub
        End If
        If Not DebugMode Then
            LogToDatabase = True
        End If

        ' If this is a partial run, then it needs to run initialization parts
        If PartialRun Then
            Test_INITIAL()
        End If

        For i = 0 To Test_Sequence.Count - 1
            If RunSingleTest And Not i = Test_Index Then
                Continue For
            End If
            'If Not Test_Sequence(i).Button.Enabled = True Then
            If Not Test_Sequence(i).Enabled = True Then
                Continue For
            End If

            ETA_Timer.Stop()
            EstimatedTestCompletionHours = 0
            If RunSingleTest Then
                EstimatedTestCompletionHours = Test_Sequence(i).Timeout
            End If
            For j = i To Test_Sequence.Count - 1
                If Test_Sequence(j).Enabled Then
                    EstimatedTestCompletionHours += Test_Sequence(j).Timeout
                End If
            Next
            TimeToCompletion = New TimeSpan(0, EstimatedTestCompletionHours * 60, 0)
            EstimatedTestCompletionDate = Now.Add(TimeToCompletion)
            EstimatedCompletionHours.Text = Math.Round(EstimatedTestCompletionHours, 1).ToString + " hours"
            EstimatedCompletionDate.Text = EstimatedTestCompletionDate.ToString
            ETA_Timer.Start()

            Test_Sequence(i).Button.BackColor = Color.Yellow
            Test_Sequence(i).Button.ForeColor = Color.Black
            RetryCnt = 0
            Retry = True
            While Retry
                AppendText("##########################################################", LogAllUuts:=True)
                If (RetryCnt = 0) Then
                    AppendText("TEST_START:  " + Test_Sequence(i).Button.Text, LogAllUuts:=True)
                Else
                    AppendText("TEST_RETRY:  " + Test_Sequence(i).Button.Text, LogAllUuts:=True)
                End If
                AppendText("START_TIME:  " + Format(Date.UtcNow, "yyyyMMddHHmmss"), LogAllUuts:=True)
                RetryCnt += 1
                TestName = Test_Sequence(i).Button.Text
                Log.WriteLine("Doing sequence " + TestName)
                Test_Status = Test_Sequence(i).Handler.Invoke()
                Log.WriteLine("Sequence complete")
                AllFailed = True
                For Each UUT In UUTs
                    If UUT("SN").Text = "" Then Continue For
                    If Not UUT("FAILED") Then
                        AllFailed = False
                    End If
                Next
                If AllFailed Then Test_Status = False
                TestTime = Now.Subtract(TestStartTime).TotalSeconds
                AppendText("END_TIME:  " + Format(Date.UtcNow, "yyyyMMddHHmmss"), LogAllUuts:=True)
                Retry = False
                If Not Test_Status Then
                    AppendText("TEST_STATUS:  FAILED", LogAllUuts:=True)
                    If RetriesAllowed And Not Stopped Then
                        If (MessageBox.Show("Test failed, click yes to retry", "RETRY?", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes) Then
                            Retry = True
                        End If
                    End If
                    If Not Retry Then
                        If Stopped Then
                            Test_Sequence(i).Button.BackColor = Color.Blue
                            Test_Sequence(i).Button.ForeColor = Color.White
                        Else
                            Test_Sequence(i).Button.BackColor = Color.Red
                            Test_Sequence(i).Button.ForeColor = Color.White
                        End If
                        Fail(TestName, TestTime)
                        Exit Sub
                    End If
                End If
            End While
            'AppendText("TEST_STATUS:  PASSED")
            Test_Sequence(i).Button.ForeColor = Color.Black
            AllPassed = True
            For Each UUT In UUTs
                If UUT("SN").Text = "" Then Continue For
                If UUT("FAILED") Then
                    AllPassed = False
                End If
            Next
            If AllPassed Then
                AppendText("TEST_STATUS:  PASSED", LogAllUuts:=True)
                Test_Sequence(i).Button.BackColor = Color.LightGreen
            Else
                AppendText("TEST_STATUS:  SOME_PASSED")
                Test_Sequence(i).Button.BackColor = Color.PaleGoldenrod
            End If

            ' Refresh the form
            Application.DoEvents()
        Next

        If AllPassed Then
            AppendText("FINAL_STATUS:  PASSED", LogAllUuts:=True)
        Else
            AppendText("FINAL_STATUS:  SOME_PASSED")
        End If
        Pass(TestTime, AllPassed)
        LogToDatabase = False
        PartialRun = False

    End Sub

    Private Sub TM8_SN_KeyDown(sender As System.Object, e As System.Windows.Forms.KeyEventArgs) Handles TM8_SN.KeyDown
        Dim TM8_SN_valid As Boolean = False
        Dim Ping_Success As Boolean

        If e.KeyCode = Keys.Enter Then
            If Regex.IsMatch(TM8_SN.Text, "^TM8\d\d\d\d\d\d$") Or Regex.IsMatch(TM8_SN.Text, "^(10|172|192)\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}$") Then
                TM8_SN_valid = True
                TM8_SN.BackColor = Color.LightGreen
            Else
                TM8_SN.BackColor = Color.Red
                MsgBox("'" + TM8_SN.Text + "' is not a valid TM8 SN")
            End If
        End If

        If TM8_SN_valid Then
            Try
                Ping_Success = My.Computer.Network.Ping(TM8_SN.Text)
            Catch ex As Exception
                Ping_Success = False
            End Try
            If Not Ping_Success Then
                TM8_SN.BackColor = Color.Pink
                MsgBox("Can't ping TM8 using " + TM8_SN.Text)
            Else
                TM8_SN.BackColor = Color.LightGreen
            End If
        End If
    End Sub

    Public Sub DebugLog(ByVal Line As String)
        If Not Line.Contains(System.Environment.NewLine) Then
            Line += System.Environment.NewLine
        End If
        LogFileWriter.Write(Line)
    End Sub

    ' Public Sub AppendText(ByVal Line As String, Optional ByVal LogToFile As Boolean = True, Optional ByVal CR As Boolean = True, Optional ByVal UUT_Result As RichTextBox = Nothing)
    Public Sub AppendText(ByVal Line As String, Optional ByVal LogToFile As Boolean = True, Optional ByVal CR As Boolean = True, Optional ByVal UUT As Object = Nothing, Optional ByVal LogAllUuts As Boolean = False, Optional ByVal LogToResults As Boolean = True)
        'TestDetails += Line
        'If Not ReportFileOpen Then
        '    LogToFile = False
        'End If
        'If LogToFile Then
        '    UUT("REPORT").Write(Line)
        'End If
        If CR Then
            If Not Line.EndsWith(System.Environment.NewLine) Then
                'If Not Line.Contains(System.Environment.NewLine) Then
                Line += System.Environment.NewLine
            End If
        End If
        If LogToResults Then Results.AppendText(Line)
        'If Not IsNothing(UUT_Result) Then
        If Not IsNothing(UUT) Then
            UUT("RESULT").AppendText(Line)
            If LogToFile Then
                UUT("REPORT").Write(Line)
                UUT("REPORT").Flush()
            End If
            'UUT_Result.AppendText(Line)
        End If
        If LogAllUuts Then
            For Each UUT In UUTs
                If Not UUT("SN").Text = "" Then
                    UUT("REPORT").Write(Line)
                End If
            Next
        End If
        If LogToResults Then
            If LogFileWriter.BaseStream Is Nothing Then
                Exit Sub
            End If
            LogFileWriter.Write(Line)
        End If

        'If CR Then
        '    If Not Line.Contains(System.Environment.NewLine) Then
        '        'If LogToFile Then
        '        '    ReportFileWriter.Write(System.Environment.NewLine)
        '        'End If
        '        Results.AppendText(System.Environment.NewLine)
        '        'TestDetails += System.Environment.NewLine
        '    End If
        'End If
    End Sub

    Sub Pass(ByVal TestTime As Integer, ByVal AllPassed As Boolean)
        Dim database As New DB
        Dim DB_success As Boolean = False

        For Each UUT In UUTs
            If UUT("SN").Text = "" Then Continue For
            Try
                If Not UUT("FAILED") Then
                    If Stopped Then
                        UUT("TAB").ImageIndex = StatusColor.FAILED
                    Else
                        UUT("TAB").ImageIndex = StatusColor.PASSED
                        If LogToDatabase Then
                            DB_success = database.Pass(UUT("SN").Text, "FINAL_TEST", TestTime)
                        End If
                    End If
                End If
            Catch ex As Exception
                AppendText("Form1.Pass(): Caught " + ex.ToString)
            End Try
        Next
        If Stopped Then
            TestStatus.BackColor = Color.Red
            TestStatus.Text = "FAILED"
        ElseIf ListBox_TestMode.Text = "Debug" Then
            TestStatus.Text = "DEBUG PASSED"
            TestStatus.BackColor = Color.Goldenrod
            TestStatus.ForeColor = Color.Black
        Else
            If AllPassed Then
                TestStatus.Text = "PASSED"
                TestStatus.BackColor = Color.Green
                TestStatus.ForeColor = Color.White
            Else
                TestStatus.Text = "SOME_PASSED"
                TestStatus.BackColor = Color.Orange
                TestStatus.ForeColor = Color.Black
            End If
        End If

        Try
            LogFileWriter.Close()
            LogFile.Close()
        Catch ex As Exception
            AppendText("Form1.Pass(): Caught " + ex.ToString)
        End Try

        For i = 0 To ReportFile.Length() - 1
            If ReportFileWriter(i) Is Nothing Then Continue For
            Try
                ReportFileWriter(i).Close()
                ReportFile(i).Close()
            Catch ex As Exception
                AppendText("Form1.Pass(): Caught " + ex.ToString)
            End Try
        Next
        'If LoggingData Then
        '    ReportFileWriter.Close()
        '    ReportFile.Close()
        '    LoggingData = False
        'End If

        Try
            EnableEntries()
            Stopped = True
            CopyUutLogs()
            'CloseAllUUT_SerialPorts()
            ETA_Timer.Stop()
            EstimatedCompletionDate.Text = Now.ToString
            EstimatedCompletionHours.Text = 0
        Catch ex As Exception
            AppendText("Form1.Pass(): Caught " + ex.ToString)
        End Try
    End Sub

    Sub DisableEntries()
        TM8_SN.Enabled = False
        ListBox_TestMode.Enabled = False
        GDA_ComportSelect.Enabled = False
        StartTest.Enabled = False
        StopButton.Enabled = True
        For i = 0 To Test_Sequence.Count - 1
            Test_Sequence(i).Button.Enabled = False
        Next
        TestButtonPanel.Enabled = False
        U1_Time_Select.Enabled = False
        U2_Time_Select.Enabled = False
    End Sub

    Sub EnableEntries()
        TM8_SN.Enabled = True
        ListBox_TestMode.Enabled = True
        GDA_ComportSelect.Enabled = True
        StartTest.Enabled = True
        If ListBox_TestMode.Text <> "Production" Then
            For i = 0 To Test_Sequence.Count - 1
                Test_Sequence(i).Button.Enabled = True
            Next
            TestButtonPanel.Enabled = True
        End If
        U1_Time_Select.Enabled = True
        U2_Time_Select.Enabled = True
    End Sub

    Sub Fail(Optional ByVal TestName As String = "", Optional ByVal TestTime As Integer = 0)
        Dim database As New DB
        Dim DB_success As Boolean = False

        For Each UUT In UUTs
            If UUT("SN").Text = "" Then Continue For

            Try
                If UUT("FAILED") Then
                    UUT("TAB").ImageIndex = StatusColor.FAILED
                    If LogToDatabase Then
                        DB_success = database.Fail(UUT("SN").Text, "FINAL_TEST", TestTime, TestName, "")
                    End If
                End If
            Catch ex As Exception
                AppendText("Form1.Fail(): Caught " + ex.ToString)
            End Try
        Next

        If Stopped Then
            TestStatus.Text = "STOPPED"
            TestStatus.BackColor = Color.Blue
            TestStatus.ForeColor = Color.White
        Else
            TestStatus.Text = "FAILED"
            TestStatus.BackColor = Color.Red
            TestStatus.ForeColor = Color.Black
        End If

        Try
            LogFileWriter.Close()
            LogFile.Close()
        Catch ex As Exception
            AppendText("Form1.Fail(): Caught " + ex.ToString)
        End Try

        For i = 0 To ReportFile.Length() - 1
            If ReportFileWriter(i) Is Nothing Then Continue For
            Try
                ReportFileWriter(i).Close()
                ReportFile(i).Close()
            Catch ex As Exception
                AppendText("Form1.Fail(): Caught " + ex.ToString)
            End Try
        Next
        'If LoggingData Then
        '    ReportFileWriter.Close()
        '    ReportFile.Close()
        '    LoggingData = False
        'End If

        Try
            EnableEntries()
            Stopped = True
            CopyUutLogs()
            CloseAllUUT_SerialPorts()
            ETA_Timer.Stop()
            EstimatedCompletionDate.Text = Now.ToString
            EstimatedCompletionHours.Text = 0
        Catch ex As Exception
            AppendText("Form1.Fail(): Caught " + ex.ToString)
        End Try
    End Sub

    Private Sub SN_entry_KeyDown(sender As System.Object, e As System.Windows.Forms.KeyEventArgs) Handles SN_entry_left_1.KeyDown,
        SN_entry_left_2.KeyDown, SN_entry_left_3.KeyDown, SN_entry_left_4.KeyDown, SN_entry_left_5.KeyDown, SN_entry_left_6.KeyDown,
        SN_entry_left_7.KeyDown, SN_entry_left_8.KeyDown, SN_entry_left_9.KeyDown, SN_entry_right_1.KeyDown, SN_entry_right_2.KeyDown,
        SN_entry_right_3.KeyDown, SN_entry_right_4.KeyDown, SN_entry_right_5.KeyDown, SN_entry_right_6.KeyDown, SN_entry_right_7.KeyDown,
        SN_entry_right_8.KeyDown, SN_entry_right_9.KeyDown

        Dim TabPage As Object
        Dim UUT_Result As Object
        Dim Msg As String
        Dim leftTabControl As Boolean = True
        If e.KeyCode = Keys.Enter Then
            If sender.Equals(SN_entry_left_1) Then
                TabPage = TabPage1
                UUT_Result = UUT_Result_1
                'SN_entry_left_1.BackColor = Color.LightGreen
                'TabPage1.Text = SN_entry_left_1.Text
            End If
            If sender.Equals(SN_entry_left_2) Then
                TabPage = TabPage2
                UUT_Result = UUT_Result_2
                'SN_entry_left_2.BackColor = Color.LightGreen
                'TabPage2.Text = SN_entry_left_2.Text
            End If
            If sender.Equals(SN_entry_left_3) Then
                TabPage = TabPage3
                UUT_Result = UUT_Result_3
                'SN_entry_left_3.BackColor = Color.LightGreen
                'TabPage3.Text = SN_entry_left_3.Text
            End If
            If sender.Equals(SN_entry_left_4) Then
                TabPage = TabPage4
                UUT_Result = UUT_Result_4
                'SN_entry_left_4.BackColor = Color.LightGreen
                'TabPage4.Text = SN_entry_left_4.Text
            End If
            If sender.Equals(SN_entry_left_5) Then
                TabPage = TabPage5
                UUT_Result = UUT_Result_5
                'SN_entry_left_5.BackColor = Color.LightGreen
                'TabPage5.Text = SN_entry_left_5.Text
            End If
            If sender.Equals(SN_entry_left_6) Then
                TabPage = TabPage6
                UUT_Result = UUT_Result_6
                'SN_entry_left_6.BackColor = Color.LightGreen
                'TabPage6.Text = SN_entry_left_6.Text
            End If
            If sender.Equals(SN_entry_left_7) Then
                TabPage = TabPage7
                UUT_Result = UUT_Result_7
                'SN_entry_left_7.BackColor = Color.LightGreen
                'TabPage7.Text = SN_entry_left_7.Text
            End If
            If sender.Equals(SN_entry_left_8) Then
                TabPage = TabPage8
                UUT_Result = UUT_Result_8
                'SN_entry_left_8.BackColor = Color.LightGreen
                'TabPage8.Text = SN_entry_left_8.Text
            End If
            If sender.Equals(SN_entry_left_9) Then
                TabPage = TabPage9
                UUT_Result = UUT_Result_9
                'SN_entry_left_9.BackColor = Color.LightGreen
                'TabPage9.Text = SN_entry_left_9.Text
            End If

            If sender.Equals(SN_entry_right_1) Then
                TabPage = TabPage11
                UUT_Result = UUT_Result_11
                leftTabControl = False
                'SN_entry_right_1.BackColor = Color.LightGreen
                'TabPage11.Text = SN_entry_right_1.Text
            End If
            If sender.Equals(SN_entry_right_2) Then
                TabPage = TabPage12
                UUT_Result = UUT_Result_12
                leftTabControl = False
                'SN_entry_right_2.BackColor = Color.LightGreen
                'TabPage12.Text = SN_entry_right_2.Text
            End If
            If sender.Equals(SN_entry_right_3) Then
                TabPage = TabPage13
                UUT_Result = UUT_Result_13
                leftTabControl = False
                'SN_entry_right_3.BackColor = Color.LightGreen
                'TabPage13.Text = SN_entry_right_3.Text
            End If
            If sender.Equals(SN_entry_right_4) Then
                TabPage = TabPage14
                UUT_Result = UUT_Result_14
                leftTabControl = False
                'SN_entry_right_4.BackColor = Color.LightGreen
                'TabPage14.Text = SN_entry_right_4.Text
            End If
            If sender.Equals(SN_entry_right_5) Then
                TabPage = TabPage15
                UUT_Result = UUT_Result_15
                leftTabControl = False
                'SN_entry_right_5.BackColor = Color.LightGreen
                'TabPage15.Text = SN_entry_right_5.Text
            End If
            If sender.Equals(SN_entry_right_6) Then
                TabPage = TabPage16
                UUT_Result = UUT_Result_16
                leftTabControl = False
                'SN_entry_right_6.BackColor = Color.LightGreen
                'TabPage16.Text = SN_entry_right_6.Text
            End If
            If sender.Equals(SN_entry_right_7) Then
                TabPage = TabPage17
                UUT_Result = UUT_Result_17
                leftTabControl = False
                'SN_entry_right_7.BackColor = Color.LightGreen
                'TabPage17.Text = SN_entry_right_7.Text
            End If
            If sender.Equals(SN_entry_right_8) Then
                TabPage = TabPage18
                UUT_Result = UUT_Result_18
                leftTabControl = False
                'SN_entry_right_8.BackColor = Color.LightGreen
                'TabPage18.Text = SN_entry_right_8.Text
            End If
            If sender.Equals(SN_entry_right_9) Then
                TabPage = TabPage19
                UUT_Result = UUT_Result_19
                leftTabControl = False
                'SN_entry_right_9.BackColor = Color.LightGreen
                'TabPage19.Text = SN_entry_right_9.Text
            End If

            If Not ValidateSN(sender.Text, Msg) Then
                UUT_Result.Text = sender.Text + ":  " + Msg
                sender.Text = ""
                sender.BackColor = Color.Red
            Else
                If CheckForDuplicateSN() Then
                    UUT_Result.Text = sender.Text + ":  DUPLICATE SN"
                    sender.Text = ""
                    sender.BackColor = Color.Red
                Else
                    TabPage.Text = sender.Text
                    sender.BackColor = Color.LightGreen

                    ' Go to the next tabpage
                    If leftTabControl = True Then
                        If TabControl1.SelectedIndex < 8 Then TabControl1.SelectedIndex += 1
                    Else
                        If TabControl2.SelectedIndex < 8 Then TabControl2.SelectedIndex += 1
                    End If
                End If
            End If

            ' For debug only: Go to the next tabpage
            'If leftTabControl = True Then
            '    If TabControl1.SelectedIndex < 8 Then TabControl1.SelectedIndex += 1
            'Else
            '    If TabControl2.SelectedIndex < 8 Then TabControl2.SelectedIndex += 1
            'End If

        End If
    End Sub

    Private Sub CloseAllUUT_SerialPorts()
        Dim UUT As Hashtable
        Dim SF As New SerialFunctions

        For Each UUT In UUTs
            If UUT("SN").Text = "" Then Continue For
            Try
                SF.Close(UUT)
            Catch ex As Exception
                AppendText(UUT("SN").Text + ": CloseAllUUT_SerialPorts() Caught " + ex.Message, UUT:=UUT)
            End Try
        Next

        If GDA_SP.IsOpen Then
            Try
                GDA_SP.Close()
            Catch ex As Exception
            End Try
        End If
    End Sub

    Private Sub EnableTest(sender As System.Object, e As System.Windows.Forms.MouseEventArgs) Handles Test_1.MouseUp,
        Test_2.MouseUp, Test_3.MouseUp, Test_4.MouseUp, Test_5.MouseUp, Test_6.MouseUp, Test_7.MouseUp,
        Test_8.MouseUp, Test_9.MouseUp, Test_10.MouseUp, Test_11.MouseUp, Test_12.MouseUp, Test_13.MouseUp
        Dim AnyTestEnabled As Boolean = False

        If Not ListBox_TestMode.Text = "Debug" Then
            Exit Sub
        End If

        If e.Button = MouseButtons.Right Then
            For i = 0 To Test_Sequence.Count - 1
                If sender.Equals(Test_Sequence(i).Button) Then
                    If Test_Sequence(i).Enabled Then
                        sender.BackColor = DefaultButtonBackColor
                        sender.ForeColor = DefaultButtonForeColor
                        sender.Enabled = False
                        Test_Sequence(i).Enabled = False
                    Else
                        sender.BackColor = Color.PaleTurquoise
                        sender.ForeColor = Color.Black
                        ' Button.Enabled = True
                        Test_Sequence(i).Enabled = True
                    End If
                    'sender.BackColor = DefaultBackColor
                    '' sender.Enabled = False
                    'Test_Sequence(i).Enabled = False
                End If
                If Test_Sequence(i).Enabled Then
                    AnyTestEnabled = True
                End If
            Next
            If AnyTestEnabled Then
                StartTest.Enabled = True
            Else
                StartTest.Enabled = False
            End If
        End If
    End Sub

    Private Sub DisableTest(sender As System.Object, e As System.Windows.Forms.MouseEventArgs) Handles TestButtonPanel.Click
        Dim button As Button

        If Not ListBox_TestMode.Text = "Debug" Then
            Exit Sub
        End If
        For i = 0 To Test_Sequence.Count - 1
            button = Test_Sequence(i).Button
            If (e.X > button.Location.X) And (e.X < button.Location.X + button.Width) And
               (e.Y > button.Location.Y) And (e.Y < button.Location.Y + button.Height) Then
                'button.BackColor = Color.LightGray
                'button.ForeColor = Color.Black
                button.BackColor = Color.PaleTurquoise
                button.ForeColor = Color.Black
                button.Enabled = True
                Test_Sequence(i).Enabled = True
                StartTest.Enabled = True
                Exit For
            End If
        Next
    End Sub

    Private Sub Form1_FormClosing(sender As System.Object, e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        For Each UUT In UUTs
            If Not UUT("SP") Is Nothing Then
                If UUT("SP").IsOpen Then
                    Try
                        UUT("SP").Close()
                    Catch ex As Exception

                    End Try
                End If
            End If
        Next
        If GDA_SP.IsOpen Then
            Try
                GDA_SP.Close()
            Catch ex As Exception
            End Try
        End If
    End Sub

    Private Sub GDA_ComportSelect_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles GDA_ComportSelect.SelectedIndexChanged
        Dim ErrorMsg As String

        GDA_SP = New SerialPort(GDA_ComportSelect.Text, 2400, 0, 8, 1)

        Try
            GDA_SP.Open()
            GDA_SP.ReadTimeout = 25
        Catch ex As Exception
            ErrorMsg = "Problem opening GDA serial port" + vbCr
            ErrorMsg += ex.ToString
            MsgBox(ErrorMsg)
        End Try

    End Sub

    Private Sub InitializeStatusIcons()
        For i = 0 To 3
            statusIconImages(i) = New Bitmap(10, 10)
        Next

        For x = 0 To statusIconImages(0).Width - 1
            For y = 0 To statusIconImages(0).Height - 1
                statusIconImages(0).SetPixel(x, y, Color.Gray)
                statusIconImages(1).SetPixel(x, y, Color.Yellow)
                statusIconImages(2).SetPixel(x, y, Color.Green)
                statusIconImages(3).SetPixel(x, y, Color.Red)
            Next
        Next

        statusIcons = New ImageList
        statusIcons.TransparentColor = Color.Gray
        For i = 0 To 3
            statusIcons.Images.Add(statusIconImages(i))
        Next
        TabControl1.ImageList = statusIcons
        TabControl2.ImageList = statusIcons

        For Each UUT In UUTs
            UUT("TAB").ImageIndex = 0
        Next
    End Sub

    Sub CopyUutLogs()
        Dim source_file As String
        Dim dest_file As String
        Dim FilePrefixesToCopy() As String = {"REC", "EVENT", "CONFIG", "H2SCAN_DATA", "U1", "U2",
                                              "VER_1000", "VER_6000", "VER_10000", "TestReport",
                                              "Test Summary Report"}
        Dim FilePrefix As String
        Dim di As DirectoryInfo
        Dim files() As FileSystemInfo
        Dim comparer As IComparer = New DateComparer()

        For Each UUT In UUTs
            If UUT("SN").Text = "" Or UUT("log_filename") = "" Then Continue For
            Try
                source_file = ReportDir + "FINAL_TEST" + "\" + UUT("SN").Text + "\" + UUT("log_filename")
                dest_file = FinalReportDir + "FINAL_TEST" + "\" + UUT("SN").Text + "\" + UUT("log_filename")
                File.Copy(source_file, dest_file, True)

                di = New DirectoryInfo(ReportDir + "FINAL_TEST" + "\" + UUT("SN").Text)
                For Each FilePrefix In FilePrefixesToCopy
                    If FilePrefix = "TestReport" Then FilePrefix = UUT("SN").Text + "_" + FilePrefix
                    If FilePrefix = "Test Summary Report" Then FilePrefix = UUT("SN").Text + " " + FilePrefix
                    files = di.GetFileSystemInfos(FilePrefix + "*")
                    Array.Sort(files, comparer)
                    If files.Length > 0 Then
                        source_file = files(0).FullName
                        dest_file = FinalReportDir + "FINAL_TEST" + "\" + UUT("SN").Text + "\" + files(0).Name
                        If FilePrefix.Contains("Test Summary Report") Then
                            dest_file = dest_file.Insert(dest_file.IndexOf(".doc"), "." + Format(Date.UtcNow, "yyyyMMddHHmmss"))
                        End If
                        If Not File.Exists(dest_file) Then
                            File.Copy(source_file, dest_file)
                        End If
                    End If
                Next
            Catch ex As Exception
                MsgBox("CopyUutLogs() caught " + ex.Message)
            End Try
        Next
    End Sub

    Private Class DateComparer
        Implements System.Collections.IComparer

        Public Function Compare(ByVal info1 As Object, ByVal info2 As Object) As Integer Implements System.Collections.IComparer.Compare
            Dim FileInfo1 As System.IO.FileInfo = DirectCast(info1, System.IO.FileInfo)
            Dim FileInfo2 As System.IO.FileInfo = DirectCast(info2, System.IO.FileInfo)

            Dim Date1 As DateTime = FileInfo1.CreationTime
            Dim Date2 As DateTime = FileInfo2.CreationTime

            If Date1 > Date2 Then Return -1
            If Date1 < Date2 Then Return 1

            Return 0
        End Function
    End Class

    Private Sub Stop_UUT_1_Click(sender As System.Object, e As System.EventArgs) Handles Stop_UUT_1.Click,
        Stop_UUT_2.Click, Stop_UUT_3.Click, Stop_UUT_4.Click, Stop_UUT_5.Click, Stop_UUT_6.Click, Stop_UUT_7.Click, Stop_UUT_8.Click, Stop_UUT_9.Click,
        Stop_UUT_11.Click, Stop_UUT_12.Click, Stop_UUT_13.Click, Stop_UUT_14.Click, Stop_UUT_15.Click, Stop_UUT_16.Click, Stop_UUT_17.Click, Stop_UUT_18.Click,
        Stop_UUT_19.Click

        Dim b As Button
        Dim uut_index As Integer
        Dim r As RichTextBox
        Dim r_index As Integer
        Dim DR As DialogResult

        b = sender
        sender.enabled = False
        uut_index = Split(b.Name, "_")(2)

        For Each UUT In UUTs
            r = UUT("RESULT")
            r_index = Split(r.Name, "_")(2)
            If r_index = uut_index Then
                DR = MessageBox.Show("Do you wish to stop testing on this UUT?", "STOP TEST?", MessageBoxButtons.YesNo)
                If DR = DialogResult.Yes Then
                    UUT("FAILED") = True
                    UUT("TAB").ImageIndex = StatusColor.FAILED
                End If
                Application.DoEvents()
                Exit For
            End If
        Next
    End Sub

    Private Sub StopButton_Click(sender As System.Object, e As System.EventArgs) Handles StopButton.Click
        Dim DR As DialogResult

        DR = MessageBox.Show("Do you wish to stop the test?", "CONFIRM STOP TEST?", MessageBoxButtons.YesNo)
        If Not DR = DialogResult.Yes Then
            Exit Sub
        End If

        StopButton.Enabled = False
        Stopped = True
        TestStatus.Text = "STOPPING"
        TestStatus.BackColor = Color.OrangeRed
        TestStatus.ForeColor = Color.Black
        If LoggingData Then
            LoggingData = False
        End If
    End Sub

    Private Sub ListBox_TestMode_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ListBox_TestMode.SelectedIndexChanged
        Dim EnableTestButtons As Boolean

        If ListBox_TestMode.SelectedIndex = 0 Then ' PRODUCTION
            EnableTestButtons = False
            TestButtonPanel.Enabled = False
        Else
            EnableTestButtons = True
            TestButtonPanel.Enabled = True
        End If
        For i = 0 To Test_Sequence.Count - 1
            Test_Sequence(i).Button.Enabled = EnableTestButtons
            Test_Sequence(i).Enabled = True
            Test_Sequence(i).Button.BackColor = Color.PaleTurquoise
            Test_Sequence(i).Button.ForeColor = Color.Black
        Next

        ' In case of a partial run, disable the selection of the INITIAL button
        If ListBox_TestMode.Text = "Partial Run" Then
            Test_Sequence(0).Button.Enabled = False
        Else
            PartialRun = False
        End If

    End Sub

    Private Function CheckForDuplicateSN() As Boolean
        Dim SN_i As String
        Dim SN_j As String

        For i = 1 To 18
            Select Case i
                Case 1
                    SN_i = SN_entry_left_1.Text
                Case 2
                    SN_i = SN_entry_left_2.Text
                Case 3
                    SN_i = SN_entry_left_3.Text
                Case 4
                    SN_i = SN_entry_left_4.Text
                Case 5
                    SN_i = SN_entry_left_5.Text
                Case 6
                    SN_i = SN_entry_left_6.Text
                Case 7
                    SN_i = SN_entry_left_7.Text
                Case 8
                    SN_i = SN_entry_left_8.Text
                Case 9
                    SN_i = SN_entry_left_9.Text
                Case 10
                    SN_i = SN_entry_right_1.Text
                Case 11
                    SN_i = SN_entry_right_2.Text
                Case 12
                    SN_i = SN_entry_right_3.Text
                Case 13
                    SN_i = SN_entry_right_4.Text
                Case 14
                    SN_i = SN_entry_right_5.Text
                Case 15
                    SN_i = SN_entry_right_6.Text
                Case 16
                    SN_i = SN_entry_right_7.Text
                Case 17
                    SN_i = SN_entry_right_8.Text
                Case 18
                    SN_i = SN_entry_right_9.Text
            End Select
            If SN_i = "" Then Continue For
            For j = i + 1 To 18
                Select Case j
                    Case 2
                        SN_j = SN_entry_left_2.Text
                    Case 3
                        SN_j = SN_entry_left_3.Text
                    Case 4
                        SN_j = SN_entry_left_4.Text
                    Case 5
                        SN_j = SN_entry_left_5.Text
                    Case 6
                        SN_j = SN_entry_left_6.Text
                    Case 7
                        SN_j = SN_entry_left_7.Text
                    Case 8
                        SN_j = SN_entry_left_8.Text
                    Case 9
                        SN_j = SN_entry_left_9.Text
                    Case 10
                        SN_j = SN_entry_right_1.Text
                    Case 11
                        SN_j = SN_entry_right_2.Text
                    Case 12
                        SN_j = SN_entry_right_3.Text
                    Case 13
                        SN_j = SN_entry_right_4.Text
                    Case 14
                        SN_j = SN_entry_right_5.Text
                    Case 15
                        SN_j = SN_entry_right_6.Text
                    Case 16
                        SN_j = SN_entry_right_7.Text
                    Case 17
                        SN_j = SN_entry_right_8.Text
                    Case 18
                        SN_j = SN_entry_right_9.Text
                End Select
                If SN_j = "" Then Continue For
                If SN_i = SN_j Then
                    Return True
                End If
            Next
        Next
        Return False
    End Function

    Function ValidateSN(ByVal SN As String, ByRef Msg As String) As Boolean
        Dim YY As Integer
        Dim WW As Integer
        Dim NNNN As Integer
        Dim Fields() As String

        If Not Regex.IsMatch(SN, "TM10\d01\d\d\d\d\d\d\d\d") Then
            Msg = "TM1 Serial Number does not match format 'TM10X01YYWWNNN'"
            Return False
        End If
        Fields = Regex.Split(SN, "TM10\d01(\d\d)(\d\d)(\d\d\d\d)")
        YY = Fields(1)
        WW = Fields(2)
        NNNN = Fields(3)
        If YY < 12 Then
            Msg = "YY field of SN using format 'TM10X01YYWWNNN' expected to be > 12"
            Return False
        End If
        If WW > 52 Then
            Msg = "WW field of SN using format 'TM10X01YYWWNNN' expected between 1 & 52"
            Return False
        End If
        Return True
    End Function

    Private Sub ETA_Timer_Tick(sender As System.Object, e As System.EventArgs) Handles ETA_Timer.Tick
        Dim TimeToCompletion As TimeSpan

        EstimatedTestCompletionHours -= 0.1
        TimeToCompletion = New TimeSpan(0, EstimatedTestCompletionHours * 60, 0)
        EstimatedTestCompletionDate = Now.Add(TimeToCompletion)
        EstimatedCompletionHours.Text = EstimatedTestCompletionHours.ToString + " hours"
        EstimatedCompletionDate.Text = EstimatedTestCompletionDate.ToString
    End Sub

    'Sub UutDataAvailable_Handler(UUT As Hashtable)
    '    MsgBox(UUT("CLASS_REF").MessagesDequeue)
    'End Sub

    Private Sub PrintTestReportsToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles PrintTestReportsToolStripMenuItem.Click
        Dim PD As New PrintDialog
        Dim app As Word.Application
        Dim filename As String
        Dim m As Object = System.Reflection.Missing.Value
        Dim doc As Word.Document

        If PD.ShowDialog() = Windows.Forms.DialogResult.Cancel Then
            Exit Sub
        End If

        For Each UUT In UUTs
            If UUT("SN").Text = "" Or UUT("FAILED") Then Continue For
            'filename = FinalReportDir + "FINAL_TEST\" + UUT("SN").Text + "\" + UUT("SN").Text + " Test Summary Report.doc"
            filename = "C:\Operations\Production\TM1\FINAL_TEST\" + UUT("SN").Text + "\" + UUT("SN").Text + " Test Summary Report.doc"
            If Not File.Exists(filename) Then
                MsgBox(UUT("SN").Text + ":  missing test report " + filename)
                Continue For
            End If
            app = New Word.Application
            Try
                app.WordBasic.FilePrintSetup(Printer:=PD.PrinterSettings.PrinterName, DoNotSetAsSysDefault:=1)
                doc = app.Documents.Open(filename, m, m, m, m, m, m, m, m, m, m, m)
                app.PrintOut()
                app.Documents.Close()
                app.Quit()
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        Next
    End Sub

    Private Sub U1_Time_Select_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles U1_Time_Select.SelectedIndexChanged
        Timeout_U1 = CDbl(U1_Time_Select.Text)
    End Sub

    Private Sub U2_Time_Select_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles U2_Time_Select.SelectedIndexChanged
        Timeout_U2 = CDbl(U2_Time_Select.Text)
    End Sub

    Private Sub TM8_SN_TextChanged(sender As System.Object, e As System.EventArgs) Handles TM8_SN.TextChanged

    End Sub

    ' This is for creating reports from cvs files
    Private Sub StartTest_Click_CreateTestReports(sender As System.Object, e As System.EventArgs)
        'Handles StartTest.Click, Test_1.Click,
        'Test_2.Click, Test_3.Click, Test_4.Click, Test_5.Click, Test_6.Click, Test_7.Click, Test_8.Click, Test_9.Click,
        'Test_10.Click, Test_11.Click, Test_12.Click, Test_13.Click

        Dim DebugMode As Boolean = False
        Dim RunSingleTest As Boolean = False
        Dim Test_Index As Integer
        Dim Retry As Boolean
        Dim RetryCnt As Integer
        Dim Test_Status As Boolean
        Dim TestName As String
        Dim TestTime As Integer = 0
        Dim TestStartTime As DateTime = Now
        Dim ReportFilePath As String
        Dim FinalFilePath As String
        Dim TimeStamp As String
        Dim LogFilePath As String
        Dim AllPassed As Boolean
        Dim ReportFileIndex As Integer
        Dim AllFailed As Boolean
        Dim TimeToCompletion As TimeSpan

        '' Inject the unit info needed for creating reports from cvs files
        ''InfoInit()
        'Dim sn As String() = {
        '    "TM1020115020081",
        '    "TM1020114500069",
        '    "TM1020114500068",
        '    "TM1020114500075",
        '    "TM1020115020091",
        '    "TM1020115020092",
        '    "TM1020115020093",
        '    "TM1020115020094",
        '    "TM1020115020095",
        '    "TM1020115020096",
        '    "TM1020115020097",
        '    "TM1020115020099",
        '    "TM1020115020100",
        '    "TM1020115040103"}
        'Dim i As Integer = 0
        'For Each s As String In sn
        '    UUTs(i)("SN").Text = sn(i)
        '    If i < 2 Then
        '        UUTs(i)("FAILED") = False
        '    Else
        '        UUTs(i)("FAILED") = True
        '    End If
        '    i += 1
        'Next

        Log.WriteLine("Test Started")

        ' This reset fail flag to enable test rerun
        If hasStarted Then ResetFailFlags()
        hasStarted = True

        Log.WriteLine("Creating Calibration log file")
        TimeStamp = Format(Date.UtcNow, "yyyyMMddHHmmss")
        LogFilePath = "C:\Temp\TM1_Calibration_Verification_log_file" + TimeStamp + ".txt"
        Try
            LogFile = New FileStream(LogFilePath, FileMode.Create, FileAccess.Write)
            LogFileWriter = New StreamWriter(LogFile)
        Catch ex As Exception
            MsgBox("Problem opening log file " + LogFilePath)
            Exit Sub
        End Try

        Stopped = False
        If ListBox_TestMode.Text = "Debug" Then
            DebugMode = True
        End If

        If Not sender.Equals(StartTest) Then
            RunSingleTest = True
            If Not DebugMode Then
                Stopped = True
                Exit Sub
            End If
        End If

        Log.WriteLine("Clearing Results")
        Results.Clear()
        TestStatus.Text = "RUNNING"
        TestStatus.BackColor = Color.Yellow
        TestStatus.ForeColor = Color.Black

        If RunSingleTest Then
            For i = 0 To Test_Sequence.Count - 1
                If sender.Equals(Test_Sequence(i).Button) Then
                    Test_Index = i
                End If
            Next
            If Test_Index < 0 Then
                AppendText("Could not find test for " + sender.Text)
                'For i = 0 To Test_Sequence.Count - 1
                '    Test_Sequence(i).Button.Enabled = True
                'Next
                Fail()
                Exit Sub
            End If
        End If

        For i = 0 To Test_Sequence.Count - 1
            If RunSingleTest And Not i = Test_Index Then
                Continue For
            End If
            'If Not Test_Sequence(i).Button.Enabled Then
            If Not Test_Sequence(i).Enabled Then
                Continue For
            End If
            'If Not CheckEntriesForTest(Test_Sequence(i)) Then
            '    Fail()
            '    Exit Sub
            'End If
        Next
        DisableEntries()

        For i = 0 To Test_Sequence.Count - 1
            If Test_Sequence(i).Enabled Then
                Test_Sequence(i).Button.BackColor = Color.PaleTurquoise
                Test_Sequence(i).Button.ForeColor = Color.Black
            End If
            'If ListBox_TestMode.Text = "Debug" Then
            '    If Test_Sequence(i).Button.Enabled Then
            '        Test_Sequence(i).Button.BackColor = Color.LightGray
            '        Test_Sequence(i).Button.ForeColor = Color.Black
            '    End If
            'Else
            '    Test_Sequence(i).Button.Enabled = True
            '    Test_Sequence(i).Button.ForeColor = DefaultForeColor
            '    Test_Sequence(i).Button.BackColor = DefaultBackColor
            'End If
        Next

        TestStatus.Text = "RUNNING"
        TestStatus.BackColor = Color.Yellow
        TestStatus.ForeColor = Color.Black

        ReportFileIndex = 0
        For Each UUT In UUTs
            If Not UUT("SN").Text = "" Then
                Log.WriteLine("Creating report for " + UUT("SN").Text)
                UUT("TAB").ImageIndex = StatusColor.RUNNING
                ReportFilePath = ReportDir + "FINAL_TEST" + "\" + UUT("SN").Text
                Try
                    Directory.CreateDirectory(ReportFilePath)
                Catch ex As Exception
                    AppendText("Problem creating directory " + ReportFilePath)
                    Fail()
                End Try

                FinalFilePath = FinalReportDir + "FINAL_TEST" + "\" + UUT("SN").Text
                Try
                    Directory.CreateDirectory(FinalFilePath)
                Catch ex As Exception
                    AppendText("Problem creating directory " + FinalFilePath)
                    Fail()
                End Try

                Try
                    ReportFile(ReportFileIndex) = New FileStream(ReportFilePath + "\" + TimeStamp + ".txt", FileMode.Create, FileAccess.Write)
                    ReportFileWriter(ReportFileIndex) = New StreamWriter(ReportFile(ReportFileIndex))
                    UUT("REPORT") = ReportFileWriter(ReportFileIndex)
                Catch ex As Exception
                    AppendText("Problem opening log file")
                    AppendText(ex.ToString)
                    Fail()
                End Try
                UUT("log_filename") = TimeStamp + ".txt"
                ReportFileIndex += 1
            Else
                UUT("TAB").ImageIndex = StatusColor.NONE
            End If
        Next
        Application.DoEvents()

        If Stopped Then
            Fail()
            Exit Sub
        End If
        If Not DebugMode Then
            LogToDatabase = True
        End If

        For i = 0 To Test_Sequence.Count - 1
            If RunSingleTest And Not i = Test_Index Then
                Continue For
            End If
            'If Not Test_Sequence(i).Button.Enabled = True Then
            If Not Test_Sequence(i).Enabled = True Then
                Continue For
            End If

            ETA_Timer.Stop()
            EstimatedTestCompletionHours = 0
            If RunSingleTest Then
                EstimatedTestCompletionHours = Test_Sequence(i).Timeout
            End If
            For j = i To Test_Sequence.Count - 1
                If Test_Sequence(j).Enabled Then
                    EstimatedTestCompletionHours += Test_Sequence(j).Timeout
                End If
            Next
            TimeToCompletion = New TimeSpan(0, EstimatedTestCompletionHours * 60, 0)
            EstimatedTestCompletionDate = Now.Add(TimeToCompletion)
            EstimatedCompletionHours.Text = Math.Round(EstimatedTestCompletionHours, 1).ToString + " hours"
            EstimatedCompletionDate.Text = EstimatedTestCompletionDate.ToString
            ETA_Timer.Start()

            Test_Sequence(i).Button.BackColor = Color.Yellow
            Test_Sequence(i).Button.ForeColor = Color.Black
            RetryCnt = 0
            Retry = True
            While Retry
                AppendText("##########################################################", LogAllUuts:=True)
                If (RetryCnt = 0) Then
                    AppendText("TEST_START:  " + Test_Sequence(i).Button.Text, LogAllUuts:=True)
                Else
                    AppendText("TEST_RETRY:  " + Test_Sequence(i).Button.Text, LogAllUuts:=True)
                End If
                AppendText("START_TIME:  " + Format(Date.UtcNow, "yyyyMMddHHmmss"), LogAllUuts:=True)
                RetryCnt += 1
                TestName = Test_Sequence(i).Button.Text
                Log.WriteLine("Doing sequence " + TestName)
                Test_Status = Test_Sequence(i).Handler.Invoke()
                Log.WriteLine("Sequence complete")
                AllFailed = True
                For Each UUT In UUTs
                    If UUT("SN").Text = "" Then Continue For
                    If Not UUT("FAILED") Then
                        AllFailed = False
                    End If
                Next
                If AllFailed Then Test_Status = False
                TestTime = Now.Subtract(TestStartTime).TotalSeconds
                AppendText("END_TIME:  " + Format(Date.UtcNow, "yyyyMMddHHmmss"), LogAllUuts:=True)
                Retry = False
                If Not Test_Status Then
                    AppendText("TEST_STATUS:  FAILED", LogAllUuts:=True)
                    If RetriesAllowed And Not Stopped Then
                        If (MessageBox.Show("Test failed, click yes to retry", "RETRY?", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes) Then
                            Retry = True
                        End If
                    End If
                    If Not Retry Then
                        If Stopped Then
                            Test_Sequence(i).Button.BackColor = Color.Blue
                            Test_Sequence(i).Button.ForeColor = Color.White
                        Else
                            Test_Sequence(i).Button.BackColor = Color.Red
                            Test_Sequence(i).Button.ForeColor = Color.White
                        End If
                        Fail(TestName, TestTime)
                        Exit Sub
                    End If
                End If
            End While
            'AppendText("TEST_STATUS:  PASSED")
            Test_Sequence(i).Button.ForeColor = Color.Black
            AllPassed = True
            For Each UUT In UUTs
                If UUT("SN").Text = "" Then Continue For
                If UUT("FAILED") Then
                    AllPassed = False
                End If
            Next
            If AllPassed Then
                AppendText("TEST_STATUS:  PASSED", LogAllUuts:=True)
                Test_Sequence(i).Button.BackColor = Color.LightGreen
            Else
                AppendText("TEST_STATUS:  SOME_PASSED")
                Test_Sequence(i).Button.BackColor = Color.PaleGoldenrod
            End If
        Next

        If AllPassed Then
            AppendText("FINAL_STATUS:  PASSED", LogAllUuts:=True)
        Else
            AppendText("FINAL_STATUS:  SOME_PASSED")
        End If
        Pass(TestTime, AllPassed)
        LogToDatabase = False

    End Sub

    Public Shared Function SkipTest() As Boolean
        Return True
    End Function

    Public Shared PartialRun As Boolean = False
    Private Sub Button1_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Test_1.MouseDown, Test_2.MouseDown,
        Test_3.MouseDown, Test_4.MouseDown, Test_5.MouseDown, Test_6.MouseDown, Test_7.MouseDown, Test_8.MouseDown, Test_9.MouseDown, Test_10.MouseDown,
        Test_11.MouseDown, Test_12.MouseDown, Test_13.MouseDown

        Dim theButton As Button = CType(sender, Button)
        Dim start As Boolean = False
        PartialRun = False

        If ListBox_TestMode.Text <> "Partial Run" Then
            theButton.PerformClick()
            Exit Sub
        End If

        If (MessageBox.Show("Do you want to start from " + theButton.Text + " test?", "Partial Test Start",
                           MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes) Then
            For i As Integer = 0 To Test_Sequence.Count - 1
                If (Test_Sequence(i).Button.Text = theButton.Text) Then
                    start = True
                    PartialRun = True
                End If

                If start Then
                    Test_Sequence(i).Enabled = True
                Else
                    Test_Sequence(i).Enabled = False
                End If
            Next

        End If

        ' Invoke start test button click
        If PartialRun Then Me.StartTest.PerformClick()

        ''Works with both buttons
        'If e.Button = Windows.Forms.MouseButtons.Right Then
        '    MsgBox(theButton.Text + " Right Mouse Button.")
        'Else
        '    MsgBox(theButton.Text + " Left Mouse Button.")
        'End If
    End Sub
End Class
'                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                     