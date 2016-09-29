Public Enum AssemblyType
    CONTROLLER_BOARD = 0
    H2SCAN = 1
    TM1 = 2
End Enum

Public Enum StatusColor
    NONE = 0
    RUNNING = 1
    PASSED = 2
    FAILED = 3
End Enum


Module Module1
    Public ReportDir As String
    Public FinalReportDir As String
    Public LoggingData As Boolean = False
    Public Stopped As Boolean = True
    Public RetriesAllowed As Boolean = False
    Public SerialPort() = Nothing
    Public UUTs As New ArrayList
    Public Products As New Dictionary(Of String, Hashtable)
    Public Product As String = "STANDARD"
    Public TM8_telnet As Telnet

    'Public Tm1RecFields() As String = {"run#|I", "Timestamp|DT", "H2_OIL.PPM|I", "H2_ROC|S", "Moisture|D",
    '                                   "RS|D", "Temperature|D", "RecordStatus|S", "H2_DGA.PPM|I", "H2.PPM|I",
    '                                   "H2_PEAK_PPM|I", "SnsrTemp|D", "PcbTemp|D", "OilTemp|D", "H2_G.PPM|I",
    '                                   "H2_SldAv|I", "DailyROC|I", "WeeklyROC|I", "MonthlyROC|I", "Status|S",
    '                                   "Mode|I", "SnsrAdc|I", "PCBAdc|I", "HCurrent|I", "ResAdc|I", "AdjRes|I",
    '                                   "H2Res.PPM|I", "H2Leg.PPM|I", "RunSecHi|I", "RunSecLo|I", "OilblockTemp|D",
    '                                   "AnalogbdTemp|D", "AmbientTemp|D", "Tach|D"}

    Public Tm1Rev0RecFields() As String = {"run#|UI", "Timestamp|DT", "H2_OIL.PPM|UI", "H2_ROC|S", "Moisture|D",
                                   "AUX1|D", "AUX2|D", "RecordStatus|S", "H2_DGA.PPM|UI", "H2.PPM|UI",
                                   "H2_PEAK_PPM|UI", "SnsrTemp|D", "PcbTemp|D", "OilTemp|D", "H2_G.PPM|UI",
                                   "H2_SldAv|UI", "DailyROC|UI", "WeeklyROC|UI", "MonthlyROC|UI", "Status|S",
                                   "Mode|S", "SnsrAdc|UI", "PCBAdc|UI", "HCurrent|UI", "ResAdc|UI", "AdjRes|UI",
                                   "H2Res.PPM|UI", "H2Leg.PPM|UI", "RunSecHi|UI", "RunSecLo|UI", "OilblockTemp|D",
                                   "AnalogbdTemp|D", "AmbientTemp|D", "Tach|UI"}
    Public Tm1Rev2RecFields() As String = {"run#|UI", "Timestamp|DT", "H2_OIL.PPM|UI", "H2_ROC|S", "Moisture|D",
                                   "Relative_Saturation|D", "External_OilTemp|D", "RecordStatus|S", "Internal_Relative_Saturation|D", "Internal_OilTemp|D", "H2_DGA.PPM|UI", "H2.PPM|UI",
                                   "H2_PEAK_PPM|UI", "SnsrTemp|D", "PcbTemp|D", "OilTemp|D", "H2_G.PPM|UI",
                                   "H2_SldAv|UI", "DailyROC|UI", "WeeklyROC|UI", "MonthlyROC|UI", "Status|S",
                                   "Mode|S", "SnsrAdc|UI", "PCBAdc|UI", "HCurrent|UI", "ResAdc|UI", "AdjRes|UI",
                                   "H2Res.PPM|UI", "H2Leg.PPM|UI", "RunSecHi|UI", "RunSecLo|UI", "OilblockTemp|D",
                                   "AnalogbdTemp|D", "AmbientTemp|D", "Tach|UI"}

    Public TM8_gas_start_in_spec As DateTime
    Public TM8_gas_in_spec As Boolean = False
    Public ConnectString As String
    Public AssemblyLevel As String = "FINAL"
    Public GDA_SP As System.IO.Ports.SerialPort
    Public statusIcons As ImageList
    Public statusIconImages(4) As Bitmap

    ' Timeout hours
    Public Timeout_GDA As Double = 4.0
    Public Timeout_U1 As Double = 36.0
    Public Timeout_U2 As Double = 20.0
    Public Timeout_VER_10000 As Double = 16.0
    Public Timeout_VER_6000 As Double = 22.0
    Public Timeout_VER_1000 As Double = 36.0
    Public EstimatedTestCompletionHours As Double = 0
    Public EstimatedTestCompletionDate As Date

    ' Test Report
    Public CREATE_TEST_REPORT_DOC As Boolean = True

End Module