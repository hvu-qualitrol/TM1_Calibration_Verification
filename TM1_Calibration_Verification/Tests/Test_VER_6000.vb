Partial Class Tests
    Public Shared Function Test_VER_6000() As Boolean
        Dim Success As Boolean
        Dim AllTm1sInSpec As Boolean
        Dim T As New TM1

        ' expect TM1 to stabilize within 12 hours to +/-10% +/-15ppm gas in oil
        ' expect last 4 hours of TM1 data to be in spec

        If Not TM8_gas_in_spec Then
            Form1.AppendText("TM8 gas ppm not in spec")
            Return False
        End If

        ' diffPpmSpec = 10%: The difference between the TM8 & TM1 measurements
        ' delPpmSpec = 4%: The delta between the two consecutive of TM1 measurements
        Dim diffPpmSpec As Double = 10.0
        Dim delPpmSpec As Double = 4.0
        Success = T.WaitForTm1PpmInSpec(6000, 7000, 5000, AllTm1sInSpec, Timeout_VER_6000, diffPpmSpec, delPpmSpec)
        'uccess = T.WaitForTm1PpmInSpec(6000, 7000, 5000, AllTm1sInSpec, 22, 10)
        'If AllTm1sInSpec Then
        '    Return Success
        'Else
        '    Return False
        'End If

        Return Success

    End Function
End Class