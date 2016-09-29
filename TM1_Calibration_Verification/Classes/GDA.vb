Imports System.Text.RegularExpressions

Public Class GDA
    Function SetGDA(ByVal ppm As Integer) As Boolean
        Dim conc As Integer = 1
        Dim B0 As Integer
        Dim B1 As Integer
        Dim B10 As Integer
        Dim GdaSet As Boolean = False
        Dim retryCnt As Integer
        Dim Buffer As String = ""
        Dim EndFound As Boolean = False
        Dim Data As Integer
        Dim GdaSetting As Integer
        Dim GdaSetting_B1 As Integer

        If ppm > 5000 Then conc = 30
        B10 = (ppm * 8191) / (conc * 10000)
        B0 = B10 And &H7F
        B1 = (B10 And &H1F80) / 128
        If conc = 30 Then
            B1 = B1 Or &H40
        End If

        Try
            If Not GDA_SP.IsOpen Then
                GDA_SP.Open()
                GDA_SP.ReadTimeout = 25
            End If

            Form1.AppendText("Attempt to set GDA to " + ppm.ToString + " ppm")
            While (Not GdaSet And retryCnt < 8)
                retryCnt += 1
                GDA_SP.DiscardInBuffer()
                Buffer = ""
                While Not EndFound
                    Try
                        Data = GDA_SP.ReadChar()
                        Buffer += Chr(Data)
                    Catch ex As Exception
                        If Regex.IsMatch(Buffer, "B0=\d+:B1=\d+#$") Then
                            EndFound = True
                            GdaSetting = Regex.Split(Buffer, "(\d+):B1=\d+#$")(1)
                            GdaSetting_B1 = Regex.Split(Buffer, "\d+:B1=(\d+)#$")(1)
                            If GdaSetting = B0 And GdaSetting_B1 = B1 Then
                                GdaSet = True
                            End If
                        Else
                            Buffer = ""
                        End If
                    End Try
                End While
                EndFound = False
                If Not GdaSetting Then
                    GDA_SP.Write(":" + Chr(B0) + Chr(B1))
                End If
            End While

            If GdaSet Then
                Form1.AppendText("GDA SET to B1=" + B1.ToString + ", B0=" + B0.ToString + " in " + retryCnt.ToString + " attempts ")
            End If

        Catch ex As Exception
            Form1.AppendText("GDC.SetGDA() caught exception: " + ex.Message)
            Return False
        End Try

        Return GdaSet
    End Function

    Function WaitForTm8PpmInSpecDebug(ByVal LL As Double, ByVal UL As Double, ByVal timeout As Double) As Boolean
        Dim Tm8RecData As New DataTable
        Dim TM8 As New TM8
        Dim startTime As DateTime
        Dim TM8_gas_ppm As Double
        Dim getRecordErrorCount As Integer = 0

        For attempt As Integer = 0 To 2
            Try
                TM8_gas_in_spec = False
                startTime = Now
                getRecordErrorCount = 0
                'While Not TM8_gas_in_spec And Now.Subtract(startTime).TotalHours < timeout
                While Now.Subtract(startTime).TotalHours < timeout
                    Application.DoEvents()
                    'Get TM8 rec Data
                    If Not TM8.GetRecdata(Tm8RecData, 3) Then
                        Form1.AppendText(TM8.ErrorMsg)
                        If (getRecordErrorCount < 10) Then
                            getRecordErrorCount += 1
                            Continue While
                        Else
                            Return False
                        End If
                    End If
                    If Tm8RecData.Rows.Count > 0 Then
                        TM8_gas_ppm = Tm8RecData.Rows(0)("gas_ppm")
                        Form1.TM8_PPM.Text = TM8_gas_ppm.ToString
                        If TM8_gas_ppm >= LL And TM8_gas_ppm <= UL Then
                            If Not TM8_gas_in_spec Then TM8_gas_start_in_spec = Tm8RecData.Rows(0)("RecTime")
                            TM8_gas_in_spec = True
                        End If
                    End If

                    ' Update the timeout and wait for 15 minutes before continuing
                    Form1.TimeoutLabel.Text = Math.Round(timeout * 60 - Now.Subtract(startTime).TotalMinutes).ToString + "m"
                    CommonLib.Delay(60)
                End While
                If Not TM8_gas_in_spec Then
                    Form1.AppendText("Timeout waiting for TM8 gas-in-gas to be between " + LL.ToString + " & " + UL.ToString + " ppm")
                Else
                    Form1.AppendText("TM8 in spec time = " + TM8_gas_start_in_spec.ToString)
                End If
                Exit For
            Catch ex As Exception
                Form1.AppendText("GDC.WaitForTm8PpmInSpec() attempt " + attempt.ToString + " caught exception: " + ex.Message)
                If (attempt = 2) Then
                    Return False
                End If
            End Try
        Next

        Return TM8_gas_in_spec

    End Function

    Function WaitForTm8PpmInSpec(ByVal LL As Double, ByVal UL As Double, ByVal timeout As Double) As Boolean
        Dim Tm8RecData As New DataTable
        Dim TM8 As New TM8
        Dim startTime As DateTime
        Dim TM8_gas_ppm As Double
        Dim getRecordErrorCount As Integer = 0
        Dim recsInSpecCount As Integer = 0
        Dim firstRunInSpec As Integer = -1

        TM8_gas_in_spec = False
        startTime = Now
        getRecordErrorCount = 0

        ' Keep waiting until the ppm is within the limits or the timeout elapses
        While Not TM8_gas_in_spec And Now.Subtract(startTime).TotalHours < timeout
            Try
                Application.DoEvents()
                'Get TM8 rec Data
                If Not TM8.GetRecdata(Tm8RecData, 3) Then
                    Form1.AppendText(TM8.ErrorMsg)
                    ' Allow upto 10 error counts before bailing out
                    If (getRecordErrorCount < 10) Then
                        getRecordErrorCount += 1
                        Continue While
                    Else
                        Return False
                    End If
                End If
                ' Only process if the table contains data
                If Tm8RecData.Rows.Count > 0 Then
                    TM8_gas_ppm = Tm8RecData.Rows(0)("gas_ppm")
                    Form1.TM8_PPM.Text = TM8_gas_ppm.ToString
                    If TM8_gas_ppm >= LL And TM8_gas_ppm <= UL Then
                        ' Record the time of the first within spec occurence 
                        If recsInSpecCount = 0 Then
                            TM8_gas_start_in_spec = Tm8RecData.Rows(0)("RecTime")
                        End If
                        ' Count the number of runs in specs
                        If Tm8RecData.Rows(0)("Run#") > firstRunInSpec Then
                            firstRunInSpec = Tm8RecData.Rows(0)("Run#")
                            recsInSpecCount += 1
                        End If
                        ' Need to have two runs in specs to make sure the gas concentration is stable at the setting
                        If recsInSpecCount >= 2 Then
                            TM8_gas_in_spec = True
                            Exit While
                        End If
                    End If
                End If

                ' Update the timeout and wait for 15 minutes before continue
                Form1.TimeoutLabel.Text = Math.Round(timeout * 60 - Now.Subtract(startTime).TotalMinutes).ToString + "m"
                CommonLib.Delay(60 * 15)
            Catch ex As Exception
                ' log the exception and then allow the process to continue
                Form1.AppendText("GDC.WaitForTm8PpmInSpec() caught exception: " + ex.Message)
            End Try
        End While

        If Not TM8_gas_in_spec Then
            Form1.AppendText("Timeout waiting for TM8 gas-in-gas to be between " + LL.ToString + " & " + UL.ToString + " ppm")
        Else
            Form1.AppendText("TM8_gas_start_in_spec = " + TM8_gas_start_in_spec.ToString + ", and recsInSpecCount = " + recsInSpecCount.ToString)
        End If

        Return TM8_gas_in_spec

    End Function

    Function WaitForTm8PpmInSpec0(ByVal LL As Double, ByVal UL As Double, ByVal timeout As Double) As Boolean
        Dim Tm8RecData As DataTable
        Dim TM8 As New TM8
        Dim startTime As DateTime
        Dim TM8_gas_ppm As Double

        Try
            TM8_gas_in_spec = False
            startTime = Now
            'While Not TM8_gas_in_spec And Now.Subtract(startTime).TotalHours < timeout
            While Now.Subtract(startTime).TotalHours < timeout
                Application.DoEvents()
                'Get TM8 rec Data
                If Not TM8.GetRecdata(Tm8RecData) Then
                    Form1.AppendText(TM8.ErrorMsg)
                    Return False
                End If
                TM8_gas_ppm = Tm8RecData.Rows(0)("gas_ppm")
                Form1.TM8_PPM.Text = TM8_gas_ppm.ToString
                If TM8_gas_ppm >= LL And TM8_gas_ppm <= UL Then
                    TM8_gas_in_spec = True
                End If
                If Not TM8_gas_in_spec Then
                    CommonLib.Delay(60 * 15)
                    Form1.TimeoutLabel.Text = Math.Round(Timeout_GDA * 60 - Now.Subtract(startTime).TotalMinutes).ToString + "m"
                Else
                    For Each row In Tm8RecData.Rows
                        If row("gas_ppm") >= LL And row("gas_ppm") <= UL Then
                            TM8_gas_start_in_spec = row("RecTime")
                        Else
                            Exit For
                        End If
                    Next
                End If
            End While
            If Not TM8_gas_in_spec Then
                Form1.AppendText("Timeout waiting for TM8 gas-in-gas to be between " + LL.ToString + " & " + UL.ToString + " ppm")
                Return False
            End If
            Form1.AppendText("TM8 in spec time = " + TM8_gas_start_in_spec.ToString)
        Catch ex As Exception
            Form1.AppendText("GDC.WaitForTm8PpmInSpec() caught exception: " + ex.Message)
            Return False
        End Try

        Return True
    End Function
End Class
