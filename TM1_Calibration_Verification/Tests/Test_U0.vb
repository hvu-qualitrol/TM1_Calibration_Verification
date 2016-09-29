Partial Class Tests
    Public Shared Function Test_U0() As Boolean
        Dim Success As Boolean
        Dim results As ReturnResults
        Dim SF As New SerialFunctions
        Dim h2scan As New H2SCAN_debug
        Dim AllFailed As Boolean = True

        ' login to UUT's
        Success = True
        For Each UUT In UUTs
            ' Skip failing units
            If UUT("SN").Text = "" Or UUT("FAILED") Then Continue For

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

                Form1.AppendText("CALIBRATION STEP U0", UUT:=UUT)
                If Not h2scan.Open(UUT) Then
                    Form1.AppendText(h2scan.ErrorMsg, UUT:=UUT)
                    UUT("FAILED") = True
                    UUT("TAB").ImageIndex = StatusColor.FAILED
                    Success = False
                    Continue For
                End If
                If Not h2scan.U0(UUT) Then
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
                Form1.AppendText(UUT("SN").Text + ": Test_U0() caught" + ex.ToString, UUT:=UUT)
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

    Public Shared Function Test_U00() As Boolean
        Dim Success As Boolean
        Dim results As ReturnResults
        Dim SF As New SerialFunctions
        Dim h2scan As New H2SCAN_debug
        Dim AllFailed As Boolean = True


        ' login to UUT's
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
                Form1.AppendText(UUT("SN").Text + ": Test_U0() caught" + ex.ToString, UUT:=UUT)
                UUT("FAILED") = True
                UUT("TAB").ImageIndex = StatusColor.FAILED
                Success = False
            End Try
        Next

        ' U0
        For Each UUT In UUTs
            If UUT("SN").Text = "" Or UUT("FAILED") Then Continue For

            Try
                Form1.AppendText("CALIBRATION STEP U0", UUT:=UUT)
                If Not h2scan.Open(UUT) Then
                    Form1.AppendText(h2scan.ErrorMsg, UUT:=UUT)
                    UUT("FAILED") = True
                    UUT("TAB").ImageIndex = StatusColor.FAILED
                    Success = False
                    Continue For
                End If
                If Not h2scan.U0(UUT) Then
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
                Form1.AppendText(UUT("SN").Text + ": Test_U0() caught" + ex.ToString, UUT:=UUT)
                UUT("FAILED") = True
                UUT("TAB").ImageIndex = StatusColor.FAILED
                Success = False
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