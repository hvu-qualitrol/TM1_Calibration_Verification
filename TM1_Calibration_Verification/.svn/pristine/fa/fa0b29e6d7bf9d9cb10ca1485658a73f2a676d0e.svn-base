Partial Class Tests
    Public Shared Function Test_GDA_1000() As Boolean
        Dim SF As New SerialFunctions
        Dim TM8 As New TM8
        Dim gda As New GDA
        Dim inSpecCount As Integer = 0

        ' Set the gas to the specified level of 1000ppm
        If Not gda.SetGDA(1000) Then
            Return False
        End If

        ' Prompt the user to check the gas source if after Timeout_GDA the gas concentration is not at the specified level
        For i As Integer = 0 To 1
            If Not gda.WaitForTm8PpmInSpec(800.0, 1200.0, Timeout_GDA) Then
                If i = 1 Then
                    Return False
                ElseIf i = 0 Then
                    If MessageBox.Show("Check gas sources and click 'Retry' to retry or 'Cancel' to abort test.",
                                    "H2 is not at the 1000ppm", MessageBoxButtons.RetryCancel) = DialogResult.Cancel Then
                        Return False
                    End If
                End If
            Else
                inSpecCount += 1
            End If

            ' Exit if having one cycle meet the specs
            If inSpecCount >= 1 Then Exit For
        Next

        Return True

    End Function

    ' For debug only 04/13/2015
    Public Shared Function TestGDA1000() As Boolean
        Dim SF As New SerialFunctions
        Dim TM8 As New TM8
        Dim gda As New GDA
        Dim inSpecCount As Integer = 0

        ' Prompt the user to check the gas source if after 4 hours the gas concentration is not at the specified level
        For i As Integer = 0 To 9
            If Not gda.WaitForTm8PpmInSpecDebug(800.0, 1200.0, 0.05) Then
                If i = 9 Then
                    Return False
                ElseIf i = 0 Then
                    If MessageBox.Show("Check gas sources and click 'Retry' to retry or 'Cancel' to abort test.",
                                    "H2 is not at the 1000ppm", MessageBoxButtons.RetryCancel) = DialogResult.Cancel Then
                        Return False
                    End If
                End If
            Else
                inSpecCount += 1
            End If

            ' Exit if four cycles have met the specs
            If inSpecCount >= 8 Then Exit For
        Next

        Return True

    End Function

End Class