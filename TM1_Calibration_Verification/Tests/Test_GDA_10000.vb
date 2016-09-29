Partial Class Tests
    Public Shared Function Test_GDA_10000() As Boolean
        Dim SF As New SerialFunctions
        Dim TM8 As New TM8
        Dim gda As New GDA
        Dim inSpecCount As Integer = 0

        ' Set the gas to the specified level of 10000ppm
        If Not gda.SetGDA(10000) Then
            Return False
        End If

        ' Prompt the user to check the gas source if after Timeout_GDA the gas concentration is not at the specified level
        For i As Integer = 0 To 1
            If Not gda.WaitForTm8PpmInSpec(8500.0, 11500.0, Timeout_GDA) Then
                If i = 1 Then
                    Return False
                ElseIf i = 0 Then
                    If MessageBox.Show("Check gas sources and click 'Retry' to retry or 'Cancel' to abort test.",
                                    "H2 is not at the 10000ppm", MessageBoxButtons.RetryCancel) = DialogResult.Cancel Then
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
End Class