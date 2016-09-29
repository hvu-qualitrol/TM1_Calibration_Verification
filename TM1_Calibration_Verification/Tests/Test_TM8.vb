


Partial Class Tests
    Public Shared Function Test_TM8() As Boolean
        Dim TM8 As New TM8
        Dim RecData As DataTable

        If Not TM8.GetRecdata(RecData) Then
            Form1.AppendText(TM8.ErrorMsg)
            Return False
        End If

        Form1.AppendText("row count " + RecData.Rows.Count.ToString)
        For Each Row In RecData.Rows
            Form1.AppendText(Row("RecTime").ToString + " " + Row("run#").ToString + " " + Row("ppm").ToString)
        Next
        'Dim TM8_Telnet As New Telnet

        'If Not TM8_Telnet.Open(Form1.TM8_SN.Text, 23) Then
        '    Form1.AppendText("Could not open telnet to system board at " + Form1.TM8_SN.Text + vbCr)
        '    Return False
        'End If

        'If Not TM8_Telnet.Command("rec -L SAMPLE 24") Then
        '    Form1.AppendText("Problem sending cmd 'rec -L 24'")
        '    Return False
        'End If
        'Form1.AppendText(TM8_Telnet.CmdResult)
        'TM8_Telnet.CloseTelnet()


        Return True
    End Function

End Class
