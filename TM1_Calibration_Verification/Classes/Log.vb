Imports System.IO

Public Module Log
    Private _fileName As String

    Public Property FileName() As String
        Get
            Return _fileName
        End Get
        Set(value As String)
            _fileName = value
        End Set
    End Property

    Public Sub Write(line As String)
        Dim path As String
        path = "C:\temp\" & _fileName & ".log"
        Using _file As New StreamWriter(path, True)
            _file.Write(line)
            _file.Close()
        End Using

    End Sub
    Public Sub WriteLine(line As String)
        Dim path As String
        path = "C:\temp\" & _fileName & ".log"
        Using _file As New StreamWriter(path, True)
            _file.WriteLine(Format(Date.UtcNow, "yyyy-MM-dd-HHmmss") + " - " + line)
            _file.Close()
        End Using

    End Sub
End Module
