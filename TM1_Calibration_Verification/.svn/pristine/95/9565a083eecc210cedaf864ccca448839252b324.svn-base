Imports System.Threading
Imports System.IO
Imports Microsoft.Office.Interop
Public Class WordWrapper
    Dim wordApp As Word.Application = Nothing
    Dim readOnlyProp As Object = False
    Event NotFound(ByVal Message As String)

    Public ReadOnly Property Version() As Double
        Get
            Return Convert.ToDouble(wordApp.Version)
        End Get
    End Property
    Public Function CreateWordDoc(ByVal FilePath As String) As Boolean
        Try
            Dim fileName As Object = FilePath
            Dim isVisible As Object = True
            ' Here is the way to handle parameters you don't care about in .NET
            Dim missing As Object = System.Reflection.Missing.Value
            ' Open the document that was chosen by the dialog    

            wordApp.Documents.Add(fileName)
            'wordDoc = wordApp.Documents.Open(fileName, missing, True, missing, missing, missing, missing, missing, missing, missing, missing, isVisible, missing, missing, missing, missing) ' Activate the document so it shows up in front aDoc.Activate();                        
        Catch ex As Exception
            MessageBox.Show("CreateWordDoc() caught " + ex.Message, "WordWrapper Error")
            Return False
        End Try

        Return True

    End Function

    Public Sub Finished(ByVal FilePath As String)
        Dim fileName As Object = FilePath
        Try
            Dim pathName As String
            pathName = FilePath.Substring(0, FilePath.Length - (FilePath.Length - FilePath.LastIndexOf("\"c) - 1))
            If Directory.Exists(pathName) = False Then
                Directory.CreateDirectory(pathName)
            End If
            wordApp.Application.ActiveDocument.SaveAs(fileName)
            'fileName = fileName.ToString.Replace("doc", "pdf")

            wordApp.Application.ActiveDocument.Close()
            wordApp.Application.Quit()
            'wordApp = Nothing
        Catch ex As Exception
            'OCPError(ex.Message, "Word Wrapper : Finished")
        Finally
            If (wordApp IsNot Nothing) Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp)
                wordApp = Nothing
                GC.Collect()
            End If
        End Try

    End Sub

    Public Sub AddTable(ByVal TableInsert As String, ByVal dTable As DataTable, Optional ByVal Vertical As Boolean = False, Optional ByVal CellArray As ArrayList = Nothing, Optional ByVal HeaderArray As ArrayList = Nothing, Optional ByVal ScriptName As String = "")
        Try
            If dTable Is Nothing Then Exit Try

            If dTable.Rows.Count > 0 Then

                If Vertical = False And ScriptName <> "" Then
                    wordApp.Application.ActiveDocument.Content.Select()
                    With wordApp.Application.Selection.Find
                        .Text = TableInsert + "_Name"
                        If .Execute = False Then
                            'OCPError("Unable to find : " + TableInsert + "_Name")
                        Else
                            Dim myFont As New Microsoft.Office.Interop.Word.Font
                            wordApp.Application.Selection.Font.Italic = 1
                            wordApp.Application.Selection.Font.Bold = 1
                            wordApp.Application.Selection.Text = ScriptName
                        End If
                    End With
                End If

                wordApp.Application.ActiveDocument.Content.Select()
                With wordApp.Application.Selection.Find
                    .ClearFormatting()
                    .Text = TableInsert
                    .Replacement.Text = ""
                    .Forward = False
                    .Wrap = Word.WdFindWrap.wdFindStop
                    .Format = False
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                    If .Execute = False Then
                        .Forward = True
                        If .Execute = False Then
                            RaiseEvent NotFound("Unable to find " & TableInsert)
                            Exit Try
                        End If
                    End If
                End With

                Dim nRows As Int32
                Dim nCols As Int32
                nRows = dTable.Rows.Count
                nCols = dTable.Columns.Count

                If Vertical = True Then
                    'DebugLogger.Debug("Creating Vertical Table for " & TableInsert, Threading.Thread.CurrentThread.ManagedThreadId)
                    Me.TableInsertVertical(nCols, nRows, dTable, HeaderArray, CellArray)
                Else
                    'DebugLogger.Debug("Creating Horizontal Table for " & TableInsert, Threading.Thread.CurrentThread.ManagedThreadId)
                    Me.TableInsertHorizontal(nCols, nRows, dTable, HeaderArray, CellArray, ScriptName)
                End If
            End If
        Catch ex As Exception
            'OCPError(ex)
        End Try
    End Sub
    Private Sub TableInsertVertical(ByVal nCols As Int32, ByVal nRows As Int32, ByVal dTable As DataTable, ByVal HeaderArray As ArrayList, ByVal CellArray As ArrayList)
        Dim tempTable As Word.Table
        Dim I As Int32
        Dim I1 As Int32
        Dim dTableLine As Int32
        If dTable.Rows.Count = 0 Then
            wordApp.Application.Selection.Range.Text = ""
            Exit Sub
        End If
        tempTable = wordApp.Application.ActiveDocument.Tables.Add( _
                   Range:=wordApp.Application.Selection.Range, NumRows:=1, NumColumns:=2)

        tempTable.AllowAutoFit = True

        dTableLine = 1
        Try
            Application.DoEvents()

            If CellArray Is Nothing Then
                Dim colCount As Int32
                CellArray = New ArrayList
                For colCount = 0 To dTable.Columns.Count - 1
                    CellArray.Add(dTable.Columns(colCount).ColumnName)
                Next
            End If

            If HeaderArray Is Nothing Then
                HeaderArray = New ArrayList
                For I1 = 0 To CellArray.Count - 1
                    HeaderArray.Add(CellArray(I1).ToString)
                Next
            End If

            For I = 0 To nRows - 1
                For I1 = 0 To HeaderArray.Count - 1
                    If dTable.Rows(I).Item(CellArray(I1).ToString).ToString.Trim <> "NA" Then

                        Application.DoEvents()
                        If I <> nRows Then tempTable.Rows.Add()
                        If I1 = 0 Then
                            ' Create the header row.
                            tempTable.Cell(tempTable.Rows.Count - 1, 1).Shading.BackgroundPatternColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdBlack
                            tempTable.Cell(tempTable.Rows.Count - 1, 2).Shading.BackgroundPatternColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdBlack
                            tempTable.Cell(tempTable.Rows.Count - 1, 1).Range.Font.Color = Word.WdColor.wdColorWhite
                            tempTable.Cell(tempTable.Rows.Count - 1, 2).Range.Font.Color = Word.WdColor.wdColorWhite
                            tempTable.Cell(tempTable.Rows.Count - 1, 1).Range.Font.Bold = 1
                            tempTable.Cell(tempTable.Rows.Count - 1, 2).Range.Font.Bold = 1
                        End If
                        tempTable.Cell(tempTable.Rows.Count - 1, 1).Range.Text = HeaderArray(I1).ToString
                        tempTable.Cell(tempTable.Rows.Count - 1, 2).Range.Text = dTable.Rows(I).Item(CellArray(I1).ToString).ToString
                    End If
                Next
                tempTable.AllowPageBreaks = True
            Next
        Catch ex As Exception
            'OCPError(ex)
        End Try
    End Sub
    Private Sub TableInsertHorizontal(ByVal nCols As Int32, ByVal nRows As Int32, ByVal dTable As DataTable, ByVal HeaderArray As ArrayList, ByVal CellArray As ArrayList, Optional ByVal ScriptName As String = "")
        Dim tempTable As Word.Table
        Dim I1 As Int32

        If dTable.Rows.Count = 0 Then
            wordApp.Application.Selection.Range.Text = ""
            Exit Sub
        End If

        Try
            Dim dTableLine As Int32
            Application.DoEvents()
            tempTable = wordApp.Application.ActiveDocument.Tables.Add( _
                Range:=wordApp.Application.Selection.Range, NumRows:=1, NumColumns:=nCols)
            tempTable.AllowAutoFit = True

            dTableLine = 0

            If Not (HeaderArray Is Nothing) Then
                With tempTable
                    Dim i3 As Int32
                    For i3 = 1 To nCols
                        .Cell(Row:=1, Column:=i3).Range.Text = HeaderArray(i3 - 1).ToString
                    Next
                End With
                dTableLine = 1
            End If

            If CellArray Is Nothing Then
                Dim colCount As Int32
                CellArray = New ArrayList
                For colCount = 0 To dTable.Columns.Count - 1
                    CellArray.Add(dTable.Columns(colCount).ColumnName)
                Next
            End If

            For I1 = 0 To nRows - 1
                Application.DoEvents()
                Dim I3 As Int32
                If I1 <> nRows Then tempTable.Rows.Add()
                For I3 = 1 To nCols
                    tempTable.Cell(tempTable.Rows.Count, I3).Range.Text = dTable.Rows(I1).Item(CellArray(I3 - 1).ToString).ToString
                    tempTable.Cell(tempTable.Rows.Count, I3).Range.Font.Size = 8
                Next
            Next

            tempTable.Style = "Table Professional"
            tempTable.AllowPageBreaks = True
        Catch ex As Exception
            'OCPError(ex)
        End Try
    End Sub

    Public Sub ReplaceItemWith(ByVal Item As DictionaryEntry)
        Dim txt As String
        Dim tempWord As Word.Application
        Dim tempDoc As Word.Document
        Dim tempRange As Word.Range

        txt = Me.StripWhiteSpace(Item.Key.ToString.Trim)
        Try
            If wordApp.Application.ActiveDocument.Bookmarks.Exists(txt) = True Then
                Dim Rng As Microsoft.Office.Interop.Word.Range
                Dim bookMarkName As Object = CType(txt, Object)
                Application.DoEvents()
                wordApp.Application.ActiveDocument.Bookmarks.Item(bookMarkName).Select()
                Rng = wordApp.Application.Selection.Range
                Rng.Text = ""
                Dim Anchor As Object = CType(Rng, Object)
                Dim Line As Microsoft.Office.Interop.Word.Paragraph
                Select Case LCase(Item.Value.ToString.Substring(Item.Value.ToString.Length - 4))
                    Case Is = ".txt"
                        If Not System.IO.Directory.Exists("C:\Temp") Then
                            Try
                                Directory.CreateDirectory("C:\Temp")
                            Catch ex As Exception
                                'OCPError(ex)
                            End Try
                        End If
                        tempWord = New Word.Application
                        tempDoc = tempWord.Documents.Add()
                        tempRange = tempDoc.Range(Start:=0, End:=0)
                        tempDoc.Application.Selection.InsertFile(Item.Value.ToString, , False)
                        tempRange = tempDoc.Range(Start:=0)
                        tempRange.Font.Size = 9
                        If Item.Value.ToString.Contains("sample gas.txt") Or Item.Value.ToString.Contains("veri gas.txt") Then
                            For Each Line In tempDoc.Paragraphs
                                If Line.Range.Text.Contains("concentration") Or Line.Range.Text.Contains("oil ppm") Or Line.Range.Text.Contains("gas ppm") Then
                                    Line.Range.Delete()
                                End If
                            Next
                        End If
                        tempDoc.SaveAs("c:\Temp\temp_dhr.doc")
                        tempDoc.Close(True)
                        tempWord.Quit()
                        'wordApp.Application.Selection.Font.Size = 9
                        'wordApp.Application.Selection.InsertFile(Item.Value.ToString)
                        wordApp.Application.Selection.InsertFile("c:\Temp\temp_dhr.doc", , False)
                    Case Is = ".bmp"
                        Dim Shp As Word.Shape
                        Shp = wordApp.Application.ActiveDocument.Shapes.AddPicture(FileName:=Item.Value.ToString, LinkToFile:=False, SaveWithDocument:=True, Anchor:=Anchor)
                        Shp.WrapFormat.Type = Word.WdWrapType.wdWrapTight
                        'WordApp.Application.ActiveDocument.Bookmarks.Item(bookMarkName).Select()
                        'WordApp.Application.Selection.Range.Text = ""                        
                        'shp.RelativeHorizontalPosition = Word.WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage

                    Case Is = ".jpg"
                        'Dim Shp As Word.Shape
                        'Shp = WordApp.Application.ActiveDocument.Shapes.AddPicture(fileName:=Item.Value.ToString, LinkToFile:=False, SaveWithDocument:=True, Anchor:=Rng)
                        'Shp.WrapFormat.Type = Word.WdWrapType.wdWrapTight                        
                        'WordApp.Application.ActiveDocument.Bookmarks.Item(CType(txt, Object)).Delete()
                        Dim Shp As Word.Shape
                        Shp = wordApp.Application.ActiveDocument.Shapes.AddPicture(FileName:=Item.Value.ToString, SaveWithDocument:=True, Anchor:=Anchor)
                        'Shp.WrapFormat.Type = Word.WdWrapType.wdWrapTight
                        'Shp.WrapFormat.Type = Word.WdWrapType.wdWrapBehind
                        Shp.RelativeHorizontalPosition = Word.WdRelativeHorizontalPosition.wdRelativeHorizontalPositionColumn
                    Case Is = ".emf"
                        Dim Shp As Microsoft.Office.Interop.Word.Shape
                        Shp = wordApp.Application.ActiveDocument.Shapes.AddPicture(FileName:=Item.Value.ToString, LinkToFile:=False, SaveWithDocument:=True, Anchor:=Anchor)
                        Shp.WrapFormat.Type = Word.WdWrapType.wdWrapTight
                        wordApp.Application.ActiveDocument.Bookmarks.Item(CType(txt, Object)).Delete()
                    Case Is = ".doc"
                        wordApp.Application.Selection.Font.Size = 9
                        wordApp.Application.Selection.InsertFile(Item.Value.ToString)
                    Case Is = ".rtf"
                        wordApp.Application.Selection.Font.Size = 9
                        wordApp.Application.Selection.InsertFile(Item.Value.ToString)
                    Case Is = ".htm"
                        wordApp.Application.Selection.InsertFile(Item.Value.ToString)
                End Select
            Else
                RaiseEvent NotFound("Unable to find : " & txt)
            End If
        Catch ex As Exception
            'OCPError(ex)
        End Try
    End Sub

    Public Sub ReplaceItemWith(ByVal ItemToFind As String, ByVal txt As String)
        Dim forward As Boolean
        Dim backward As Boolean

        Do
            backward = False
            wordApp.Application.ActiveDocument.Content.Select()
            With wordApp.Application.Selection.Find
                .ClearFormatting()
                .Text = ItemToFind
                .Replacement.Text = txt
                .Forward = False
                .Wrap = Word.WdFindWrap.wdFindStop
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
                If .Execute = True Then
                    wordApp.Application.Selection.Text = txt
                    backward = True
                End If
            End With
        Loop Until backward = False

        Do
            forward = False
            wordApp.Application.ActiveDocument.Content.Select()
            With wordApp.Application.Selection.Find
                .ClearFormatting()
                .Text = ItemToFind
                .Replacement.Text = txt
                .Forward = True
                .Wrap = Word.WdFindWrap.wdFindStop
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
                If .Execute = True Then
                    wordApp.Application.Selection.Text = txt
                    forward = True
                End If
            End With
        Loop Until forward = False
    End Sub

    Public Sub ReplaceItemWith(ByVal HashOfSpecValues As Hashtable, Optional ByVal roundValue As Int32 = 0)
        For Each entry As DictionaryEntry In HashOfSpecValues
            Try
                If roundValue > 0 Then
                    If IsNumeric(entry.Value.ToString) Then
                        Dim numValue As Decimal = Convert.ToDecimal(entry.Value.ToString)
                        numValue = Decimal.Round(numValue, roundValue)
                        ReplaceItemWith(entry.Key.ToString(), String.Format("{0:f" + roundValue.ToString + "}", numValue))
                        'ReplaceItemWith(entry.Key.ToString(), numValue.ToString)
                    End If
                Else
                    ReplaceItemWith(entry.Key.ToString(), entry.Value.ToString())
                End If
            Catch ex As Exception
                'OCPError(ex.Message)
            End Try
        Next
    End Sub
    Private Function StripWhiteSpace(ByVal txt As String) As String
        Dim I As Int32
        Dim txt2() As String
        txt2 = txt.Split(" "c)
        txt = ""
        For I = 0 To UBound(txt2)
            txt += txt2(I)
        Next
        Return txt
    End Function

    Protected Overrides Sub Finalize()
        If (wordApp IsNot Nothing) Then
            System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp)
            wordApp = Nothing
            GC.Collect()
        End If

        MyBase.Finalize()

        'If wordApp IsNot Nothing Then
        '    Try
        '        wordApp.Application.Documents.Close()
        '        wordApp.
        '    Catch ex As Exception
        '        Form1.AppendText("WordWrapper.Finalize(): Caught " + ex.ToString)
        '    End Try
        '    Try
        '        wordApp.Quit()
        '    Catch ex As Exception
        '        Form1.AppendText("WordWrapper.Finalize(): Caught " + ex.ToString)
        '    End Try
        '    wordApp = Nothing
        'End If

    End Sub

    Public Sub New()
        wordApp = New Word.Application
    End Sub
End Class

