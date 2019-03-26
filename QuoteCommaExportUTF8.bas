Attribute VB_Name = "Módulo1"
Sub QuoteCommaExportUTF8()
   ' Dimension all variables.
   Dim DestFile As String
   Dim FileNum As Integer
   Dim ColumnCount As Integer
   Dim RowCount As Integer
   
   ' Select actual Used Range
    SelectActualUsedRange
   ' Prompt user for destination file name.
   DestFile = "C:\CSV\" & InputBox("Enter the filename" _
      & Chr(10) & "(with complete path):", "NEWNEWQuote-Comma Exporter") _
      & Replace(ActiveWorkbook.Name, ".xlsx", " ") & ActiveSheet.Name _
      & ".csv"

   ' Create File
   Dim objStream
   Set objStream = CreateObject("ADODB.Stream")
   objStream.Charset = "utf-8"
   objStream.Open


   ' Turn error checking off.
   On Error Resume Next

   ' Attempt to open destination file for output.
   ' Open DestFile For Output As #FileNum

   ' If an error occurs report it and end.
   If Err <> 0 Then
      MsgBox "Cannot open filename " & DestFile
      End
   End If

   ' Turn error checking on.
   On Error GoTo 0

   ' Loop for each row in selection.
   For RowCount = 1 To Selection.Rows.Count

      ' Loop for each column in selection.
      For ColumnCount = 1 To Selection.Columns.Count

         ' Write current cell's text to file with quotation marks.
         objStream.WriteText """" & Selection.Cells(RowCount, ColumnCount).Text & """"

         ' Check if cell is in last column.
         If ColumnCount = Selection.Columns.Count Then
            ' If so, then write a blank line.
            objStream.WriteText vbCrLf
         Else
            ' Otherwise, write a comma.
            objStream.WriteText ","
            
         End If
      ' Start next iteration of ColumnCount loop.
      Next ColumnCount
   ' Start next iteration of RowCount loop.
   Next RowCount

   ' Save file.
   objStream.SaveToFile DestFile, 2
   
   MsgBox "Complete"
End Sub

Sub SelectActualUsedRange()
  Dim FirstCell As Range, LastCell As Range
  Set LastCell = Cells(Cells.Find(What:="*", SearchOrder:=xlRows, _
      SearchDirection:=xlPrevious, LookIn:=xlValues).Row, _
      Cells.Find(What:="*", SearchOrder:=xlByColumns, _
      SearchDirection:=xlPrevious, LookIn:=xlValues).Column)
  Set FirstCell = Cells(Cells.Find(What:="*", After:=LastCell, SearchOrder:=xlRows, _
      SearchDirection:=xlNext, LookIn:=xlValues).Row, _
      Cells.Find(What:="*", After:=LastCell, SearchOrder:=xlByColumns, _
      SearchDirection:=xlNext, LookIn:=xlValues).Column)
  Range(FirstCell, LastCell).Select
End Sub

