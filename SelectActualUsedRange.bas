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

