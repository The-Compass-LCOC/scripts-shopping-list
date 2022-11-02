Option Explicit

On Error Resume Next

ShoppingListMailMerge

Sub ShoppingListMailMerge() 

  Dim xlApp 
  Dim xlBook
  Dim xlBook2
  Dim wdApp
  Dim objDoc
  Dim curDir
  Dim usedRange
  Dim usedRange2

  curDir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
  Set xlApp = CreateObject("Excel.Application") 
  Set xlBook = xlApp.Workbooks.Open(curDir & "\Shopping List.xlsx", 0, false)
  Set xlBook2 = xlApp.Workbooks.Open(curDir & "\shopping-list-find-and-replace-macro.xlsm", 0, True)
  
  xlApp.Visible = True
  xlApp.Run "'shopping-list-find-and-replace-macro.xlsm'!find_and_replace"
  xlBook.Save
  xlBook.Close False
  xlBook2.Close False
  xlApp.Quit


  Set wdApp = CreateObject("Word.Application") 
  Set objDoc = wdApp.Documents.Open(curDir & "\shopping-list-mail-merge.docx")
  
  objDoc.MailMerge.OpenDataSource(curDir & "\Shopping List.xlsx")
  objDoc.MailMerge.Execute
  objDoc.Close False
  wdApp.Visible = True

  Set xlBook = xlApp.Workbooks.Open(curDir & "\Shopping List.xlsx", 0, false)
  Set xlBook2 = xlApp.Workbooks.Open(curDir & "\shopping-list-split-column-macro.xlsm", 0, True)
  xlApp.Visible = True

  xlBook.Worksheets(1).Range("A:A").EntireColumn.Delete
  xlBook.Worksheets(1).Range("D:D").EntireColumn.Delete
  xlBook.Worksheets(1).Range("G:G").EntireColumn.Delete
  xlBook.Worksheets(1).Range("E:E").EntireColumn.Delete
  xlBook.Worksheets(1).Range("E:E").EntireColumn.Delete
  xlBook.Worksheets(1).Range("E:E").EntireColumn.Delete
  xlBook.Worksheets(1).Range("B:B").EntireColumn.Delete
  xlBook.Worksheets(1).Range("A:A").EntireColumn.Cut
  xlBook.Worksheets(1).Paste xlBook.Worksheets(1).Range("Z1")
  xlBook.Worksheets(1).Range("C:C").EntireColumn.Cut
  xlBook.Worksheets(1).Paste xlBook.Worksheets(1).Range("A1")
  xlBook.Worksheets(1).Range("Z:Z").EntireColumn.Cut
  xlBook.Worksheets(1).Paste xlBook.Worksheets(1).Range("C1")
  xlApp.Run "'shopping-list-split-column-macro.xlsm'!split_name"
  xlApp.Run "'shopping-list-split-column-macro.xlsm'!replace_time"
  xlBook.Worksheets(1).Range("1:1").EntireRow.Delete
  xlBook.Worksheets(1).Range("D:D").EntireColumn.Cut
  xlBook.Worksheets(1).Paste xlBook.Worksheets(1).Range("Z1")
  xlBook.Worksheets(1).Range("B:B").EntireColumn.Cut
  xlBook.Worksheets(1).Paste xlBook.Worksheets(1).Range("D1")
  xlBook.Worksheets(1).Range("C:C").EntireColumn.Cut
  xlBook.Worksheets(1).Paste xlBook.Worksheets(1).Range("B1")
  xlBook.Worksheets(1).Range("Z:Z").EntireColumn.Cut
  xlBook.Worksheets(1).Paste xlBook.Worksheets(1).Range("C1")
  xlBook.Worksheets(1).Cells.Interior.ColorIndex = 6
  xlBook2.Close False

  Set xlBook2 = xlApp.Workbooks.Open(curDir & "\Arrival Time Detail Report.xlsx", 0, false)

  xlBook.Worksheets(1).UsedRange.Copy
  xlBook2.Worksheets(1).Range("A" & xlBook2.Worksheets(1).UsedRange.Rows.Count + 1).pasteSpecial
  xlBook2.Worksheets(1).Range("A" & xlBook2.Worksheets(1).UsedRange.Rows.Count + 1).Value = "Total Shopping"
  xlBook2.Worksheets(1).Range("E" & xlBook2.Worksheets(1).UsedRange.Rows.Count).Value = xlBook.Worksheets(1).UsedRange.Rows.Count
  xlBook2.Worksheets(1).Range("A" & xlBook2.Worksheets(1).UsedRange.Rows.Count + 1).Value = "Total Pickup"
  xlBook2.Worksheets(1).Range("E" & xlBook2.Worksheets(1).UsedRange.Rows.Count).Value = "=" & xlBook2.Worksheets(1).UsedRange.Rows.Count - xlBook.Worksheets(1).UsedRange.Rows.Count -3 & " - COUNTBLANK(B2:B101)"
  xlBook2.Worksheets(1).Range("A" & xlBook2.Worksheets(1).UsedRange.Rows.Count + 1).Value = "Total"
  xlBook2.Worksheets(1).Range("E" & xlBook2.Worksheets(1).UsedRange.Rows.Count).Value = "=" & xlBook.Worksheets(1).UsedRange.Rows.Count + xlBook2.Worksheets(1).UsedRange.Rows.Count - xlBook.Worksheets(1).UsedRange.Rows.Count -4 & " - COUNTBLANK(B2:B101)"
  xlApp.CutCopyMode = False
  xlBook.Close False
  xlBook2.Worksheets(1).Cells.Font.Size = 12

  Set xlApp = Nothing 
  Set xlBook = Nothing 
  Set xlBook2 = Nothing 
  Set wdApp = Nothing
  Set objDoc=Nothing
  Set curDir=Nothing 

End Sub 