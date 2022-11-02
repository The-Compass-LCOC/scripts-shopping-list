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

  xlBook.Worksheets("ShoppingList").Range("A:A").EntireColumn.Delete
  xlBook.Worksheets("ShoppingList").Range("D:D").EntireColumn.Delete
  xlBook.Worksheets("ShoppingList").Range("G:G").EntireColumn.Delete
  xlBook.Worksheets("ShoppingList").Range("E:E").EntireColumn.Delete
  xlBook.Worksheets("ShoppingList").Range("E:E").EntireColumn.Delete
  xlBook.Worksheets("ShoppingList").Range("E:E").EntireColumn.Delete
  xlBook.Worksheets("ShoppingList").Range("B:B").EntireColumn.Delete
  xlBook.Worksheets("ShoppingList").Range("A:A").EntireColumn.Cut
  xlBook.Worksheets("ShoppingList").Paste xlBook.Worksheets("ShoppingList").Range("Z1")
  xlBook.Worksheets("ShoppingList").Range("C:C").EntireColumn.Cut
  xlBook.Worksheets("ShoppingList").Paste xlBook.Worksheets("ShoppingList").Range("A1")
  xlBook.Worksheets("ShoppingList").Range("Z:Z").EntireColumn.Cut
  xlBook.Worksheets("ShoppingList").Paste xlBook.Worksheets("ShoppingList").Range("C1")
  xlApp.Run "'shopping-list-split-column-macro.xlsm'!split_name"
  xlApp.Run "'shopping-list-split-column-macro.xlsm'!replace_time"
  xlBook.Worksheets("ShoppingList").Range("1:1").EntireRow.Delete
  xlBook2.Close False

  Set xlBook2 = xlApp.Workbooks.Open(curDir & "\Arrival Time Detail Report.xlsx", 0, false)

  'xlBook.Worksheets("ShoppingList").UsedRange.Copy
  MsgBox xlBook.Worksheets("ShoppingList").UsedRange.Rows.Count
  xlBook2.Worksheets("Arrival Time Detail Report (TOM").Range("B11").EntireRow.Insert
  'xlBook2.Worksheets("Arrival Time Detail Report (TOM").Range("B11").Insert

  Set xlApp = Nothing 
  Set xlBook = Nothing 
  Set xlBook2 = Nothing 
  Set wdApp = Nothing
  Set objDoc=Nothing
  Set curDir=Nothing 

End Sub 