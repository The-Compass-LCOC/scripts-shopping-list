Option Explicit

On Error Resume Next

ShoppingListMailMerge

Sub ShoppingListMailMerge() 

  Dim xlApp 
  Dim xlBook
  Dim xlBook2
  Dim wdApp
  Dim objDoc
  Dim dirShell, curDir

  Set dirShell = CreateObject("WScript.Shell")
  curDir = dirShell.CurrentDirectory
  Set xlApp = CreateObject("Excel.Application") 
  Set xlBook = xlApp.Workbooks.Open(curDir & "\Shopping List.xlsx", 0, false)
  Set xlBook2 = xlApp.Workbooks.Open(curDir & "\shopping-list-find-and-replace-macro.xlsm", 0, True)
  
  xlApp.Visible = True
  xlApp.Run "'shopping-list-find-and-replace-macro.xlsm'!find_and_replace"
  xlBook.Save
  xlBook.Close False
  xlBook2.Close False
  xlApp.Quit

  Set xlBook = Nothing 
  Set xlApp = Nothing 

  Set wdApp = CreateObject("Word.Application") 
  Set objDoc = wdApp.Documents.Open(curDir & "\Shopping List - Mail Merge.docx")
  
  objDoc.MailMerge.OpenDataSource(curDir & "\Shopping List.xlsx")
  objDoc.MailMerge.Execute
  objDoc.Close False

  wdApp.Visible = True

  Set objWord=Nothing
  Set objFso=Nothing 

End Sub 