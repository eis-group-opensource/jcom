Attribute VB_Name = "Module1"

'** VBのサンプルプログラム
Sub HelloJCom()
On Error GoTo ErrorHandler
  Set xlApp = CreateObject("Excel.Application")        ' EXCELを起動
  xlApp.Visible = True
  Set xlBooks = xlApp.Workbooks()
  Set xlBook = xlBooks.Add()               ' 新しいブックを作成
  Set xlSheets = xlBook.Worksheets()
  Set xlSheet = xlSheets.Item(1)
  Set xlRanges = xlSheet.Cells()
  xlRanges.Item(1, 1).Value = "はじめてのJCom"
  xlRanges.Item(2, 1).Value = "これはJavaから書きました"
  Call MsgBox("ボタンを押すと終了します")
  Call xlBook.Close(False, Null, False)
  Call xlApp.Quit
  Exit Sub
ErrorHandler:
  Call MsgBox(Err.Number & ":" & Err.Description)
End Sub

Sub main()
  Call HelloJCom
End Sub
