Attribute VB_Name = "Module1"

'** VB�̃T���v���v���O����
Sub HelloJCom()
On Error GoTo ErrorHandler
  Set xlApp = CreateObject("Excel.Application")        ' EXCEL���N��
  xlApp.Visible = True
  Set xlBooks = xlApp.Workbooks()
  Set xlBook = xlBooks.Add()               ' �V�����u�b�N���쐬
  Set xlSheets = xlBook.Worksheets()
  Set xlSheet = xlSheets.Item(1)
  Set xlRanges = xlSheet.Cells()
  xlRanges.Item(1, 1).Value = "�͂��߂Ă�JCom"
  xlRanges.Item(2, 1).Value = "�����Java���珑���܂���"
  Call MsgBox("�{�^���������ƏI�����܂�")
  Call xlBook.Close(False, Null, False)
  Call xlApp.Quit
  Exit Sub
ErrorHandler:
  Call MsgBox(Err.Number & ":" & Err.Description)
End Sub

Sub main()
  Call HelloJCom
End Sub
