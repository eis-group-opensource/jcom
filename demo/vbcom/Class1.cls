VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function testByte(ByVal a As Byte, ByRef b As Byte) As Byte
    b = a
    testByte = b
End Function

Public Function testInteger(ByVal a As Integer, ByRef b As Integer) As Integer
    b = a
    testInteger = b
End Function

Public Function testLong(ByVal a As Long, ByRef b As Long) As Long
    b = a
    testLong = b
End Function

Public Function testSingle(ByVal a As Single, ByRef b As Single) As Single
    b = a
    testSingle = b
End Function

Public Function testDouble(ByVal a As Double, ByRef b As Double) As Double
    b = a
    testDouble = b
End Function


Public Function testBoolean(ByVal a As Boolean, ByRef b As Boolean) As Boolean
    b = a
    testBoolean = b
End Function

Public Function testString(ByVal a As String, ByRef b As String) As String
    b = a
    testString = b
End Function

Public Function testDate(ByVal a As Date, ByRef b As Date) As Date
    b = a
    testDate = b
End Function

Public Function testCurrency(ByVal a As Currency, ByRef b As Currency) As Currency
    b = a
    testCurrency = b
End Function

Public Sub testVoid()
    Call MsgBox("testVoid")
End Sub


