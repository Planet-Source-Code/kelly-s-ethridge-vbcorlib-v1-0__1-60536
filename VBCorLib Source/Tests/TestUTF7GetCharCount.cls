VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestUTF7GetCharCount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestUTF7GetCharCount
'
Option Explicit

Private e As UTF7Encoding
Private b() As Byte

Public Sub Setup()
    Set e = New UTF7Encoding
End Sub

Public Sub TestGetPreamble()
    Dim b() As Byte
    
    b = e.GetPreamble
    Assert.IsFalse carray.IsNull(b)
    Assert.Equals 0, carray.GetLength(b)
End Sub

Public Sub TestPlus()
    b = carray.NewArray(ciByte, 43)
    Assert.Equals 0, e.GetCharCount(b)
End Sub

Public Sub TestOneNonDirect()
    b = carray.NewArray(ciByte, 33)
    Assert.Equals 1, e.GetCharCount(b)
End Sub

Public Sub TestPlusMinus()
    b = carray.NewArray(ciByte, 43, 45)
    Assert.Equals 1, e.GetCharCount(b)
End Sub

Public Sub TestSimple()
    b = e.GetBytes("hello")
    Assert.Equals 5, e.GetCharCount(b)
End Sub

Public Sub TestJapaneseWord_nihongo()
    Dim ch() As Integer
    ch = carray.NewArray(ciInteger, &H65E5, &H672C, &H8A9E)
    b = e.GetBytes(ch)
    Assert.Equals 3, e.GetCharCount(b)
End Sub

Public Sub TestCombo()
    b = e.GetBytes(ChrW$(&H65E5) & "hi!" & ChrW$(&H8A9E))
    Assert.Equals 5, e.GetCharCount(b)
End Sub

Public Sub TestDecoder()
    Dim ch(10) As Integer
    Dim i As Long
    Dim c As Decoder
    
    Set c = e.GetDecoder
    
    b(0) = 43
    Assert.Equals 0, c.GetCharCount(b, 0, 1)
    b(0) = 33
    Assert.Equals 1, c.GetCharCount(b, 0, 1)
    
    b = e.GetBytes(ChrW$(&H6535))
    Assert.Equals 0, c.GetCharCount(b, 0, 1)
    Assert.Equals 0, c.GetCharCount(b, 0, 2)
    Assert.Equals 0, c.GetCharCount(b, 0, 3)
    Assert.Equals 1, c.GetCharCount(b, 0, 4)
    Assert.Equals 1, c.GetCharCount(b, 0, 5)
End Sub
