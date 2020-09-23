VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestGuid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestGuid
'
Option Explicit

Private Const TEST_GUID As String = "{C200E360-38C5-11CE-AE62-08002B2B79EF}"
Private Const TEST_GUID_LOWER As String = "{C200E360-38C5-11CE-AE62-08002B2B79EE}"
Private Const TEST_GUID_UPPER As String = "{C200E360-38C5-11CE-AE62-08002B2B79FF}"


Public Sub TestToString()
    Dim s As String
    Dim g As VBCorLib.Guid
    
    Set g = Guid.Parse(TEST_GUID)
    
    Assert.Equals "{C200E360-38C5-11CE-AE62-08002B2B79EF}", g.ToString
    Assert.Equals "{C200E360-38C5-11CE-AE62-08002B2B79EF}", g.ToString("B")
    Assert.Equals "(C200E360-38C5-11CE-AE62-08002B2B79EF)", g.ToString("P")
    Assert.Equals "C200E360-38C5-11CE-AE62-08002B2B79EF", g.ToString("D")
    Assert.Equals "C200E36038C511CEAE6208002B2B79EF", g.ToString("N")
    
End Sub

Public Sub TestParse()
    Dim g As VBCorLib.Guid
    
    Set g = Guid.Parse("{C200E360-38C5-11CE-AE62-08002B2B79EF}")
    Assert.Equals TEST_GUID, g.ToString
    
    Set g = Guid.Parse("(C200E360-38C5-11CE-AE62-08002B2B79EF)")
    Assert.Equals TEST_GUID, g.ToString
    
    Set g = Guid.Parse("C200E360-38C5-11CE-AE62-08002B2B79EF")
    Assert.Equals TEST_GUID, g.ToString
    
    Set g = Guid.Parse("C200E36038C511CEAE6208002B2B79EF")
    Assert.Equals TEST_GUID, g.ToString

End Sub

Public Sub TestEquals()
    Dim g1 As VBCorLib.Guid
    Dim g2 As VBCorLib.Guid
    
    Set g1 = Guid.Parse(TEST_GUID)
    Set g2 = Guid.Parse(TEST_GUID)
    
    Assert.IsTrue g1.Equals(g2)
End Sub

Public Sub TestNotEquals()
    Dim g1 As VBCorLib.Guid
    Dim g2 As VBCorLib.Guid
    
    Set g1 = Guid.NewGuid
    Set g2 = Guid.Parse(TEST_GUID)
    
    Assert.IsFalse g1.Equals(g2)
End Sub

Public Sub TestCompareToEqual()
    Dim g1 As VBCorLib.Guid
    Dim g2 As VBCorLib.Guid
    
    Set g1 = Guid.Parse(TEST_GUID)
    Set g2 = Guid.Parse(TEST_GUID)
    
    Assert.Equals 0, g1.CompareTo(g2)
    Assert.Equals 0, g2.CompareTo(g1)
End Sub

Public Sub TestCompareTo()
    Dim lo As VBCorLib.Guid
    Dim hi As VBCorLib.Guid
    
    Set lo = Guid.Parse(TEST_GUID_LOWER)
    Set hi = Guid.Parse(TEST_GUID_UPPER)
    
    Assert.IsTrue lo.CompareTo(hi) < 0
    Assert.IsTrue hi.CompareTo(lo) > 0
End Sub

Public Sub TestGetHashCode()
    Dim g As VBCorLib.Guid
    
    Set g = Guid.Parse(TEST_GUID)
    Assert.Equals -221908415, g.GetHashCode
End Sub

Public Sub TestToByteArray()
    Dim g As VBCorLib.Guid
    Dim b() As Byte
    
    Set g = Guid.Parse(TEST_GUID)
    b = g.ToByteArray
    Assert.Equals &H60, b(0)
    Assert.Equals &HE3, b(1)
    Assert.Equals 0, b(2)
    Assert.Equals &HC2, b(3)
    Assert.Equals &HC5, b(4)
    Assert.Equals &H38, b(5)
    Assert.Equals &HCE, b(6)
    Assert.Equals &H11, b(7)
    Assert.Equals &HAE, b(8)
    Assert.Equals &H62, b(9)
    Assert.Equals &H8, b(10)
    Assert.Equals &H0, b(11)
    Assert.Equals &H2B, b(12)
    Assert.Equals &H2B, b(13)
    Assert.Equals &H79, b(14)
    Assert.Equals &HEF, b(15)
End Sub

Public Sub TestFromByteArray()
    Dim g As VBCorLib.Guid
    Dim b() As Byte
    
    b = carray.NewArray(ciByte, &H60, &HE3, &H0, &HC2, &HC5, &H38, &HCE, &H11, &HAE, &H62, &H8, &H0, &H2B, &H2B, &H79, &HEF)
    
    Set g = Guid.FromByteArray(b)
    
    Assert.Equals "{C200E360-38C5-11CE-AE62-08002B2B79EF}", g.ToString
End Sub

Public Sub TestPersistence()
    Dim g1 As VBCorLib.Guid
    Dim g2 As VBCorLib.Guid
    Dim pp1 As New PropertyBag
    Dim pp2 As New PropertyBag
    
    Set g1 = Guid.NewGuid
    pp1.WriteProperty "g", g1
    pp2.Contents = pp1.Contents
    
    Set g2 = pp2.ReadProperty("g")
    
    Assert.IsTrue g2.Equals(g1), "Guids should be equal."
    
End Sub
