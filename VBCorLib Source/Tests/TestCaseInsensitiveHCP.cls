VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestCaseInsensitiveHCP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestCaseInsensitiveHCP
'
Option Explicit

Private mHcp As CaseInsensitiveHashCodePrvdr


Public Sub Setup()
    Set mHcp = New CaseInsensitiveHashCodePrvdr
End Sub

Public Sub TestStrings()
    Assert.Equals mHcp.GetHashCode("abc"), mHcp.GetHashCode("ABC")
    Assert.Equals mHcp.GetHashCode("aBc"), mHcp.GetHashCode("AbC")
End Sub

Public Sub TestIntegers()
    Assert.Equals 10, mHcp.GetHashCode(10)
End Sub

Public Sub TestObjects()
    Dim o As cObject
    
    Set o = New Int32
    
    Assert.Equals o.GetHashCode, mHcp.GetHashCode(o)
End Sub

