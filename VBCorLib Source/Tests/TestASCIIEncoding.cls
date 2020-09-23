VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestASCIIEncoding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestASCIIEncoding
'
Option Explicit

Private en As Encoding

Public Sub Setup()
    Set en = Encoding.ASCII
End Sub

Public Sub TestGetBytes()
    Dim b() As Byte
    
    b = en.GetBytes("hello")
    Assert.Equals 4, UBound(b)
    
End Sub
