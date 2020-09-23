VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestMathExt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestMathExt
'
Option Explicit

Public Sub TestCeiling()
    Assert.Equals 3#, Ceiling(2.9)
    Assert.Equals -2#, Ceiling(-2.9)
End Sub

Public Sub TestFloor()
    Assert.Equals 2#, Floor(2.9)
    Assert.Equals -3#, Floor(-2.9)
End Sub

Public Sub TestAcos()
    Assert.IsCloseTo 1.33403041100045, ACos(0.23456), 0.000000001
End Sub

Public Sub TestMax()
    Assert.Equals 5, Max(2, 5)
    Assert.Equals 5, Max(5, 2)
End Sub

Public Sub TestMaxObject()
    Dim x As New Int32
    Dim y As New Int32
    
    x.mValue = 2
    y.mValue = 5
    
    Assert.Equals y, Max(x, y)
    Assert.Equals y, Max(y, x)
End Sub

Public Sub TestMin()
    Assert.Equals 2, Min(2, 5)
    Assert.Equals 2, Min(5, 2)
End Sub

Public Sub TestMinObject()
    Dim x As New Int32
    Dim y As New Int32
    
    x.mValue = 2
    y.mValue = 5
    
    Assert.Equals x, Min(x, y)
    Assert.Equals x, Min(y, x)
End Sub
