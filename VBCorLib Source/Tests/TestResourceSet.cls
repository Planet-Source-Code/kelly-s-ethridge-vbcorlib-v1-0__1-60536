VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestResourceSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestResourceSet
'
Option Explicit

Private mResourceSet As ResourceSet


Public Sub Setup()
    Set mResourceSet = Cor.NewResourceSet(Cor.NewCultureInfo("en-US"), Path.Combine(App.Path, "VBCorLibTests.res"))
End Sub

Public Sub Teardown()
    mResourceSet.CloseResourceSet
End Sub

Public Sub TestGetString()
    Assert.Equals "hello", mResourceSet.GetString(101)
End Sub

Public Sub TestGetStringNotFound()
    Assert.Equals "", mResourceSet.GetString(999)
End Sub

Public Sub TestGetObjectCustom()
    Dim b() As Byte
    Dim i As Long
    
    b = mResourceSet.GetObject(101, "Custom")
    Assert.Equals 10, carray.GetLength(b)
    For i = 0 To 9
        Assert.Equals i + 48, b(i)
    Next i
End Sub

Public Sub TestGetObjectNotFound()
    Assert.IsEmpty mResourceSet.GetObject(101, "NotFound")
End Sub
