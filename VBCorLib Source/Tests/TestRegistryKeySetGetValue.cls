VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestRegistryKeySetGetValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestRegistryKeySetGetValue
'
Option Explicit

Private Key As RegistryKey

Public Sub Setup()
    Set Key = Registry.CurrentUser.CreateSubKey("software\vb and vba program settings\vbcorlib_SetAndGetTests")
End Sub

Public Sub Teardown()
    Key.CloseKey
    Registry.CurrentUser.DeleteSubKeyTree "software\vb and vba program settings\vbcorlib_SetAndGetTests"
End Sub


Public Sub TestMissingNoDefault()
    Assert.IsEmpty Key.GetValue("Missing")
End Sub

Public Sub TestString()
    Key.SetValue "Name", "Kelly"
    Assert.Equals "Kelly", Key.GetValue("Name")
End Sub

Public Sub TestStringMissingWithDefault()
    Assert.Equals "Default", Key.GetValue("Missing", "Default")
End Sub

Public Sub TestLong()
    Key.SetValue "Age", 36
    Assert.Equals 36, Key.GetValue("Age")
End Sub

Public Sub TestLongMissingWithDefault()
    Assert.Equals 25, Key.GetValue("Missing", 25)
End Sub

Public Sub TestBinary()
    Dim b() As Byte
    
    Key.SetValue "Bytes", cArray.NewArray(ciByte, 1, 2, 3)
    b = Key.GetValue("Bytes")
    Assert.Equals 0, LBound(b), "Should have an lbound of 0."
    Assert.Equals 2, UBound(b), "Should have a ubound of 2."
    Assert.Equals 1, b(0)
    Assert.Equals 2, b(1)
    Assert.Equals 3, b(2)
End Sub

Public Sub TestBinaryMissingWithDefault()
    Dim b() As Byte
    
    b = Key.GetValue("Missing", cArray.NewArray(ciByte, 3, 2, 1))
    Assert.Equals 0, LBound(b), "should have an lbound of 0."
    Assert.Equals 2, UBound(b), "should have a ubound of 2."
    Assert.Equals 3, b(0)
    Assert.Equals 2, b(1)
    Assert.Equals 1, b(2)
End Sub

Public Sub TestMultiString()
    Dim s() As String
    
    Key.SetValue "Strings", Split("one,two,three", ",")
    s = Key.GetValue("Strings")
    Assert.Equals 0, LBound(s), "should have an lbound of 0."
    Assert.Equals 2, UBound(s), "should have a ubound of 2."
    Assert.Equals "one", s(0)
    Assert.Equals "two", s(1)
    Assert.Equals "three", s(2)
End Sub

Public Sub TestMultiStringMissingWithDefault()
    Dim s() As String
    
    s = Key.GetValue("Missing", Split("three,two,one", ","))
    Assert.Equals 0, LBound(s), "should have an lbound of 0."
    Assert.Equals 2, UBound(s), "should have a ubound of 2."
    Assert.Equals "three", s(0)
    Assert.Equals "two", s(1)
    Assert.Equals "one", s(2)
End Sub

Public Sub TestWrongDataType()
    On Error GoTo errTrap
    Key.SetValue "Currency", 1@
    On Error GoTo 0
    Assert.Fail "An exception should be thrown."
    
errTrap:
    Dim ex As Exception
    Assert.IsTrue Catch(ex), "An exception should be caught."
    Assert.IsTrue TypeOf ex Is ArgumentException, "Wrong exception type caught."
End Sub

Public Sub TestObjectToString()
    Dim ts As TimeSpan
    
    Set ts = TimeSpan.FromHours(25)
    
    Key.SetValue "Object", ts
    Assert.Equals ts.ToString, Key.GetValue("Object")
End Sub

Public Sub TestKeyDefaultValue()
    Key.SetValue "", "My Default Value"
    Assert.Equals "My Default Value", Key.GetValue("")
End Sub
