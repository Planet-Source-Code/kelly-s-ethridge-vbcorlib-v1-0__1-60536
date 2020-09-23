VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestRegistryDeleteValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestRegistryDeleteValue
'
Option Explicit

Private Key As RegistryKey


Public Sub Setup()
    Set Key = Registry.CurrentUser.CreateSubKey("software\vb and vba program settings\VBCorLib_DeleteValue")
    Key.SetValue "Name", "Kelly"
    Key.SetValue "Age", 36
End Sub

Public Sub Teardown()
    Key.CloseKey
    DeleteSetting "VBCorLib_DeleteValue"
End Sub


Public Sub TestDelete()
    Key.DeleteValue "Name"
    Assert.IsEmpty Key.GetValue("Name")
    
    Key.DeleteValue "Age"
    Assert.IsEmpty Key.GetValue("Age")
End Sub

Public Sub TestDeleteMissing()
    On Error GoTo errTrap
    Key.DeleteValue "Missing", False
    Exit Sub
    
errTrap:
    Assert.Fail "No exception should be thrown."
End Sub

Public Sub TestDeleteMissingThrowable()
    On Error GoTo errTrap
    Key.DeleteValue "Missing"
    On Error GoTo 0
    Assert.Fail "An exception should be thrown."
    
errTrap:
    Dim ex As Exception
    Assert.IsTrue Catch(ex), "An exception should be caught."
    Assert.IsTrue TypeOf ex Is ArgumentException, "Wrong exception type caught."
End Sub
