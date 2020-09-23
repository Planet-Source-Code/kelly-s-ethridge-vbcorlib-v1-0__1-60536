VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestRegistryRootKeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestRegistryRootKeys
'
Option Explicit

Public Sub TestClassesRoot()
    Assert.Equals "HKEY_CLASSES_ROOT", Registry.ClassesRoot.Name
End Sub

Public Sub TestLocalMachine()
    Assert.Equals "HKEY_LOCAL_MACHINE", Registry.LocalMachine.Name
End Sub

Public Sub TestDynDataOn9x()
    If Environment.OSVersion.Platform = Win32NT Then Exit Sub
    Assert.Equals "HKEY_DYN_DATA", Registry.DynData.Name
End Sub

Public Sub TestDynDataOnNT()
    If Not Environment.OSVersion.Platform = Win32NT Then Exit Sub
    
    On Error GoTo errTrap
    Assert.Equals "HKEY_DYN_DATA", Registry.DynData.Name
    On Error GoTo 0
    Assert.Fail "An exception should have been thrown."
    
errTrap:
    Dim ex As Exception
    Assert.IsTrue Catch(ex), "An exception should have been caught."
    Assert.IsTrue TypeOf ex Is IOException, "Wrong exception type caught."
End Sub

Public Sub TestSubKeyCount()
    Assert.IsTrue Registry.CurrentUser.SubKeyCount > 0, "Should be more than 0 sub keys."
End Sub
