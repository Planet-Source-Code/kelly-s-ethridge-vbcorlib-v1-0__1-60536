VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestRegistryKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestRegistryKey
'
Option Explicit

Private root As RegistryKey


Public Sub Setup()
    Set root = Registry.CurrentUser
    
    DeleteSubKeys
End Sub

Public Sub Teardown()
    DeleteSubKeys
End Sub

Public Sub TestOpenSubKey()
    Dim rk As RegistryKey
    
    Set rk = root.OpenSubKey("software\vb and vba program settings\")
    Assert.IsNotNothing rk
    Assert.Equals "HKEY_CURRENT_USER\software\vb and vba program settings", rk.Name
    Assert.IsTrue rk.SubKeyCount > 0, "Should be more than 0 sub keys."
    rk.CloseKey
End Sub

Public Sub TestCreateSubKeyFullPath()
    Dim rk As RegistryKey
    
    Set rk = Registry.CurrentUser.CreateSubKey("software\vb and vba program settings\VBCorLibTests_CreateSubKey")
    
    Assert.IsNotNothing rk
    Assert.Equals 0, rk.SubKeyCount
    Assert.Equals 0, rk.ValueCount
    Assert.Equals "HKEY_CURRENT_USER\software\vb and vba program settings\VBCorLibTests_CreateSubKey", rk.Name
    rk.CloseKey
End Sub

Public Sub TestCreateSubKeyWritable()
    Dim rk As RegistryKey
    
    Set rk = root.OpenSubKey("software\vb and vba program settings", True)
    Set rk = rk.CreateSubKey("VBCorLibTests_CreateSubKey")
    Assert.IsNotNothing rk
    Assert.Equals 0, rk.ValueCount
    Assert.Equals 0, rk.SubKeyCount
    Assert.Equals "HKEY_CURRENT_USER\software\vb and vba program settings\VBCorLibTests_CreateSubKey", rk.Name
    rk.CloseKey
End Sub

Public Sub TestCreateSubKeyNonWritable()
    Dim rk As RegistryKey
    
    Set rk = root.OpenSubKey("software\vb and vba program settings")
    
    On Error GoTo errTrap
    Set rk = rk.CreateSubKey("VBCorLibTests_CreateSubKey")
    On Error GoTo 0
    Assert.Fail "An exception should be thrown."
    
errTrap:
    Dim ex As Exception
    Assert.IsTrue Catch(ex), "An exception should be caught."
    Assert.IsTrue TypeOf ex Is UnauthorizedAccessException, "Wrone exception type caught."
End Sub

Public Sub TestSubKeyNotFound()
    Dim rk As RegistryKey
    
    Set rk = root.OpenSubKey("software\vb and vba program settings\notfound")
    Assert.IsNothing rk
End Sub

Public Sub TestDeleteSubKey()
    Dim rk As RegistryKey
    
    Set rk = root.OpenSubKey("software\vb and vba program settings", True)
    Assert.IsNotNothing rk.CreateSubKey("VBCorLibTests_DeleteSubKey")
    
    rk.DeleteSubKey "VBCorLibTests_DeleteSubKey"
    Assert.IsNothing rk.OpenSubKey("VBCorLibTests_DeleteSubKey")
    rk.CloseKey
End Sub

Public Sub TestDeleteMissingSubKeyThrowable()
    Dim rk As RegistryKey
    
    Set rk = root.OpenSubKey("software\vb and vba program settings", True)
    On Error GoTo errTrap
    rk.DeleteSubKey "missing"
    On Error GoTo 0
    Assert.Fail "An exception should be thrown."
    
errTrap:
    Dim ex As Exception
    Assert.IsTrue Catch(ex), "An exception should be caught."
    Assert.IsTrue TypeOf ex Is ArgumentException, "Wrong exception type caught."
End Sub

Public Sub TestDeleteSubKeyWithSubKeysThrowable()
    Dim rk As RegistryKey
    
    Set rk = root.CreateSubKey("software\vb and vba program settings\VBCorLibTests_GetSubKeyNames")
    rk.CreateSubKey "one"
    rk.CreateSubKey "two"
    rk.CreateSubKey "three"
    rk.CloseKey
    
    Set rk = root.OpenSubKey("software\vb and vba program settings", True)
    On Error GoTo errTrap
    rk.DeleteSubKey "VBCorLibTests_GetSubKeyNames"
    On Error GoTo 0
    Assert.Fail "An exception should be thrown."
        
errTrap:
    Dim ex As Exception
    Assert.IsTrue Catch(ex), "an exception should be caught."
    Assert.IsTrue TypeOf ex Is InvalidOperationException, "Wrong exception type caught."
End Sub

Public Sub TestGetSubKeyNames()
    Dim rk As RegistryKey
    Dim s() As String
    Dim i As Long
    
    Set rk = root.CreateSubKey("software\vb and vba program settings\VBCorLibTests_GetSubKeyNames")
    rk.CreateSubKey "one"
    rk.CreateSubKey "two"
    rk.CreateSubKey "three"
    
    s = rk.GetSubKeyNames
    rk.CloseKey
    
    Assert.Equals 3, cArray.GetLength(s), "Should contain 3 elements."
    Assert.IsTrue (s(0) <> s(1)) And (s(0) <> s(2)) And (s(1) <> s(2)), "Elements should be unique."
    For i = LBound(s) To UBound(s)
        Assert.IsIn s(i), "one", "two", "three"
    Next i
End Sub

Public Sub TestGetSubKeyNamesEmpty()
    Dim rk As RegistryKey
    Dim s() As String
    
    Set rk = root.CreateSubKey("software\vb and vba program settings\VBCorLibTests_GetSubKeyNames")
    s = rk.GetSubKeyNames
    rk.CloseKey
    
    Assert.Equals 0, cArray.GetLength(s), "Should be an empty array."
End Sub

Public Sub TestDeleteSubKeyTree()
    Dim rk As RegistryKey
    
    Set rk = root.CreateSubKey("software\vb and vba program settings\VBCorLibTests_DeleteTree")
    rk.CreateSubKey "one"
    rk.CreateSubKey "two"
    rk.CreateSubKey "three"
    rk.CloseKey
    
    Set rk = root.OpenSubKey("software\vb and vba program settings", True)
    rk.DeleteSubKeyTree "VBCorLibTests_DeleteTree"
    Assert.IsNothing rk.OpenSubKey("VBCorLibTests_DeleteTree")
End Sub

Public Sub TestCreateExistingSubKey()
    Dim rk As RegistryKey
    
    Set rk = root.CreateSubKey("software\vb and vba program settings\VBCorLibTests_Exists")
    rk.CloseKey
    
    Set rk = root.CreateSubKey("software\vb and vba program settings\VBCorLibTests_Exists")
    Assert.IsNotNothing rk
    rk.CloseKey
End Sub


'Public Sub TestRemote()
'    Dim rk As RegistryKey
'    Dim s() As String
'    Dim i As Long
'
'    Set rk = RegistryKey.OpenRemoteBaseKey(Users, "\\Soze")
'    Assert.IsNotNothing rk
'    Set rk = rk.OpenSubKey(".Default")
'    s = rk.GetSubKeyNames
'    For i = 0 To UBound(s)
'        SimplyVBUnit.Console.WriteLine s(i)
'    Next i
'    rk.CloseKey
'End Sub


Private Sub DeleteSubKeys()
    On Error Resume Next
    DeleteSetting "VBCorLibTests_CreateSubKey"
    DeleteSetting "VBCorLibTests_GetSubKeyNames"
    DeleteSetting "VBCorLibTests_DeleteTree"
    DeleteSetting "VBCorLibTests_Exists"
End Sub
