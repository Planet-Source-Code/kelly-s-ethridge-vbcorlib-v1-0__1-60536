VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestDirectoryInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestDirectoryInfo
'
Option Explicit

Private info As DirectoryInfo

Public Sub TestFixtureSetup()
    SimplyVBUnit.Console.WriteLine ""
    SimplyVBUnit.Console.WriteLine "***** DIRECTORYINFO *****"
End Sub

Public Sub Setup()
    Set info = NewDirectoryInfo(App.Path)
End Sub

Public Sub TestAttributes()
    Assert.Equals 16, info.Attributes And FileAttributes.DirectoryAttr
End Sub

Public Sub TestCreationTime()
    SimplyVBUnit.Console.WriteLine ""
    SimplyVBUnit.Console.WriteLine "--- VISUAL INSPECTION ---"
    SimplyVBUnit.Console.WriteLine "CreationTime: " & info.CreationTime.ToString
End Sub

Public Sub TestCreationTimeLocalVsUTC()
    Dim loc As cDateTime
    Dim utc As cDateTime
    Dim offset As TimeSpan
    
    Set loc = info.CreationTime
    Set utc = info.CreationTimeUtc
    Set offset = TimeZone.CurrentTimeZone.GetUtcOffset(loc)
    
    SimplyVBUnit.Console.WriteLine cString.Format("Local: {0}", loc)
    SimplyVBUnit.Console.WriteLine cString.Format("UTC: {0}", utc)
    SimplyVBUnit.Console.WriteLine cString.Format("Offset: {0}", offset)
    
    Assert.Equals offset.TotalHours, loc.Subtract(utc).TotalHours
End Sub

Public Sub TestName()
    Assert.Equals "Tests", info.Name
End Sub

Public Sub TestParentName()
    Dim p As DirectoryInfo
    
    Set p = info.Parent
    Assert.IsNotNothing p
    Assert.Equals "VBCorLibGroup", p.Name
End Sub

Public Sub TestRoot()
    Dim r As DirectoryInfo
    
    Set r = info.Root
    Assert.IsNotNothing r, "No root information in path."
    Assert.Equals "c:\", LCase$(r.Name)
End Sub

Public Sub TestGetDirectories()
    Dim di() As DirectoryInfo
    Dim i As Long
    
    di = info.GetDirectories
    Assert.IsFalse carray.IsNull(di), "Should not be a null array."
    
    SimplyVBUnit.Console.WriteLine ""
    SimplyVBUnit.Console.WriteLine "--- VISUAL INSPECTION ---"
    SimplyVBUnit.Console.WriteLine "GetDirectories"
    SimplyVBUnit.Console.WriteLine "Count: " & carray.GetLength(di)
    For i = 0 To UBound(di)
        SimplyVBUnit.Console.WriteLine di(i).Name
    Next i
End Sub

Public Sub TestNoParent()
    Set info = cor.NewDirectoryInfo("c:\")
    Assert.IsNothing info.Parent
    
    Set info = cor.NewDirectoryInfo("\\server\share")
    Assert.IsNothing info.Parent
End Sub

Public Sub TestGetDirectoriesNoMatch()
    Dim di() As DirectoryInfo
    
    di = info.GetDirectories("abc123")
    Assert.IsFalse carray.IsNull(di)
    Assert.Equals 1, carray.GetLength(di)
    Assert.Equals Path.GetPathRoot(Path.GetFullPath("abc123")), di(0).ToString
End Sub

Public Sub TestGetFilesNoMatch()
    Dim fi() As FileInfo
    
    fi = info.GetFiles("abc123")
    Assert.IsFalse carray.IsNull(fi)
    Assert.Equals 0, carray.GetLength(fi)
End Sub
