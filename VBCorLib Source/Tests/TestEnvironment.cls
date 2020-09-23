VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestEnvironment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestEnvironment
'
Option Explicit

' These don't really test the Environment class since the results
' will vary from machine to machine. Instead they simply display
' the results for the current machine for visual inspection.
Public Sub TestFixtureSetup()
    SimplyVBUnit.Console.WriteLine ""
    SimplyVBUnit.Console.WriteLine "***** ENVIRONMENT CLASS *****"
End Sub

Public Sub TestCommandLine()
    SimplyVBUnit.Console.WriteLine ""
    SimplyVBUnit.Console.WriteLine "--- VISUAL INSPECTION ---"
    SimplyVBUnit.Console.WriteLine "Command Line: " & environment.CommandLine
End Sub

Public Sub TestMachineName()
    SimplyVBUnit.Console.WriteLine ""
    SimplyVBUnit.Console.WriteLine "--- VISUAL INSPECTION ---"
    SimplyVBUnit.Console.WriteLine "Machine Name: " & environment.MachineName
End Sub

Public Sub TestOSVersion()
    Dim os As OperatingSystem
    
    Set os = environment.OSVersion
    SimplyVBUnit.Console.WriteLine ""
    SimplyVBUnit.Console.WriteLine "--- VISUAL INSPECTION ---"
    SimplyVBUnit.Console.WriteLine "OS: " & os.ToString
End Sub

Public Sub TestSystemDirectory()
    SimplyVBUnit.Console.WriteLine ""
    SimplyVBUnit.Console.WriteLine "--- VISUAL INSPECTION ---"
    SimplyVBUnit.Console.WriteLine "System Directory: " & environment.SystemDirectory
End Sub

Public Sub TestTickCount()
    SimplyVBUnit.Console.WriteLine ""
    SimplyVBUnit.Console.WriteLine "--- VISUAL INSPECTION ---"
    SimplyVBUnit.Console.WriteLine "TickCount: " & environment.TickCount
End Sub

Public Sub TestUserName()
    SimplyVBUnit.Console.WriteLine ""
    SimplyVBUnit.Console.WriteLine "--- VISUAL INSPECTION ---"
    SimplyVBUnit.Console.WriteLine "UserName: " & environment.userName
End Sub

Public Sub TestUserDomainName()
    SimplyVBUnit.Console.WriteLine ""
    SimplyVBUnit.Console.WriteLine "--- VISUAL INSPECTION ---"
    SimplyVBUnit.Console.WriteLine "UserDomainName: " & environment.UserDomainName
End Sub

Public Sub TestUserInteractive()
    SimplyVBUnit.Console.WriteLine ""
    SimplyVBUnit.Console.WriteLine "--- VISUAL INSPECTION ---"
    SimplyVBUnit.Console.WriteLine "UserInteractive: " & environment.UserInteractive
End Sub

Public Sub TestWorkingSet()
    SimplyVBUnit.Console.WriteLine ""
    SimplyVBUnit.Console.WriteLine "--- VISUAL INSPECTION ---"
    SimplyVBUnit.Console.WriteLine "Returns 0 for 95, 98 and ME installations."
    SimplyVBUnit.Console.WriteLine "WorkingSet (Memory Usage):" & environment.WorkingSet
End Sub

Public Sub TestExpandEnvironmentVariables()
    SimplyVBUnit.Console.WriteLine ""
    SimplyVBUnit.Console.WriteLine "--- VISUAL INSPECTION ---"
    SimplyVBUnit.Console.WriteLine "--- ExpandingEnvironmentVariables ---"
    SimplyVBUnit.Console.WriteLine "'My OS is %OS%' becomes: " & environment.ExpandEnvironmentVariables("My OS is %OS%")
End Sub

Public Sub TestGetCommandLineArgs()
    Dim s() As String
    Dim i As Long
    
    s = environment.GetCommandLineArgs
    
    SimplyVBUnit.Console.WriteLine ""
    SimplyVBUnit.Console.WriteLine "--- VISUAL INSPECTION ---"
    SimplyVBUnit.Console.WriteLine "--- GetCommandLingArgs ---"
    SimplyVBUnit.Console.WriteLine "Should have no empty args unless quoted."
    SimplyVBUnit.Console.WriteLine "Count: " & carray.GetLength(s)
    SimplyVBUnit.Console.WriteLine "Program: " & s(0)
    For i = 1 To UBound(s)
        SimplyVBUnit.Console.WriteLine "Arg(" & i & ") = " & s(i)
    Next i
End Sub

Public Sub TestGetEnvironmentVariable()
    SimplyVBUnit.Console.WriteLine ""
    SimplyVBUnit.Console.WriteLine "--- VISUAL INSPECTION ---"
    SimplyVBUnit.Console.WriteLine "GetEnvironmentVariable (os): " & environment.GetEnvironmentVariable("os")
End Sub

Public Sub TestGetEnvironmentVariables()
    Dim ht As Hashtable
    Dim i As Long
    
    SimplyVBUnit.Console.WriteLine ""
    SimplyVBUnit.Console.WriteLine "--- VISUAL INSPECTION ---"
    SimplyVBUnit.Console.WriteLine "--- TestGetEnvironmentVariables ---"
    
    Set ht = environment.GetEnvironmentVariables
    SimplyVBUnit.Console.WriteLine "Count: " & ht.Count
    
    Dim de As DictionaryEntry
    For Each de In ht
        i = i + 1
        SimplyVBUnit.Console.WriteLine "Variable(" & i & "): " & de.Key & " = " & de.value
    Next de
End Sub

Public Sub TestGetFolderPath()
    SimplyVBUnit.Console.WriteLine ""
    SimplyVBUnit.Console.WriteLine "--- VISUAL INSPECTION ---"
    SimplyVBUnit.Console.WriteLine "--- TestGetFolderPath ---"
    SimplyVBUnit.Console.WriteLine "My Documents: " & environment.GetFolderPath(Personal)
    SimplyVBUnit.Console.WriteLine "Desktop: " & environment.GetFolderPath(DesktopDirectory)
    SimplyVBUnit.Console.WriteLine "Windows: " & environment.GetFolderPath(Windows)
    SimplyVBUnit.Console.WriteLine "System: " & environment.GetFolderPath(SystemDir)
End Sub

Public Sub TestGetLogicalDrives()
    Dim s() As String
    Dim i As Long
    
    s = environment.GetLogicalDrives
    SimplyVBUnit.Console.WriteLine ""
    SimplyVBUnit.Console.WriteLine "--- VISUAL INSPECTION ---"
    SimplyVBUnit.Console.WriteLine "--- GetLogicalDrives ---"
    For i = 0 To UBound(s)
        SimplyVBUnit.Console.WriteLine "Drive - " & s(i)
    Next i
End Sub
