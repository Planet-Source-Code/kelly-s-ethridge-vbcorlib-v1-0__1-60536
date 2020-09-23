VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestResourceReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestResourceReader
'
Option Explicit

Private mReader As ResourceReader


Public Sub Setup()
    Set mReader = Cor.NewResourceReader(Path.Combine(App.Path, "VBCorLibTests.res"))
End Sub

Public Sub Teardown()
    mReader.CloseReader
End Sub

Public Sub TestEnumeration()
    Dim de As DictionaryEntry
    Dim ws As Win32Resource
    Dim b() As Byte
    Dim i As Long
    
    For Each de In mReader
        Assert.IsTrue TypeOf de.value Is Win32Resource, "Wrong value type in DictionaryEntry."
        If de.Key = 101 Then
            Set ws = de.value
            Select Case ws.LanguageID
                Case 1033
                    If ws.ResourceType = StringResource Then Assert.Equals "hello", ws.value
                    If ws.ResourceType = UserDefined Then
                        b = ws.value
                        Assert.Equals 10, carray.GetLength(b)
                        Assert.Equals "CUSTOM", ws.ResourceTypeName
                        
                        For i = 0 To 9
                            Assert.Equals i + 48, b(i)
                        Next i
                    End If
                Case 1096
                    Assert.Equals 6, ws.ResourceType
                    Assert.Equals "aloha", ws.value
                    
                Case Else
                    Assert.Fail "Found wrong language ID <" & de.Key & ">"
            End Select
        ElseIf de.Key = 102 Then
            Set ws = de.value
            Assert.Equals 1033, ws.LanguageID
            Assert.Equals "CUSTOM2", ws.ResourceTypeName
            
            b = ws.value
            Assert.Equals 100, carray.GetLength(b)
            
            For i = 1 To 100
                Assert.Equals (i Mod 10) + 48, b(i - 1)
            Next i
        Else
            Assert.Fail "Wrong resource ID"
        End If
    Next de
End Sub


