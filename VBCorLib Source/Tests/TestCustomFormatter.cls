VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestCustomFormatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestCustomFormatter
'
Option Explicit
Implements IFormatProvider
Implements ICustomFormatter


Public Sub TestFormat()
    Dim s As String
    
    s = cString.Formatex(Me, "Name = {0}, Age = {1}", "Kelly", 36)
    Assert.Equals "Name = ylleK, Age = 36", s
End Sub



' format a value when requested, or return nothing if we didn't format it.
Private Function ICustomFormatter_Format(ByVal fmt As String, arg As Variant, ByVal provider As VBCorLib.IFormatProvider) As String
    ' we only format strings.
    If VarType(arg) = vbString Then
        ' we simply reverse a string value.
        ICustomFormatter_Format = StrReverse(arg)
    End If
End Function

' we only support ICustomFormatter types with this class.
Private Function IFormatProvider_GetFormat(ByVal FormatType As String) As Object
    If LCase$(FormatType) = "icustomformatter" Then Set IFormatProvider_GetFormat = Me
End Function
