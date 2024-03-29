VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StringWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'    CopyRight (c) 2005 Kelly Ethridge
'
'    This file is part of VBCorLib.
'
'    VBCorLib is free software; you can redistribute it and/or modify
'    it under the terms of the GNU Library General Public License as published by
'    the Free Software Foundation; either version 2.1 of the License, or
'    (at your option) any later version.
'
'    VBCorLib is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Library General Public License for more details.
'
'    You should have received a copy of the GNU Library General Public License
'    along with Foobar; if not, write to the Free Software
'    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'    Module: StringWriter
'

''
' Provides a set of functions for writing to a string.
'
' @see Constructors
' @see TextWriter
'
Option Explicit
Implements cObject
Implements TextWriter

Private mSB As StringBuilder
Private mProvider As IFormatProvider
Private mEncoding As UnicodeEncoding
Private mNewLine As String
Private mIsClosed As Boolean



''
' Returns the encoding in which the output is written.
'
' @return The encoding of the output string.
'
Public Property Get Encoding() As Encoding
    If mEncoding Is Nothing Then Set mEncoding = New UnicodeEncoding
    Set Encoding = mEncoding
End Property

''
' Returns the format provider the writer uses for formatting values.
'
' @return The format provider used by the writer.
' @remarks The format provider used is the returned by the current culture.
'
Public Property Get FormatProvider() As IFormatProvider
    If mProvider Is Nothing Then Set mProvider = CultureInfo.CurrentCulture
    Set FormatProvider = mProvider
End Property

''
' Returns the string currently used to represent a new-line set of characters.
'
' @return The new-line set of characters.
' @remarks The default is CarriageReturn-LineFeed.
'
Public Property Get NewLine() As String
    NewLine = mNewLine
End Property

''
' Sets the string to be used as the new-line set of characters.
'
' @param RHS The new-line set of characters.
' @remarks The default is CarriageReturn-LineFeed.
'
Public Property Let NewLine(ByVal RHS As String)
    mNewLine = RHS
End Property

''
' Closes the writer
'
Public Sub CloseWriter()
    mIsClosed = True
End Sub

''
' Writes any buffered data to the underlying string.
'
' @remarks Since the data is directly written to the underlying string,
' this function does nothing. It is here to maintain interface consistency.
'
Public Sub Flush()
    ' does nothing
End Sub

''
' Returns the underlying stringbuilder being used to write to the string.
'
' @return The StringBuilder being used to write to the string.
'
Public Function GetStringBuilder() As StringBuilder
    VerifyBuilder
    Set GetStringBuilder = mSB
End Function

''
' Writes a value to the string.
'
' @param value The value to be written to the stream.
' @param args The arguments to be used in formatting the Value.
' @remarks The <i>args</i> are used when argument indexes are supplied
' in the value to be written. Indexes are indicate by enclosing the
' specific index in curly braces {}. The indexes start at 0.
'
Public Sub WriteValue(ByRef Value As Variant, ParamArray args() As Variant)
    Dim a() As Variant
    Helper.Swap4 ByVal ArrPtr(a), ByVal Helper.DerefEBP(16)
    InternalWriteValue Value, a
End Sub

''
' Writes a value to the string, followed by a new-line set of characters.
'
' @param Values The value to be written to the string
' @remarks The first argument in <i>values</i> is the output to the
' string. Any additional arguments are used as formatting information.
' An argument can be refered to through formatting indexes. A
' formatting index is an integer enclosed in curly braces {}. The
' indexes start at 0.
' <p>Numbers and Dates will be formatted to the culture that was
' passed in using Cor.NewStringWriter. If no culture was specified, then
' the current culture for this computer is used.</p>
'
Public Sub WriteLine(ParamArray values() As Variant)
    Dim a() As Variant
    Helper.Swap4 ByVal ArrPtr(a), ByVal Helper.DerefEBP(12)
    InternalWriteLine a
End Sub

''
' Returns a string representation of this object instance.
'
' @return String representing this instance.
Public Function ToString() As String
    VerifyBuilder
    ToString = mSB.ToString
End Function

''
' Returns a boolean indicating if the value and this object
' instance are the same instance.
'
' @param value The value to compare equality to.
' @return Boolean indicating equality.
Public Function Equals(ByRef Value As Variant) As Boolean
    If IsObject(Value) Then
        Equals = (Value Is Me)
    End If
End Function

''
' Returns a pseudo-unique number identifying this instance.
'
' @return Pseudo-unique number identifying this instance.
Public Function GetHashCode() As Long
    GetHashCode = ObjPtr(Me)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal sb As StringBuilder, ByVal provider As IFormatProvider)
    Set mSB = sb
    Set mProvider = provider
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub VerifyBuilder()
    If mSB Is Nothing Then Set mSB = New StringBuilder
End Sub

Private Sub VerifyIsOpen()
    If mIsClosed Then Throw Cor.NewObjectDisposedException("StringWriter", "Cannot write to a closed StringWriter.")
End Sub

Private Function GetString(ByRef Value As Variant) As String
    Dim nf As NumberFormatInfo
    Dim fm As IFormattable
    Dim o As cObject
    
    Select Case VarType(Value)
        Case vbLong, vbInteger, vbByte, vbDouble, vbSingle
            Set nf = FormatProvider.GetFormat("NumberFormatInfo")
            If nf Is Nothing Then
                GetString = Value
            Else
                GetString = nf.Format(Value)
            End If
        Case vbDate
            GetString = cDateTime.AscDateTime(Value).ToString("", FormatProvider)
        Case vbObject
            If TypeOf Value Is IFormattable Then
                Set fm = Value
                GetString = fm.ToString("", FormatProvider)
            ElseIf TypeOf Value Is cObject Then
                Set o = Value
                GetString = o.ToString
            Else
                Throw Cor.NewArgumentException("Value cannot be written to the StringWriter.", "Value")
            End If
        Case Else
            GetString = Value
    End Select
End Function

Private Sub InternalWrite(ByRef Value As String, ByRef args() As Variant, ByVal NewLine As Boolean)
    Dim s As String
    
    VerifyIsOpen
    VerifyBuilder
    If NewLine Then s = mNewLine
    s = Value & s
    
    mSB.InternalAppendFormat Nothing, s, args
End Sub

Private Sub InternalWriteLine(ByRef args() As Variant)
    Dim ub As Long
    Dim s As String
    
    ub = UBound(args)
    
    If ub >= 0 Then s = GetString(args(0))
    If ub > 0 Then
        Dim i As Long
        For i = 1 To ub
            Helper.MoveVariant args(i - 1), args(i)
        Next i
        ReDim Preserve args(UBound(args) - 1)
    End If
    
    InternalWrite s, args, True
End Sub

Private Sub InternalWriteValue(ByRef Value As Variant, ByRef args() As Variant)
    InternalWrite GetString(Value), args, False
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    mNewLine = Environment.NewLine
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   cObject Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function cObject_Equals(Value As Variant) As Boolean
    cObject_Equals = Equals(Value)
End Function

Private Function cObject_GetHashcode() As Long
    cObject_GetHashcode = GetHashCode
End Function

Private Function cObject_ToString() As String
    cObject_ToString = ToString
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   TextWriter Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub TextWriter_CloseWriter()
    CloseWriter
End Sub

Private Property Get TextWriter_Encoding() As Encoding
    Set TextWriter_Encoding = Encoding
End Property

Private Function TextWriter_Equals(Value As Variant) As Boolean
    TextWriter_Equals = Equals(Value)
End Function

Private Sub TextWriter_Flush()
    Flush
End Sub

Private Property Get TextWriter_FormatProvider() As IFormatProvider
    Set TextWriter_FormatProvider = FormatProvider
End Property

Private Function TextWriter_GetHashCode() As Long
    TextWriter_GetHashCode = GetHashCode
End Function

Private Property Let TextWriter_NewLine(ByVal RHS As String)
    NewLine = RHS
End Property

Private Property Get TextWriter_NewLine() As String
    TextWriter_NewLine = NewLine
End Property

Private Function TextWriter_ToString() As String
    TextWriter_ToString = ToString
End Function

Private Sub TextWriter_WriteLine(ParamArray Value() As Variant)
    Dim a() As Variant
    Helper.Swap4 ByVal ArrPtr(a), ByVal Helper.DerefEBP(12)
    InternalWriteLine a
End Sub

Private Sub TextWriter_WriteValue(ByRef Value As Variant, ParamArray args() As Variant)
    Dim a() As Variant
    Helper.Swap4 ByVal ArrPtr(a), ByVal Helper.DerefEBP(16)
    InternalWriteValue Value, a
End Sub
