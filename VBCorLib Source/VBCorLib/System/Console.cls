VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Console"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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
'    Module: Console
'

''
' Provides a mechanism for handling input and output for either a screen or Stream.
'
' @remarks Using the <b>Console</b> class allows input and output to a console window.
' By default, a VB6 application will create a new console window for the <b>Console</b>
' class to use. Most times it is desired to run an application in an already open
' console window. If a VB6 application is run from within an existing console window,
' a new window will be created again. To have the VB6 application to use the console
' window from which is it run, the EXE needs to be converted to a console application.
' This can be accomplished by hand using the <i>Link</i> program that is included with
' Visual Basic 6. In a console window use the following command:
' <p>link /edit /subsystem:console <EXE></p>
' <p>Replace <EXE> with the path and executable to be flagged as a console application.
' Be sure the <i>Link</i> and application are both in paths the console window can find.
' Once the command has been executed, the VB6 application will run inside the console
' window from which it was run. If no window existed, then a new one will be created.</p>
' <p>The Console can be used to display text on the screen, and retrieve
' user input from a keyboard. The input and output can be redirected to a new
' source to allow for reading and writing from a specified Stream.</p>
' <p><b>WARNING</b>: Do not close the console through with <b>Close Window</b> button while
' working in the IDE. This will cause the IDE to crash and all data will be lost.</p>
'
Option Explicit
Private Declare Function AllocConsole Lib "kernel32.dll" () As Long
Private Declare Function FreeConsole Lib "kernel32.dll" () As Long
Private Declare Function GetConsoleOutputCP Lib "kernel32.dll" () As Long
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Private Declare Function SetConsoleCtrlHandler Lib "kernel32.dll" (ByVal HandlerRoutine As Long, ByVal Add As Long) As Long

Private Const STD_OUTPUT_HANDLE As Long = -11&
Private Const STD_INPUT_HANDLE As Long = -10&
Private Const STD_ERROR_HANDLE As Long = -12&

''
' A list of break types returned by the console.
'
' @param ControlC The CTRL+C key combination was pressed.
' @param ControlBreak The CTRL+Break key combination was pressed.
' @param CloseWindow The window is being closed.
' @param LogOffUser The user is logging off the system.
' @param ShutDownSystem The system is being shut down.
'
Public Enum ConsoleBreakType
    ControlC = 0
    ControlBreak = 1
    CloseWindow = 2
    LogOffUser = 5
    ShutDownSystem = 6
End Enum

Private mOwnsConsole As Boolean
Private mOut As TextWriter
Private Min As TextReader
Private mError As TextWriter



''
' Returns the cause of the break event in the console
'
' @return The type of break that caused the interuption.
'
Public Property Get BreakType() As ConsoleBreakType
    BreakType = modConsoleCallbacks.BreakType
End Property

''
' Returns if a break has been caused.
'
' @return If a break has been caused in the console.
' @remarks This property will be set to true when one of
' several actions occurs. The actions are defined by the
' enumeration <b>ConsoleBreakType</b>. No DoEvents need
' to be called for this property to be set. Simply check
' this property when necessary to respond to any of the events.
'
Public Property Get Break() As Boolean
    Break = modConsoleCallbacks.Break
End Property

''
' Sets the break flag.
'
' @param RHS The break flag value.
' @remarks Set Break to false once it has been handled so the
' next console break can set it to true and be detected.
'
Public Property Let Break(ByVal RHS As Boolean)
    modConsoleCallbacks.Break = RHS
End Property

''
' Returns the Error output currently being used by the Console.
'
' @return An object that implements the TextWriter interface.
' @remarks The default Error writer writes to the Console error output.
'
Public Property Get Error() As TextWriter
    If mError Is Nothing Then Set mError = Cor.NewStreamWriter(OpenStandardError, Encoding.GetEncoding(GetConsoleOutputCP))
    Set Error = mError
End Property

''
' Returns the output writer currently being used by the Console.
'
' @return Returns an object that implements the TextWriter interface.
' @remarks The default output for the Console is to a Console window displayed
' on the screen. The output can be redirected to a new Stream using SetOut.
'
Public Property Get Out() As TextWriter
    If mOut Is Nothing Then
        Dim writer As StreamWriter
        Set writer = Cor.NewStreamWriter(OpenStandardOutput, Encoding.GetEncoding(GetConsoleOutputCP))
        writer.AutoFlush = True
        Set mOut = writer
    End If
    Set Out = mOut
End Property

''
' Returns the input source currently used by the Console.
'
' @return Returns an object that implements the TextReader interface
' @remarks The default source of the input is the Keyboard. This can
' be changed by using the SetIn function.
'
Public Property Get In_() As TextReader
    If Min Is Nothing Then Set Min = Cor.NewStreamReader(OpenStandardInput, Encoding.GetEncoding(GetConsoleOutputCP))
    Set In_ = Min
End Property

''
' Reads a character from the current input source.
'
' @return The next character read from the input source, or -1 if no more characters exists.
' @remarks The Read method does not return until the Return key is pressed.
' Once the function returns, it can be called repeatedly until all key presses
' upto the return key are retrieved. The return key is also returned
' as a carriage-return (13) followed by a line-feed (10).
' @include "..\Includes\ConsoleRead.txt"
Public Function Read() As Long
    Read = In_.Read
End Function

''
' Reads all of the characters from the current input source until
' a NewLine break is reached.
'
' @return Returns all of the available characters upto a NewLine break,
' not including the NewLine characters.
' @remarks The ReadLine function will block and not return until a NewLine
' characters is pressed. This is usually the Return key.
' @include "..\Includes\ConsoleReadLine.txt"
Public Function ReadLine() As String
    ReadLine = In_.ReadLine
End Function

''
' Sets the Console output to a new output TextWriter object.
'
' @param NewOut The TextWriter object that console output will be written to.
'
Public Sub SetOut(ByVal NewOut As TextWriter)
    If NewOut Is Nothing Then _
        Throw Cor.NewArgumentNullException("Cannot set output to Nothing.", "NewOut")
    
    Set mOut = NewOut
End Sub

''
' Sets the Console input source to a new TextReader object.
'
' @param NewIn The TextReader object the console retrieves input from.
'
Public Sub SetIn(ByVal NewIn As TextReader)
    If NewIn Is Nothing Then _
        Throw Cor.NewArgumentNullException("Cannot set input to Nothing.", "NewIn")
        
    Set Min = NewIn
End Sub

''
' Sets the Console error output stream to the supplied TextWriter object.
'
' @param NewError The TextWriter object the console error data is written to.
'
Public Sub SetError(ByVal NewError As TextWriter)
    If NewError Is Nothing Then _
        Throw Cor.NewArgumentNullException("Cannot set error output to Nothing.", "NewError")
        
    Set mError = NewError
End Sub

''
' Returns the standard output stream.
'
' @return The standard output stream to write user data to.
' @remarks The standard output is the screen console.
'
Public Function OpenStandardOutput() As Stream
    Set OpenStandardOutput = GetConsoleStream(STD_OUTPUT_HANDLE)
End Function

''
' Returns the standard input stream.
'
' @return The standard input stream object to retrieve user input from.
' @remarks The standard input is from the keyboard console.
'
Public Function OpenStandardInput() As Stream
    Set OpenStandardInput = GetConsoleStream(STD_INPUT_HANDLE)
End Function

''
' Returns the standard error stream.
'
' @return The standard error output stream that error data is written to.
' @remarks The standard error output is to the screen console.
'
Public Function OpenStandardError() As Stream
    Set OpenStandardError = GetConsoleStream(STD_ERROR_HANDLE)
End Function

''
' Writes a value to the current output stream.
'
' @param values A set of zero or more parameters used to create a formatted string output.
'
' @remarks The first parameter is the string that will be written to
' the current output stream. All additional parameters will be used as
' formatting information in the first parameter. If no parameters are
' provided, then an empty string is written with a NewLine.
' @include "..\Includes\ConsoleWriteLine.txt"
Public Sub WriteLine(ParamArray values() As Variant)
    Dim a() As Variant
    Helper.Swap4 ByVal ArrPtr(a), ByVal Helper.DerefEBP(12)
    InternalWriteLine a
End Sub

''
' Writes a string to the Console output stream without a NewLine break.
'
' @param value The value to write the string version of to the output stream.
' @param Args Values to be used in a formatted output.
' @remarks Values are converted to their string representation for text output.
' A formatted output can be created similar to <b>cString.Format</b>.
' @include "..\Includes\ConsoleWriteValue.txt"
Public Sub WriteValue(ByRef Value As Variant, ParamArray args() As Variant)
    Dim a() As Variant
    Helper.Swap4 ByVal ArrPtr(a), ByVal Helper.DerefEBP(16)
    InternalWriteValue Value, a
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetConsoleStream(ByVal StreamType As Long) As Stream
    Dim ret As New ConsoleStream
    ret.Init StreamType
    Set GetConsoleStream = ret
End Function

Private Sub InternalWrite(ByVal Value As String, ByRef args() As Variant, ByVal NewLine As Boolean)
    If Not cArray.IsNull(args) Then
        If UBound(args) >= 0 Then Value = cString.FormatArray(Value, args)
    End If
    If NewLine Then Value = Value & Out.NewLine
    Out.WriteValue Value
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
    Else
        Erase args
    End If
    
    InternalWrite s, args, True
End Sub

Private Sub InternalWriteValue(ByRef Value As Variant, ByRef args() As Variant)
    InternalWrite GetString(Value), args, False
End Sub

Private Function GetString(ByRef Value As Variant) As String
    Dim nf As NumberFormatInfo
    Dim fm As IFormattable
    Dim o As cObject
    
    Select Case VarType(Value)
        Case vbLong, vbInteger, vbByte, vbDouble, vbSingle
            Set nf = Out.FormatProvider.GetFormat("NumberFormatInfo")
            If nf Is Nothing Then
                GetString = Value
            Else
                GetString = nf.Format(Value)
            End If
        Case vbDate
            GetString = cDateTime.AscDateTime(Value).ToString("", Out.FormatProvider)
        Case vbObject
            If TypeOf Value Is IFormattable Then
                Set fm = Value
                GetString = fm.ToString("", Out.FormatProvider)
            ElseIf TypeOf Value Is cObject Then
                Set o = Value
                GetString = o.ToString
            Else
                Throw Cor.NewArgumentException("Value cannot be written to the StreamWriter.", "Value")
            End If
        Case Else
            GetString = Value
    End Select
End Function

Private Sub Init()
    mOwnsConsole = (AllocConsole <> BOOL_FALSE)
    Sleep 10    ' give it time to be created
    SetConsoleCtrlHandler AddressOf ControlBreakHandler, True
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Init
End Sub

Private Sub Class_Terminate()
    SetConsoleCtrlHandler AddressOf ControlBreakHandler, False
    If mOwnsConsole Then FreeConsole
End Sub
