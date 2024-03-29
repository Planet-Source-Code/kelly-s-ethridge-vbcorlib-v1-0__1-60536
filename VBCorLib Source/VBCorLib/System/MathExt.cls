VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MathExt"
Attribute VB_GlobalNameSpace = True
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
'    Module: Math2
'

''
' Provides extended math functionality.
'
' @remarks The functions in this class are global. To access them simply call
' the functions directly as any normal function. The MathExt class does not need
' to be referenced.
'
Option Explicit

Private Const ATN_1     As Double = 0.785398163397448
Private Const LOG_10    As Double = 2.30258509299405
Private Const NEG_ZERO  As Currency = -922337203685477.5807@ - 0.0001@   ' VB won't accept -922337203685477.5808@ as a constant
Private Const C_PI      As Double = 3.14159265358979

Private mNegativeZero As Double


''
' Returns the next highest whole integer value.
'
' @param Value The number to round to the next integer.
' @return The next integer.
' @remarks Negative values will round towards zero. If the
' value is already a whole integer value then that value will be returned.
'
Public Function Ceiling(ByVal Value As Double) As Double
    If Value < 0# Then
        Ceiling = Fix(Value)
    Else
        Ceiling = -Int(-Value)
    End If
End Function

''
' Returns the next lowest whole integer value.
'
' @param Value the number to be rounded downward.
' @return The next lowest whole integer value.
' @remarks Negative values will be rounded away from zero. -2.1 will be
' rounded down to -3.
'
Public Function Floor(ByVal Value As Double) As Double
    Floor = Int(Value)
End Function

''
' Shifts the bits of an integer left.
'
' @param Value The value to be shifted left.
' @param Count The number of bits to shift.
' @return The newly shifted value.
'
Public Function LShift(ByVal Value As Long, ByVal Count As Long) As Long
    LShift = Helper.ShiftLeft(Value, Count)
End Function

''
' Shifts the bits of an integer right.
'
' @param Value The value to be shifted right.
' @param Count The number of bits to shift.
' @return The newly shifted value.
'
Public Function RShift(ByVal Value As Long, ByVal Count As Long) As Long
    RShift = Helper.ShiftRight(Value, Count)
End Function

''
' Divides two integers, placing the remainder in a supplied variable.
'
' @param a The dividend.
' @param b The divosor.
' @param Remainder The variable to place the remainder of the division.
' @return The quotient of the division.
'
Public Function DivRem(ByVal a As Long, ByVal b As Long, ByRef Remainder As Long) As Long
    Remainder = a Mod b
    DivRem = a \ b
End Function

''
' Represents the natural logarithmic base, specified by the constant, <b><i>e</i></b>.
'
' @return <b><i>e</i></b> (Approx 2.71828182845905)
'
Public Property Get E() As Double
    E = 2.71828182845905
End Property

''
' Represents the ratio of the circumference of a circle to its diameter.
'
' @return Pi (approx 3.14159265358979)
'
Public Property Get PI() As Double
    PI = C_PI
End Property

''
' Returns the maximum of two values.
'
' @param x Value to test as potential max value.
' @param y Value to test as potential max value.
' @return The max value.
' @remarks No type checking is performed to determine the maximum of two
' values given. A string compared to an integer is a legal pair of values
' to test. The only type checked is for an <b>IComparable</b> object. If
' <b>x</b> is an <b>IComparable</b>, then the object is used to test for
' the maximum value.
' <p>If both values are equavalent then <b>x</b> is returned.
'
Public Function Max(ByRef x As Variant, ByRef y As Variant) As Variant
    On Error GoTo errTrap
    
    If Not IsObject(x) Then
        If x < y Then
            Max = y
        Else
            Max = x
        End If
    ElseIf TypeOf x Is IComparable Then
        Dim c As IComparable
        Set c = x
        If c.CompareTo(y) < 0 Then
            Set Max = y
        Else
            Set Max = x
        End If
    Else
        On Error GoTo 0
        Throw Cor.NewArgumentException("Values could not be compared.")
    End If
    Exit Function
    
errTrap:
    Throw Cor.NewArgumentException("Values could not be compared.")
End Function

''
' Returns the minimum of two values.
'
' @param x Value to test as potential min value.
' @param y Value to test as potential min value.
' @return The min value.
' @remarks No type checking is performed to determine the minimum of two
' values given. A string compared to an integer is a legal pair of values
' to test. The only type checked is for an <b>IComparable</b> object. If
' <b>x</b> is an <b>IComparable</b>, then the object is used to test for
' the minimum value.
' <p>If both values are equavalent then <b>x</b> is returned.
'
Public Function Min(ByRef x As Variant, ByRef y As Variant) As Variant
    On Error GoTo errTrap
    
    If Not IsObject(x) Then
        If x > y Then
            Min = y
        Else
            Min = x
        End If
    ElseIf TypeOf x Is IComparable Then
        Dim c As IComparable
        Set c = x
        If c.CompareTo(y) > 0 Then
            Set Min = y
        Else
            Set Min = x
        End If
    Else
        On Error GoTo 0
        Throw Cor.NewArgumentException("Values could not be compared.")
    End If
    Exit Function
    
errTrap:
    Throw Cor.NewArgumentException("Values could not be compared.")
End Function

''
' Returns the angle from a Cosine.
'
' @param d A number representing a Cosine.
' @return The angle in radians.
'
Public Function ACos(ByVal d As Double) As Double
    ACos = Atn(-d / Sqr(-d * d + 1)) + 2 * ATN_1
End Function

''
' Returns the angle from a Sine.
'
' @param d A number representing a Sine.
' @return The angle in radians.
'
Public Function ASin(ByVal d As Double) As Double
    ASin = Atn(d / Sqr(-d * d + 1))
End Function

''
' Returns the hyperbolic cosine of the specified angle.
'
' @param d An angle in radians.
' @return The hyberbolic cosine.
'
Public Function Cosh(ByVal d As Double) As Double
    Cosh = (Exp(d) + Exp(-d)) / 2
End Function

''
' Returns the hyberbolic sine of the specified angle.
'
' @param d An angle in radians.
' @return The hyberbolic sine.
'
Public Function Sinh(ByVal d As Double) As Double
    Sinh = (Exp(d) - Exp(-d)) / 2
End Function

''
' Returns the hyberbolic tangent of the specified angle.
'
' @param d An angle in radians.
' @return The hyperbolic tangent.
'
Public Function Tanh(ByVal d As Double) As Double
    Dim posExp As Double
    Dim negExp As Double
    
    posExp = Exp(d)
    negExp = Exp(-d)
    
    Tanh = (posExp - negExp) / (posExp + negExp)
End Function

''
' Returns a value in the specified base.
'
' @param d The number to calculate the log of.
' @param NewBase The base of the log value.
' @return The log value.
'
Public Function LogBase(ByVal d As Double, ByVal NewBase As Double) As Double
    LogBase = Log(d) / Log(NewBase)
End Function

''
' Returns a log value in a base 10 log.
'
' @param d The number to calculate the log of.
' @return A base 10 log value.
'
Public Function Log10(ByVal d As Double) As Double
    Log10 = Log(d) / LOG_10
End Function

''
' Returns the remainder resulting from the division of a specified number by another specified number.
'
' @param x A dividend.
' @param y A devisor.
' @return The remainder of the division.
'
Public Function IEEERemainder(ByVal x As Double, ByVal y As Double) As Double
    Dim ret As Double
    ret = (x - (y * Round(x / y)))
    If ret = 0# And x < 0# Then ret = mNegativeZero
    IEEERemainder = ret
End Function

''
' Converts radians to degrees.
'
' @param Radians An angle in radians.
' @return An angle in degrees.
'
Public Function CDeg(ByVal Radians As Double) As Double
    CDeg = 180# * Radians / C_PI
End Function

''
' Converts degrees to radians.
'
' @param Degrees An angle in degrees.
' @return An angle in radians.
'
Public Function CRad(ByVal Degrees As Double) As Double
    CRad = C_PI * Degrees / 180#
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    mNegativeZero = AsDouble(NEG_ZERO)
End Sub
