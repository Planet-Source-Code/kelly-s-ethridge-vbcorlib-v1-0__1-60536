VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NumberFormatInfoStatic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'    CopyRight (c) 2004 Kelly Ethridge
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
'    Module: NumberFormatInfoStatic
'

''
' Provides methods to retrieve the various types of NumberFormatInfo objects.
'
' @see NumberFormatInfo
'
Option Explicit

''
' Returns the NumberFormatInfo object of the threads current culture.
'
Public Property Get CurrentInfo() As NumberFormatInfo
    Set CurrentInfo = CultureInfo.CurrentCulture.NumberFormat
End Property

''
' Returns the NumberFormatInfo object of the variant culture type. This
' is the same no matter what culture is current.
'
Public Property Get InvariantInfo() As NumberFormatInfo
    Set InvariantInfo = CultureInfo.InvariantCulture.NumberFormat
End Property

''
' Returns the NumberFormatInfo object from a specific provider.
'
' @param FormatProvider The provider of the NumberFormatInfo object.
' @return A NumberFormatInfo object from the provider.
' @remarks If the provider does not provide a NumberFormatInfo object, then
' the NumberFormatInfo object for the current culture is returned.
'
Public Function GetInstance(ByVal FormatProvider As IFormatProvider) As NumberFormatInfo
    If Not FormatProvider Is Nothing Then
        Set GetInstance = FormatProvider.GetFormat("NumberFormatInfo")
        If GetInstance Is Nothing Then Set GetInstance = CurrentInfo
    Else
        Set GetInstance = CurrentInfo
    End If
End Function

''
' Returns a read-only version of the NumberFormatInfo object.
'
' @param nfi The NumberFormatInfo object to create a read-only version of.
' @return A read-only version of the NumberFormatInfo object.
'
Public Function ReadOnly(ByVal nfi As NumberFormatInfo) As NumberFormatInfo
    If nfi Is Nothing Then _
        Throw New ArgumentNullException
    
    If nfi.IsReadOnly Then
        Set ReadOnly = nfi
    Else
        Set ReadOnly = nfi.Clone
        ReadOnly.IsReadOnly = True
    End If
End Function

