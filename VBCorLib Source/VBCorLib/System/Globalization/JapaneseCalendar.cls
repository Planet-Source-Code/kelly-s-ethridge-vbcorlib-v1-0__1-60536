VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JapaneseCalendar"
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
'    Module: JapaneseCalendar
'

''
' Provides functions for manipulate Japanese dates.
'
' @see Calendar
'
Option Explicit
Implements cObject
Implements Calendar

Private Const ERA_MEIJI As Currency = 58938451200000@
Private Const ERA_TAISHO As Currency = 60323443200000@
Private Const ERA_SHOWA As Currency = 60777993600000@
Private Const ERA_HEISEI As Currency = 62735817600000@

Private Const ERA_HEISEI_YR2 As Currency = 62766748800000@
Private Const ERA_SHOWA_YR2 As Currency = 60778598400000@
Private Const ERA_TAISHO_YR2 As Currency = 60336835200000@
Private Const ERA_MEIJI_YR2 As Currency = 58948387200000@


Private mEras() As Long
Private mTwoDigitYearMax As Long
Private mEraLengths() As Long
Private mEraStart() As Currency
Private mEraStartYR2() As Currency



Public Property Get Eras() As Long()
    Eras = mEras
End Property

Public Property Get TwoDigitYearMax() As Long
    TwoDigitYearMax = mTwoDigitYearMax
End Property

Public Property Let TwoDigitYearMax(ByVal RHS As Long)
    If RHS < 100 Or RHS > 9999 Then _
        Throw Cor.NewArgumentOutOfRangeException("Year must be between 100 and 9999 inclusively.", "TwoDigitYearMax", RHS)
    
    mTwoDigitYearMax = RHS
End Property

Public Function AddDays(ByRef Time As Variant, ByVal Days As Long) As cDateTime
    Set AddDays = cDateTime.GetcDateTime(Time).AddDays(Days)
End Function

Public Function AddHours(ByRef Time As Variant, ByVal Hours As Long) As cDateTime
    Set AddHours = cDateTime.GetcDateTime(Time).AddHours(Hours)
End Function

Public Function AddMilliseconds(ByRef Time As Variant, ByVal Milliseconds As Double) As cDateTime
    Set AddMilliseconds = cDateTime.GetcDateTime(Time).AddMilliseconds(Milliseconds)
End Function

Public Function AddMinutes(ByRef Time As Variant, ByVal Minutes As Long) As cDateTime
    Set AddMinutes = cDateTime.GetcDateTime(Time).AddMinutes(Minutes)
End Function

Public Function AddMonths(ByRef Time As Variant, ByVal Months As Long) As cDateTime
    Set AddMonths = cDateTime.GetcDateTime(Time).AddMonths(Months)
End Function

Public Function AddSeconds(ByRef Time As Variant, ByVal Seconds As Long) As cDateTime
    Set AddSeconds = cDateTime.GetcDateTime(Time).AddSeconds(Seconds)
End Function

Public Function AddWeeks(ByRef Time As Variant, ByVal Weeks As Long) As cDateTime
    Set AddWeeks = AddDays(Time, Weeks * 7)
End Function

Public Function AddYears(ByRef Time As Variant, ByVal Years As Long) As cDateTime
    Set AddYears = cDateTime.GetcDateTime(Time).AddYears(Years)
End Function

Public Function GetDayOfMonth(ByRef Time As Variant) As Long
    GetDayOfMonth = cDateTime.GetcDateTime(Time).Day
End Function

Public Function GetDayOfWeek(ByRef Time As Variant) As DayOfWeek
    GetDayOfWeek = cDateTime.GetcDateTime(Time).DayOfWeek
End Function

Public Function GetDayOfYear(ByRef Time As Variant) As Long
    GetDayOfYear = cDateTime.GetcDateTime(Time).DayOfYear
End Function

Public Function GetEra(ByRef Time As Variant) As Long
    Dim dt As cDateTime
    Dim i As Long
    Dim ms As Currency
    
    Set dt = cDateTime.GetcDateTime(Time)
    VerifyDate dt
    
    ms = dt.TotalMilliseconds
    i = UBound(mEraStart)
    Do While i >= 0
        If ms >= mEraStart(i) Then Exit Do
        i = i - 1
    Loop
    If i < 0 Then i = 0
    GetEra = i + 1
End Function

Public Function GetHour(ByRef Time As Variant) As Long
    GetHour = cDateTime.GetcDateTime(Time).Hour
End Function

Public Function GetMinute(ByRef Time As Variant) As Long
    GetMinute = cDateTime.GetcDateTime(Time).Minute
End Function

Public Function GetMonth(ByRef Time As Variant) As Long
    GetMonth = cDateTime.GetcDateTime(Time).Month
End Function

Public Function GetSecond(ByRef Time As Variant) As Long
    GetSecond = cDateTime.GetcDateTime(Time).Second
End Function

Public Function GetMilliseconds(ByRef Time As Variant) As Double
    GetMilliseconds = cDateTime.GetcDateTime(Time).Millisecond
End Function

Public Function GetDaysInMonth(ByVal Year As Long, ByVal Month As Long, Optional ByRef Era As Variant) As Long
    Year = GetGregorianYear(Year, Era)
    GetDaysInMonth = cDateTime.DaysInMonth(Year, Month)
End Function

Public Function GetYear(ByRef Time As Variant) As Long
    Dim dt As cDateTime
    Dim i As Long
    Dim ms As Currency
    
    Set dt = cDateTime.GetcDateTime(Time)
    VerifyDate dt
    
    ms = dt.TotalMilliseconds
    i = UBound(mEraStart)
    Do While i >= 0
        If ms >= mEraStart(i) Then Exit Do
        i = i - 1
    Loop
    If i < 0 Then i = 0
    
    If ms < mEraStartYR2(i) Then
        GetYear = 1
    Else
        GetYear = cDateTime.FromMilliseconds(ms - mEraStart(i)).Year + 1
    End If
End Function

Public Function IsLeapYear(ByVal Year As Long, Optional ByRef Era As Variant) As Boolean
    IsLeapYear = cDateTime.IsLeapYear(GetGregorianYear(Year, GetCurrentEra(Era)))
End Function

Public Function IsLeapMonth(ByVal Year As Long, ByVal Month As Long, Optional ByRef Era As Variant) As Boolean
    If Month < 1 Or Month > 12 Then _
        Throw Cor.NewArgumentOutOfRangeException("Month must be between 1 and 12.", "Month", Month)
    
    If Month = 2 Then
        IsLeapMonth = IsLeapYear(Year, Era)
    End If
End Function

Public Function IsLeapDay(ByVal Year As Long, ByVal Month As Long, ByVal Day As Long, Optional ByRef Era As Variant) As Boolean
    If Day < 1 Or Day > Me.GetDaysInMonth(Year, Month, Era) Then _
        Throw Cor.NewArgumentOutOfRangeException("Day must be between 1 and " & GetDaysInMonth(Year, Month, Era) & " for the specified month.")
    
    If Day = 29 Then IsLeapDay = IsLeapMonth(Year, Month, Era)
End Function

Public Function GetWeekOfYear(ByRef Time As Variant, ByVal Rule As CalendarWeekRule, ByVal FirstDayOfWeek As DayOfWeek) As Long
    GetWeekOfYear = InternalGetWeekOfYear(Time, Rule, FirstDayOfWeek, Me)
End Function

Public Function GetMonthsInYear(ByVal Year As Long, Optional ByRef Era As Variant) As Long
    Year = GetGregorianYear(Year, GetCurrentEra(Era))
    GetMonthsInYear = 12
End Function

Public Function ToDateTime(ByVal Year As Long, ByVal Month As Long, ByVal Day As Long, ByVal Hour As Long, ByVal Minute As Long, ByVal Second As Long, ByVal Millisecond As Long, Optional ByRef Era As Variant) As cDateTime
    Year = GetGregorianYear(Year, GetCurrentEra(Era))
    Set ToDateTime = Cor.NewcDateTime(Year, Month, Day, Hour, Minute, Second, Millisecond)
End Function

Public Function ToFourDigitYear(ByVal Year As Long) As Long
    If Year < 1 Or Year > 8011 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_Range, 1, 8011), "Year", Year)
        
    ToFourDigitYear = Year
End Function

Public Function GetDaysInYear(ByVal Year As Long, Optional ByRef Era As Variant) As Long
    If IsLeapYear(Year, Era) Then
        GetDaysInYear = 366
    Else
        GetDaysInYear = 365
    End If
End Function

''
' Returns a string representation of this object instance.
'
' @return String representing this instance.
Public Function ToString() As String
    ToString = App.Title & "." & TypeName(Me)
End Function

''
' Returns a boolean indicating if the value and this object
' instance are the same instance.
'
' @param value The value to compare equalit to.
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
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub VerifyDate(ByRef dt As cDateTime)
    If dt.TotalMilliseconds < ERA_MEIJI Then Throw Cor.NewArgumentOutOfRangeException("The Japanese calendar only supports dates from 9/8/1868 to 12/31/9999 12:59:59pm.", "Time", dt.ToString)
End Sub

Private Function GetGregorianYear(ByVal Year As Long, ByVal Era As Long) As Long
    Dim ms As Currency
    
    If Year < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum), "Year", Year)
    
    Era = Era - 1
    If Era < 0 Or Era > UBound(mEraLengths) Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_Range, 1, UBound(mEraLengths) + 1), "Era", Era)
    
    ms = mEraStart(Era)
    If Year > mEraLengths(Era) Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_Range, 0, mEraLengths(Era)), "Year", Year)
    
    GetGregorianYear = cDateTime.FromMilliseconds(ms).Year + Year - 1
End Function

Private Function GetCurrentEra(ByVal Era As Variant) As Long
    If IsMissing(Era) Then
        GetCurrentEra = 1
    Else
        If Era = 0 Then Era = 1
        If Era < 1 Or Era > UBound(mEras) + 1 Then _
            Throw Cor.NewArgumentOutOfRangeException("Japanese eras range from 1 to 4.", "Era", Era)
        GetCurrentEra = Era
    End If
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    mTwoDigitYearMax = 99
    mEras = cArray.NewArray(ciLong, 4, 3, 2, 1)
    mEraLengths = cArray.NewArray(ciLong, 45, 15, 64, 8011)
    mEraStart = cArray.NewArray(ciCurrency, ERA_MEIJI, ERA_TAISHO, ERA_SHOWA, ERA_HEISEI)
    mEraStartYR2 = cArray.NewArray(ciCurrency, ERA_MEIJI_YR2, ERA_TAISHO_YR2, ERA_SHOWA_YR2, ERA_HEISEI_YR2)
End Sub

Private Sub Class_ReadProperties(PropBag As PropertyBag)
    With PropBag
        mTwoDigitYearMax = .ReadProperty("TwoDigitYearMax")
    End With
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "TwoDigitYearMax", mTwoDigitYearMax
    End With
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
'   Calendar Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function Calendar_AddDays(Time As Variant, ByVal Days As Long) As cDateTime
    Set Calendar_AddDays = AddDays(Time, Days)
End Function

Private Function Calendar_AddHours(Time As Variant, ByVal Hours As Long) As cDateTime
    Set Calendar_AddHours = AddHours(Time, Hours)
End Function

Private Function Calendar_AddMilliseconds(Time As Variant, ByVal Milliseconds As Double) As cDateTime
    Set Calendar_AddMilliseconds = AddMilliseconds(Time, Milliseconds)
End Function

Private Function Calendar_AddMinutes(Time As Variant, ByVal Minutes As Long) As cDateTime
    Set Calendar_AddMinutes = AddMinutes(Time, Minutes)
End Function

Private Function Calendar_AddMonths(Time As Variant, ByVal Months As Long) As cDateTime
    Set Calendar_AddMonths = AddMonths(Time, Months)
End Function

Private Function Calendar_AddSeconds(Time As Variant, ByVal Seconds As Long) As cDateTime
    Set Calendar_AddSeconds = AddSeconds(Time, Seconds)
End Function

Private Function Calendar_AddWeeks(Time As Variant, ByVal Weeks As Long) As cDateTime
    Set Calendar_AddWeeks = AddWeeks(Time, Weeks)
End Function

Private Function Calendar_AddYears(Time As Variant, ByVal Years As Long) As cDateTime
    Set Calendar_AddYears = AddYears(Time, Years)
End Function

Private Function Calendar_Equals(Value As Variant) As Boolean
    Calendar_Equals = Equals(Value)
End Function

Private Property Get Calendar_Eras() As Long()
    Calendar_Eras = Eras
End Property

Private Function Calendar_GetDayOfMonth(Time As Variant) As Long
    Calendar_GetDayOfMonth = GetDayOfMonth(Time)
End Function

Private Function Calendar_GetDayOfWeek(Time As Variant) As DayOfWeek
    Calendar_GetDayOfWeek = GetDayOfWeek(Time)
End Function

Private Function Calendar_GetDayOfYear(Time As Variant) As Long
    Calendar_GetDayOfYear = GetDayOfYear(Time)
End Function

Private Function Calendar_GetDaysInMonth(ByVal Year As Long, ByVal Month As Long, Optional Era As Variant) As Long
    Calendar_GetDaysInMonth = GetDaysInMonth(Year, Month, Era)
End Function

Private Function Calendar_GetDaysInYear(ByVal Year As Long, Optional Era As Variant) As Long
    Calendar_GetDaysInYear = GetDaysInYear(Year, Era)
End Function

Private Function Calendar_GetEra(Time As Variant) As Long
    Calendar_GetEra (Time)
End Function

Private Function Calendar_GetHashCode() As Long
    Calendar_GetHashCode = GetHashCode
End Function

Private Function Calendar_GetHour(Time As Variant) As Long
    Calendar_GetHour = GetHour(Time)
End Function

Private Function Calendar_GetMilliseconds(Time As Variant) As Double
    Calendar_GetMilliseconds = GetMilliseconds(Time)
End Function

Private Function Calendar_GetMinute(Time As Variant) As Long
    Calendar_GetMinute = GetMinute(Time)
End Function

Private Function Calendar_GetMonth(Time As Variant) As Long
    Calendar_GetMonth = GetMonth(Time)
End Function

Private Function Calendar_GetMonthsInYear(ByVal Year As Long, Optional Era As Variant) As Long
    Calendar_GetMonthsInYear = GetMonthsInYear(Year, Era)
End Function

Private Function Calendar_GetSecond(Time As Variant) As Long
    Calendar_GetSecond = GetSecond(Time)
End Function

Private Function Calendar_GetWeekOfYear(Time As Variant, ByVal Rule As CalendarWeekRule, ByVal FirstDayOfWeek As DayOfWeek) As Long
    Calendar_GetWeekOfYear = GetWeekOfYear(Time, Rule, FirstDayOfWeek)
End Function

Private Function Calendar_GetYear(Time As Variant) As Long
    Calendar_GetYear = GetYear(Time)
End Function

Private Function Calendar_IsLeapDay(ByVal Year As Long, ByVal Month As Long, ByVal Day As Long, Optional Era As Variant) As Boolean
    Calendar_IsLeapDay = IsLeapDay(Year, Month, Day, Era)
End Function

Private Function Calendar_IsLeapMonth(ByVal Year As Long, ByVal Month As Long, Optional Era As Variant) As Boolean
    Calendar_IsLeapMonth = IsLeapMonth(Year, Month, Era)
End Function

Private Function Calendar_IsLeapYear(ByVal Year As Long, Optional Era As Variant) As Boolean
    Calendar_IsLeapYear = IsLeapYear(Year, Era)
End Function

Private Function Calendar_ToDateTime(ByVal Year As Long, ByVal Month As Long, ByVal Day As Long, ByVal Hour As Long, ByVal Minute As Long, ByVal Second As Long, ByVal Millisecond As Long, Optional Era As Variant) As cDateTime
    Set Calendar_ToDateTime = ToDateTime(Year, Month, Day, Hour, Minute, Second, Millisecond, Era)
End Function

Private Function Calendar_ToFourDigitYear(ByVal Year As Long) As Long
    Calendar_ToFourDigitYear = ToFourDigitYear(Year)
End Function

Private Function Calendar_ToString() As String
    Calendar_ToString = ToString
End Function

Private Property Let Calendar_TwoDigitYearMax(ByVal RHS As Long)
    TwoDigitYearMax = RHS
End Property

Private Property Get Calendar_TwoDigitYearMax() As Long
    Calendar_TwoDigitYearMax = TwoDigitYearMax
End Property

