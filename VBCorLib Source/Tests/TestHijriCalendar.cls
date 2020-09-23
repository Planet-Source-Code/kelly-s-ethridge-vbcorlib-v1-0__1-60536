VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestHijriCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestHijriCalendar
'
Option Explicit

Private cal As HijriCalendar

Public Sub Setup()
    Set cal = New HijriCalendar
End Sub

Public Sub TestIsLeapYear()
    Assert.IsFalse cal.IsLeapYear(2000)
    Assert.IsTrue cal.IsLeapYear(2001)
    Assert.IsFalse cal.IsLeapYear(2002)
End Sub

Public Sub TestGetYear()
    Assert.Equals 1421, cal.GetYear(#1/1/2001#)
    Assert.Equals 1422, cal.GetYear(#12/31/2001#)
End Sub

Public Sub TestGetMonth()
    Assert.Equals 10, cal.GetMonth(#1/1/2001#)
    Assert.Equals 1, cal.GetMonth(#4/1/2002#)
End Sub

Public Sub TestGetDayOfMonth()
    Assert.Equals 6, cal.GetDayOfMonth(#1/1/2001#)
    Assert.Equals 19, cal.GetDayOfMonth(#4/1/2002#)
End Sub

Public Sub TestGetDayOfYear()
    Assert.Equals 272, cal.GetDayOfYear(#1/1/2001#)
    Assert.Equals 19, cal.GetDayOfYear(#4/1/2002#)
End Sub

Public Sub TestToDateTime()
    Assert.Equals #12/27/2562 1:01:01 AM#, cal.ToDateTime(2001, 1, 1, 1, 1, 1, 1).ToOADate
    Assert.Equals #3/15/2564 1:01:01 AM#, cal.ToDateTime(2002, 4, 1, 1, 1, 1, 1).ToOADate
End Sub

Public Sub TestAddYears()
    Assert.Equals #12/21/2000#, cal.AddYears(#1/1/2000#, 1).ToOADate
    Assert.Equals #3/22/2003#, cal.AddYears(#4/1/2002#, 1).ToOADate
End Sub

Public Sub TestAddDays()
    Assert.Equals #9/27/2002#, cal.AddDays(#1/1/2000#, 1000).ToOADate
    Assert.Equals #12/26/2004#, cal.AddDays(#4/1/2002#, 1000).ToOADate
End Sub

Public Sub TestDayOfYear()
    Dim d As cDateTime
    Dim sr As StreamReader
    Dim i As Long
    
    Set sr = file.OpenText(App.Path & "\Dummy Test Files\TesthijriCalendar_DayOfYear.txt")
    Set d = cDateTime.FromOADate(#1/1/2000#)
    
    For i = 1 To 500
        Assert.Equals CLng(sr.ReadLine), cal.GetDayOfYear(d)
        Set d = cal.AddDays(d, 1)
    Next i
    sr.CloseReader
End Sub

Public Sub TestGetWeekOfYear()
    Dim sr As StreamReader
    Dim i As Long
    Dim dt As cDateTime
    Dim j As Long

    Set sr = NewStreamReader(App.Path & "\Dummy Test Files\TestHijriCalendar_WeekOfYear.txt")
    Set dt = cDateTime.FromOADate(#1/1/2000#)

    ' tests using CalendarWeekRule.FirstDay and first day of week of DayOfWeek.Friday
    For i = 1 To 366
        j = sr.ReadLine
        Assert.Equals j, cal.GetWeekOfYear(dt, CalendarWeekRule.FirstDay, DayOfWeek.Friday), "Wrong week for " & dt.ToString & " during first test."
        Set dt = cal.AddDays(dt, 1)
    Next i

    ' tests using CalendarWeekRule.FirstFourDayWeek and first day of week of DayOfWeek.Thursday
    For i = 1 To 365
        j = sr.ReadLine
        Assert.Equals j, cal.GetWeekOfYear(dt, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Thursday), "Wrong week for " & dt.ToString & " during second test."
        Set dt = cal.AddDays(dt, 1)
    Next i

    ' tests using CalendarWeekRule.FirstFullWeek and first day of week of DayOfWeek.Wednesday
    For i = 1 To 365
        j = sr.ReadLine
        Assert.Equals j, cal.GetWeekOfYear(dt, CalendarWeekRule.FirstFullWeek, DayOfWeek.Wednesday), "Wrong week for " & dt.ToString & " during third test."
        Set dt = cal.AddDays(dt, 1)
    Next i

    sr.CloseReader
End Sub

