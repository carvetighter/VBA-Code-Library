Attribute VB_Name = "Mod_Date"
Option Explicit
Option Base 1
Function Date_GetHoursMinutesSeconds(stringTime As String, intHoursMinutesSeconds) As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' this function will at a time formatted string and depending on the flag <intHoursMinutesSeconds> determine which number to return
' hours, minutes or seconds
'
' Requirements:
' none
'
' Inputs:
' stringTime
' Type: string
' Desc: this is a string formatted as [h]:mm:ss.##
'
' intHoursMinutesSeconds
' Type: integer
' Desc: this is a flag to dtermine which number to return, hours, minutes or seconds
' hours -> 0
' minutes -> 1
' seconds -> 2
'
' Important Info:
' None
'
' Return:
' Type: long
' Desc: the number based on the flag passed, hours, minutes or seconds
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
 
' declare
Dim intColon01 As Integer, intColon02 As Integer
Dim longHours As Long, longMinutes As Long, longSeconds As Long, longReturnValue As Long
 
' initialize
intColon01 = 0
intColon02 = 0
longHours = 0
longMinutes = 0
longSeconds = 0
longReturnValue = 0
 
' begin
stringTime = Trim(stringTime)
 
' what number is desired, hours, minutes, seconds
Select Case intHoursMinutesSeconds
    ' hours
    Case 0:
       intColon01 = InStr(1, stringTime, ":", vbBinaryCompare)
        longReturnValue = CLng(Left(stringTime, intColon01 - 1))
    ' minutes
    Case 1:
        intColon01 = InStr(1, stringTime, ":", vbBinaryCompare)
        longHours = CLng(Left(stringTime, intColon01 - 1))
        longMinutes = CLng(Mid(stringTime, intColon01 + 1, 2))
        longReturnValue = (longHours * 60) + longMinutes
    ' seconds
    Case 2:
        intColon01 = InStr(1, stringTime, ":", vbBinaryCompare)
        intColon02 = InStr(intColon01 + 1, stringTime, ":", vbBinaryCompare)
        longHours = CLng(Left(stringTime, intColon01 - 1))
        longMinutes = CLng(Mid(stringTime, intColon01 + 1, 2))
        longSeconds = CLng(CDbl(Mid(stringTime, intColon02 + 1, Len(stringTime) - intColon02)))
        longReturnValue = (longHours * 120) + (longMinutes * 60) + longSeconds
    Case Else ' do nothing
End Select
 
' return value
Date_GetHoursMinutesSeconds = longReturnValue
End Function

