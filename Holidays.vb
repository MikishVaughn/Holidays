Imports System.Math

Public Class EventDates

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Public Shared Function GetEventUtcDateString(eventName As String, eventYear As Integer) As String

        Dim midnight As String = "T00:00:00"
        Dim temp As Date
        eventName = LCase(eventName)

        ' Seasons and Easter - (Date-time is Calculated)
        If eventName = "spring" Or eventName = "summer" Or eventName = "autumn" Or eventName = "winter" Or eventName = "easter" Then
            Select Case eventName
                Case "spring"
                    temp = GetSpringUtcDate()
                Case "summer"
                    temp = GetSummerUtcDate()
                Case "autumn"
                    temp = GetAutumnUtcDate()
                Case "winter"
                    temp = GetWinterUtcDate()
                Case "easter"
                    temp = GetEasterUtcDate()
            End Select
            Return CStr(temp.Year) + "-" + Right("00" + CStr(temp.Month), 2) + "-" + Right("00" + CStr(temp.Day), 2) + " " + Right("00" + CStr(temp.Hour), 2) + ":" + Right("00" + CStr(temp.Minute), 2) + ":" + Right("00" + CStr(temp.Second), 2)
        End If

        ' Holidays - (Date is Fixed)
        Select Case eventName
            Case "new years"
                Return CStr(eventYear) + "-01-01" + midnight
            Case "valentines"
                Return CStr(eventYear) + "-02-14" + midnight
            Case "independence"
                Return CStr(eventYear) + "-07-04" + midnight
            Case "veterans"
                Return CStr(eventYear) + "-11-11" + midnight
            Case "halloween"
                Return CStr(eventYear) + "-10-31" + midnight
            Case "christmas"
                Return CStr(eventYear) + "-12-25" + midnight
        End Select
        ' Holidays - (Date is Calculated)
        Select Case eventName

            ' Month, Week, Day (Sun = 1)
            Case "elections"
                Return CalculateDateString(11, 1, 3)
            Case "mothers"
                Return CalculateDateString(5, 2, 1)
            Case "fathers"
                Return CalculateDateString(6, 3, 1)
            Case "memorial"
                'Memorial Day is usually the 4 Monday of the Month,
                'but sometimes it is the 5th Monday. The Replace() catches
                'those times and returns the correct date.
                Return Replace(CalculateDateString(5, 4, 2), "05-24", "05-31")
            Case "labor"
                Return CalculateDateString(9, 1, 2)
            Case "thanksgiving"
                Return CalculateDateString(11, 4, 5)
            Case "martinLutherKing"
                Return CalculateDateString(1, 3, 2)
            Case "presidents"
                Return CalculateDateString(2, 3, 2)
        End Select
        ' Birthdays Presidents (Date is Fixed)
        Select Case eventName
            Case "lincoln"
                Return CStr(eventYear) + "-02-12" + midnight
            Case "washington"
                Return CStr(eventYear) + "-02-22" + midnight
            Case "jefferson"
                Return CStr(eventYear) + "-04-13" + midnight
        End Select
        ' Birthdays Military (Date is Fixed)
        Select Case eventName
            Case "army"
                Return CStr(eventYear) + "-06-14" + midnight
            Case "airforce"
                Return CStr(eventYear) + "-09-18" + midnight
            Case "navy"
                Return CStr(eventYear) + "-10-13" + midnight
            Case "marinecorp"
                Return CStr(eventYear) + "-11-10" + midnight
            Case "nationalguard"
                Return CStr(eventYear) + "-12-13" + midnight
        End Select

        ' Return a default Date if target Date not found (Shouldn't happen)
        Return "1953-03-31" + midnight

    End Function

    Public Shared Function CalculateEasterUtcDate(Optional eventYear As Integer = 0) As Date

        ' Was a year passed in argument?
        ' If not, default to the current year.
        If eventYear = 0 Then
            eventYear = Date.UtcNow.Year
        End If

        Dim H = ((19 * (eventYear Mod 19)) + (eventYear \ 100) - ((eventYear \ 100) \ 4) - (((8 * (eventYear \ 100)) + 13) \ 25) + 15) Mod 30
        Dim M = ((eventYear Mod 19) + (11 * H)) \ 319
        Dim L = ((2 * ((eventYear \ 100) Mod 4)) + (2 * ((eventYear Mod 100) \ 4)) - ((eventYear Mod 100) Mod 4) - H + M + 32) Mod 7

        Return New DateTime(eventYear, ((H - M + L + 90) \ 25), ((H - M + L + ((H - M + L + 90) \ 25) + 19) Mod 32))

    End Function
    Public Shared Function CalculateSeasonUtcDate(seasonName As String, Optional eventYear As Integer = 0) As Date

        ' Was a year passed in argument?
        ' If not, default to the current year.
        If eventYear = 0 Then
            eventYear = Date.UtcNow.Year
        End If

        ' Define variables for the raised the power of the year Y, Y², Y³, Y⁴.
        Dim Y As Double = (eventYear - 2000) / 1000
        Dim Y2 As Double = Y * Y
        Dim Y3 As Double = Y2 * Y
        Dim Y4 As Double = Y3 * Y

        ' For the years +1000 to +3000.
        Dim JDE As Double

        Select Case LCase(seasonName)
            Case "spring" ' March Equinox (Spring)
                JDE = 2451623.80984 + (365242.37404 * Y) + (0.05169 * Y2) - (0.00411 * Y3) - (0.00057 * Y4)
            Case "summer" ' June Solstice (Summer)
                JDE = 2451716.56767 + (365241.62603 * Y) + (0.00325 * Y2) + (0.00888 * Y3) - (0.0003 * Y4)
            Case "autumn" ' September Equinox (Autumn)
                JDE = 2451810.21715 + (365242.01767 * Y) - (0.11575 * Y2) + (0.00337 * Y3) + (0.00078 * Y4)
            Case "winter" ' December Solstice (Winter)
                JDE = 2451900.05952 + (365242.74049 * Y) - (0.06223 * Y2) - (0.00823 * Y3) + (0.00032 * Y4)
        End Select

        ' Periodic Term array used below to find the sum 'S' for calculating the 'mean' Equinox or Solstice.
        Dim A As Double() = {485, 203, 199, 182, 156, 136, 77, 74, 70, 58, 52, 50, 45, 44, 29, 18, 17, 16, 14, 12, 12, 12, 9, 8}
        Dim B As Double() = {324.96, 337.23, 342.08, 27.85, 73.14, 171.52, 222.54, 296.72, 243.58, 119.81, 297.17, 21.02, 247.54, 325.15, 60.93, 155.12, 288.79, 198.04, 199.76, 95.39, 287.11, 320.81, 227.73, 15.45}
        Dim C As Double() = {1934.136, 32964.467, 20.186, 445267.112, 45036.886, 22518.443, 65928.934, 3034.906, 9037.513, 33718.147, 150.678, 2281.226, 29929.562, 31555.956, 4443.417, 67555.328, 4562.452, 62894.029, 31436.921, 14577.848, 31931.756, 34777.259, 1222.114, 16859.074}

        ' Define 'T' variable.
        Dim T As Double = (JDE - 2451545.0) / 36525

        ' S will hold the Equinox or Solstice Periodic Term sum after loop through the Terms arrays.
        ' i is the count iteration when reading the 24 A, B, and C Periotic Terms array data.
        Dim S As Double = 0
        Dim i As Integer = 0

        ' Loop through and add up the Periodic Terms.
        Do While i < 24

            ' The Periodic Term equation is:
            ' "S = ∑ A Cos(B + C * T)" Meaning the array of 0-23 terms are summed to S to calculate a 'mean'.
            ' NOTE: The argument for Cos() is given in degrees needing the "(({arg}* PI) / 180)" conversions.
            S += (A(i) * Cos(((B(i) * PI) / 180) + (((C(i) * PI) / 180) * T)))
            i += 1

        Loop ' (Now we have the sum S of the Equinox or Solstice Periodic Terms and can continue.)

        ' Define 'W' and 'Δλ (SW)' variables.
        Dim W As Double = 35999.373 * T - 2.47
        Dim SW As Double = 1 + 0.0334 * Cos(((W * PI) / 180)) + 0.0007 * Cos(2 * ((W * PI) / 180))

        ' Note: "2415019 + 0.5" = Start of Julian Century.
        Dim eventDateJDE = JDE - 2415019 + 0.5 + (0.00001 * S / SW)

        ' Conversion from 'Double' to 'Date' requires calling the 'Date.FromOADate'
        Dim eventUtcDate As Date = Date.FromOADate(eventDateJDE)

        '   MsgBox(eventDateJDE) ' For Debugging

        Return eventUtcDate

    End Function
    Private Shared Function CalculateDateString(eventMonth As Integer, eventWeek As Integer, eventDay As Integer) As String

        ' Build Date String  (based on Month, Week, and Day of Event with 1 = Sunday)
        ' Handy for events falling on days like the second Tuesday of a Month.
        '
        ' Input:
        '   eventMonth = Month of the Event
        '   eventWeek = Week number of Event (ie. Event is in the third week of the month.)
        '   evenDay = Day of Event. Sunday = 1, Monday = 2,...
        '
        ' Output:
        '   EventDate String
        '
        Dim eventYear As Integer = Date.UtcNow.Year
        Dim wkday As Integer = Weekday(Date.Parse(CStr(eventYear) + "-" + CStr(eventMonth) + "-" + CStr(8 - eventDay), System.Globalization.CultureInfo.InvariantCulture))
        Dim eventDate As Date = Date.Parse(CStr(eventYear) + "-" + CStr(eventMonth) + "-" + CStr(1 + (7 * eventWeek) - wkday), System.Globalization.CultureInfo.InvariantCulture).ToUniversalTime

        ' Catch to see if we are beyond the event date for this year.
        ' If so, then add a year and countdown to next year's event.
        If Date.UtcNow.Month > eventDate.Month Then
            eventYear = Date.UtcNow.Year + 1
        Else
            If Date.UtcNow.Month = eventDate.Month And Date.UtcNow.Day > eventDate.Day Then
                eventYear = Date.UtcNow.Year + 1
            End If
        End If
        wkday = Weekday(Date.Parse(CStr(eventYear) + "-" + CStr(eventMonth) + "-" + CStr(8 - eventDay), System.Globalization.CultureInfo.InvariantCulture))
        eventDate = Date.Parse(CStr(eventYear) + "-" + CStr(eventMonth) + "-" + CStr(1 + (7 * eventWeek) - wkday), System.Globalization.CultureInfo.InvariantCulture).ToUniversalTime

        Return CStr(eventDate.Year) + "-" + Right("00" + CStr(eventDate.Month), 2) + "-" + Right("00" + CStr(eventDate.Day), 2) + "T00:00:00"

    End Function

    Public Shared Function GetEasterUtcDate(Optional eventYear As Integer = 0) As Date

        Dim eventUtcDate As Date

        If eventYear > 0 Then ' was a year passed?
            eventUtcDate = CalculateEasterUtcDate(eventYear)
        Else
            ' no year passed, use the current year and check
            ' to see if we are beyond event date for this year.
            eventUtcDate = CalculateEasterUtcDate(Date.UtcNow.Year)
            If eventUtcDate.AddDays(1) < Date.UtcNow Then
                eventUtcDate = CalculateEasterUtcDate(Date.UtcNow.Year + 1)
            End If
        End If

        Return (eventUtcDate)

    End Function
    Public Shared Function GetSpringUtcDate(Optional eventYear As Integer = 0) As Date

        Dim eventSeason As String = "spring"
        Dim eventDate As Date

        If eventYear > 0 Then ' was a year passed?
            eventDate = CalculateSeasonUtcDate(eventSeason, eventYear)
        Else
            ' no year passed, use the current year and check
            ' to see if we are beyond event date for this year.
            eventDate = CalculateSeasonUtcDate(eventSeason, Date.UtcNow.Year)
            If eventDate.AddDays(1) < Date.UtcNow Then
                eventDate = CalculateSeasonUtcDate(eventSeason, Date.UtcNow.Year + 1)
            End If
        End If

        Return (eventDate)

    End Function
    Public Shared Function GetSummerUtcDate(Optional eventYear As Integer = 0) As Date

        Dim eventSeason As String = "summer"
        Dim eventDate As Date

        If eventYear > 0 Then ' was a year passed?
            eventDate = CalculateSeasonUtcDate(eventSeason, eventYear)
        Else
            ' no year passed, use the current year and check
            ' to see if we are beyond event date for this year.
            eventDate = CalculateSeasonUtcDate(eventSeason, Date.UtcNow.Year)
            If eventDate.AddDays(1) < Date.UtcNow Then
                eventDate = CalculateSeasonUtcDate(eventSeason, Date.UtcNow.Year + 1)
            End If
        End If

        Return (eventDate)

    End Function
    Public Shared Function GetAutumnUtcDate(Optional eventYear As Integer = 0) As Date

        Dim eventSeason As String = "autumn"
        Dim eventDate As Date

        If eventYear > 0 Then ' was a year passed?
            eventDate = CalculateSeasonUtcDate(eventSeason, eventYear)
        Else
            ' no year passed, use the current year and check
            ' to see if we are beyond event date for this year.
            eventDate = CalculateSeasonUtcDate(eventSeason, Date.UtcNow.Year)
            If eventDate.AddDays(1) < Date.UtcNow Then
                eventDate = CalculateSeasonUtcDate(eventSeason, Date.UtcNow.Year + 1)
            End If
        End If

        Return (eventDate)

    End Function
    Public Shared Function GetWinterUtcDate(Optional eventYear As Integer = 0) As Date

        Dim eventSeason As String = "winter"
        Dim eventDate As Date

        If eventYear > 0 Then ' was a year passed?
            eventDate = CalculateSeasonUtcDate(eventSeason, eventYear)
        Else
            ' no year passed, use the current year and check
            ' to see if we are beyond event date for this year.
            eventDate = CalculateSeasonUtcDate(eventSeason, Date.UtcNow.Year)
            If eventDate.AddDays(1) < Date.UtcNow Then
                eventDate = CalculateSeasonUtcDate(eventSeason, Date.UtcNow.Year + 1)
            End If
        End If

        Return eventDate

    End Function

    Public Overrides Function Equals(obj As Object) As Boolean
        Return MyBase.Equals(obj)
    End Function
    Public Overrides Function GetHashCode() As Integer
        Return MyBase.GetHashCode()
    End Function
    Public Overrides Function ToString() As String
        Return MyBase.ToString()
    End Function

End Class