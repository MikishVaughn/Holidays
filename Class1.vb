Public Class EasterDate
    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    '*****************
    '* GetEasterDate *
    '*****************
    '
    ' Returns the EventDate of the next Easter
    '
    Public Shared Function GetEasterDate() As Date

        Dim EventDate As Date = CalculateEasterDate(Date.UtcNow.Year)
        If EventDate.AddDays(1) < Date.UtcNow Then
            EventDate = CalculateEasterDate(Date.UtcNow.Year + 1)
        End If

        Return (EventDate)

    End Function


    '**********************
    '* CalculateEventDate *
    '**********************
    '
    ' Returns the EventDate of the next Easter based on the year being passed.
    '
    Public Shared Function CalculateEasterDate(eventYear As Integer) As Date

        Dim H = ((19 * (eventYear Mod 19)) + (eventYear \ 100) - ((eventYear \ 100) \ 4) - (((8 * (eventYear \ 100)) + 13) \ 25) + 15) Mod 30
        Dim M = ((eventYear Mod 19) + (11 * H)) \ 319
        Dim L = ((2 * ((eventYear \ 100) Mod 4)) + (2 * ((eventYear Mod 100) \ 4)) - ((eventYear Mod 100) Mod 4) - H + M + 32) Mod 7

        Return New DateTime(eventYear, ((H - M + L + 90) \ 25), ((H - M + L + ((H - M + L + 90) \ 25) + 19) Mod 32))

    End Function

    Public Overrides Function ToString() As String
        Return MyBase.ToString()
    End Function

    Public Overrides Function Equals(obj As Object) As Boolean
        Return MyBase.Equals(obj)
    End Function

    Public Overrides Function GetHashCode() As Integer
        Return MyBase.GetHashCode()
    End Function


End Class




