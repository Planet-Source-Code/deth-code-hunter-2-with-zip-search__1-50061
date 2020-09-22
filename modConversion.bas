Attribute VB_Name = "modConversion"
Option Explicit

'calculates gb,mb,kb,b from total
Function FormatBytes(ByVal BytesCount As Double) As String

    If BytesCount >= 1024000000 Then
        FormatBytes = Format$(BytesCount / 1024000000, "###,###,###,###,##0.00") & " Gb"
    ElseIf BytesCount >= 1024000 Then    'Checks to see if its big enough to convert into MB
        FormatBytes = Format$(BytesCount / 1024000, "###,###,##0.00") & " Mb"
    ElseIf BytesCount >= 1024 Then    'Checks to see if file is big enough to convert into KB
        FormatBytes = Format$(BytesCount / 1024, "0.00") & " Kb"
    Else
        FormatBytes = CStr(BytesCount) & " Bytes"
    End If

End Function

'calulates hours,minutes and seconds from number of milliseconds
Function CalculateTime(ByVal sngTotal As Single) As String

  Dim lhour As Single, lmin As Single, lsec As Single

    If sngTotal >= 3600000 Then
        lhour = sngTotal \ 3600000
        sngTotal = sngTotal - (lhour * 3600000)
        CalculateTime = CStr(lhour) & " Hr. "
    End If

    If sngTotal >= 60000 Then
        lmin = sngTotal \ 60000
        sngTotal = sngTotal - (lmin * 60000)
        CalculateTime = CalculateTime & CStr(lmin) & " Min. "
    End If

    If sngTotal >= 1000 Then
        lsec = sngTotal \ 1000
        CalculateTime = CalculateTime & CStr(lsec) & " Sec."
    Else
        CalculateTime = CalculateTime & ".0" & Left$(CStr(sngTotal), 1) & " Sec."
    End If

End Function


