'Elfproef voor banknr of sofinr 11 / 11 proof 
Function IsSofinr(strNum As String) As Boolean
    Dim aNUM(1 To 9) As Integer, i As Integer, j As Integer, Tot As Integer
    
    IsSofinr = True
    
    If Len(Trim(strNum)) > 9 Then
        IsSofinr = False
        Exit Function
    End If
    
    strNum = Format(strNum, "000000000")
    'MsgBox strNum
    j = UBound(aNUM)
    For i = 1 To 9
        aNUM(i) = Val(Mid(strNum, i, 1)) * j
        If i <> 9 Then Tot = Tot + aNUM(i)
        j = j - 1
    Next
    
    If (Tot Mod 11) = aNUM(9) Then
        IsSofinr = True
    Else
        IsSofinr = False
    End If
    
End Function
