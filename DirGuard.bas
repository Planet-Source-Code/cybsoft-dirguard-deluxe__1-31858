Attribute VB_Name = "modDirGuard"
Public Function CAttr(AttNum As Integer) As String
AttrValue = AttNum


        Select Case AttrValue
            Case 0
                Attresult = "None"
            Case 1
                Attresult = "+R"
            Case 2
                Attresult = "+H"
            Case 3
                Attresult = "+H +R"
            Case 4
                Attresult = "+S"
            Case 5
                Attresult = "+R +S"
            Case 6
                Attresult = "+H +S"
            Case 7
                Attresult = "+H +R +S"
            Case 32
                Attresult = "+A"
            Case 33
                Attresult = "+A +R"
            Case 34
                Attresult = "+A +H"
            Case 35
                Attresult = "+A +H +R"
            Case 36
                Attresult = "+A +S"
            Case 37
                Attresult = "+A +R +S"
            Case 38
                Attresult = "+A +H +S"
            Case 39
                Attresult = "+A +H +R +S"
            Case 2048                           ' 20xx-series WinNT only
                Attresult = "+C"
            Case 2049
                Attresult = "+C +R"
            Case 2050
                Attresult = "+C + H"
            Case 2051
                Attresult = "+C +H +R"
            Case 2080
                Attresult = "+A +C"
            Case 2081
                Attresult = "+A +C +R"
            Case 2082
                Attresult = "+A +C +H"
            Case 2083
                Attresult = "+A +C +H +R"
            Case 2087
                Attresult = "+A +C +H +R +S"
         End Select
         
    CAttr = Attresult

End Function

Public Function SizeResult(StSize As Double, NewSize As Double) As String
StoredSz = StSize
NewSz = NewSize

Difference = NewSz - StoredSz

If Difference > 0 Then
    Difference = "+" & Difference & " Bits"
    Else
    Difference = Difference & " Bits"
End If
SizeResult = Difference
End Function
