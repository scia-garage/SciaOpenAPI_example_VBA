Attribute VB_Name = "guid"
Public Function Get_NewGUID() As String
    'Returns GUID as string 36 characters long

    Randomize

    Dim r1a As Long
    Dim r1b As Long
    Dim r2 As Long
    Dim r3 As Long
    Dim r4 As Long
    Dim r5a As Long
    Dim r5b As Long
    Dim r5c As Long

    'randomValue = CInt(Math.Floor((upperbound - lowerbound + 1) * Rnd())) + lowerbound
    r1a = RandomBetween(0, 65535)
    r1b = RandomBetween(0, 65535)
    r2 = RandomBetween(0, 65535)
    r3 = RandomBetween(16384, 20479)
    r4 = RandomBetween(32768, 49151)
    r5a = RandomBetween(0, 65535)
    r5b = RandomBetween(0, 65535)
    r5c = RandomBetween(0, 65535)

    Get_NewGUID = (PadHex(r1a, 4) & PadHex(r1b, 4) & "-" & PadHex(r2, 4) & "-" & PadHex(r3, 4) & "-" & PadHex(r4, 4) & "-" & PadHex(r5a, 4) & PadHex(r5b, 4) & PadHex(r5c, 4))

End Function

Public Function Floor(ByVal X As Double, Optional ByVal Factor As Double = 1) As Double
    'From: http://www.tek-tips.com/faqs.cfm?fid=5031
    ' X is the value you want to round
    ' Factor is the multiple to which you want to round
        Floor = Int(X / Factor) * Factor
End Function

Public Function RandomBetween(ByVal StartRange As Long, ByVal EndRange As Long) As Long
    'Based on https://msdn.microsoft.com/en-us/library/f7s023d2(v=vs.90).aspx
    '         randomValue = CInt(Math.Floor((upperbound - lowerbound + 1) * Rnd())) + lowerbound
        RandomBetween = CLng(Floor((EndRange - StartRange + 1) * Rnd())) + StartRange
End Function

Public Function PadLeft(text As Variant, totalLength As Integer, padCharacter As String) As String
    'Based on https://stackoverflow.com/questions/12060347/any-method-equivalent-to-padleft-padright
    ' with a little more checking of inputs

    Dim s As String
    Dim inputLength As Integer
    s = CStr(text)
    inputLength = Len(s)

    If padCharacter = "" Then
        padCharacter = " "
    ElseIf Len(padCharacter) > 1 Then
        padCharacter = Left(padCharacter, 1)
    End If

    If inputLength < totalLength Then
        PadLeft = String(totalLength - inputLength, padCharacter) & s
    Else
        PadLeft = s
    End If

End Function

Public Function PadHex(number As Long, length As Integer) As String
    PadHex = PadLeft(Hex(number), 4, "0")
End Function
