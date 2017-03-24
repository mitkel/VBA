'---------------------------------------------------------------------------------------------
Function swieto(d As Date) As Boolean

If Day(d) = 1 And Month(d) = 1 Then swieto = True 'nowy rok
If Day(d) = 6 And Month(d) = 1 Then swieto = True 'trzech króli
If Day(d) = 1 And Month(d) = 5 Then swieto = True '1 maj
If Day(d) = 3 And Month(d) = 5 Then swieto = True '3 maj
If Day(d) = 15 And Month(d) = 8 Then swieto = True '15 sierpien
If Day(d) = 1 And Month(d) = 11 Then swieto = True '1 listopad
If Day(d) = 11 And Month(d) = 11 Then swieto = True '11 listopad
If Day(d) = 25 And Month(d) = 12 Then swieto = True '25 grudzien
If Day(d) = 26 And Month(d) = 12 Then swieto = True '26 grudzien

'===================wyznaczanie Wielkanocy i Bozego Ciala=====================================
If Month(d) = 3 Or Month(d) = 4 Or Month(d) = 5 Or Month(d) = 6 Then
   sta = 24 'stała dla dat z przedziału 1900 - 2099
   stb = 5  'stała dla dat z przedziału 1900 - 2099
   a = Year(d) Mod 19
   b = Year(d) Mod 4
   c = Year(d) Mod 7
   g = (a * 19 + sta) Mod 30
   e = (2 * b + 4 * c + 6 * g + stb) Mod 7
   f = g + e
   egg = DateAdd("d", f, DateSerial(Year(d), 3, 22))
   If g = 29 And e = 6 Then egg = DateSerial(Year(d), 4, 19)
   If g = 28 And e = 6 Then egg = DateSerial(Year(d), 4, 18)
   pon_wielkanocny = DateAdd("d", 1, egg)  'poniedzialek wielkanocny
   boze_cialo = DateAdd("d", 60, egg) 'Boze Cialo
End If
'===============================================================================================

If d = pon_wielkanocny Then swieto = True
If d = boze_cialo Then swieto = True

End Function

Function NastepnyRoboczy(d As Date) As Date
Dim d1 As Date

d1 = d + 1
line1:
  If Weekday(d1, 2) = 6 Or Weekday(d1, 2) = 7 Or swieto(d1) = True Then
     d1 = d1 + 1
     GoTo line1
  End If

NastepnyRoboczy = d1

End Function

Function PoprzedniRoboczy(d As Date) As Date
Dim d1 As Date

d1 = d - 1
line1:
  If Weekday(d1, 2) = 6 Or Weekday(d1, 2) = 7 Or swieto(d1) = True Then
     d1 = d1 - 1
     GoTo line1
  End If

PoprzedniRoboczy = d1

End Function
Function CzyRoboczy(d As Date) As Boolean
    
    If PoprzedniRoboczy(NastepnyRoboczy(d)) = d Then
        CzyRoboczy = True
    Else
        CzyRoboczy = False
    End If

End Function
Function Month_end(d As Date) As Date

    d1 = DateSerial(Year(d), Month(d), 1)
    Month_end = PoprzedniRoboczy(DateSerial(Year(DateAdd("m", 1, d1)), Month(DateAdd("m", 1, d1)), 1))

End Function
Function Year_end(d As Date) As Date

    d1 = DateSerial(Year(d), Month(d), 1)
    Year_end = PoprzedniRoboczy(DateSerial(Year(DateAdd("y", 1, d1)), 1, 1))

End Function