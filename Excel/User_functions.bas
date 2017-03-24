'...zwraca nazwę kolumny jako napis
Public Function ColumnLetter(ColumnNumber As Integer) As String
    Dim n As Integer
    Dim c As Byte
    Dim s As String

    n = ColumnNumber
    Do
        c = ((n - 1) Mod 26)
        s = Chr(c + 65) & s
        n = (n - c) \ 26
    Loop While n > 0
    ColumnLetter = s
End Function

'......Przełączenie między widokiem R1C1 a A1
'......skrót ctrl+shift+R
Public Sub ChangeReferenceStyle()
    With Application
        If (.ReferenceStyle = xlA1) Then
            .ReferenceStyle = xlR1C1
        End
        Else:
            .ReferenceStyle = xlA1
        End If
    End With
End Sub

Sub usun_funkcje()
'Zamienia w zaznaczeniu # na =
'ctrl + shift + e

    Selection.Replace What:="=", Replacement:="#", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

End Sub

Sub przywroc_funckje()
'zamienia w zaznaczeniu = na #
'ctrl + shift + w

   Selection.Replace What:="#", Replacement:="=", LookAt:=xlPart, _
       SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
       ReplaceFormat:=False

End Sub

Sub zamien_na_wartosci()
'zamienia formuły na wartości
'ctrl + shift + s

    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
End Sub

Sub kopiowanie( _
    ByVal sPath1 As String, _
    ByVal sPath2 As String, _
    Optional ByVal sRplWhat As String = "", _
    Optional ByVal sRplWith As String = "")
    'sPath1, sPath2 - ścieżki do folderów [odp. źródłowego i docelowego] z plikami do przekopiowania
    'folder docelowy nie musi istnieć (wtedy pojawi się zapytane, czy przekopiować do istniejącego już folderu)
    'ścieżki MOGĄ, ale NIE MUSZĄ, kończyć się backslashem
    '
    'sRplWhat - fragment nazwy starych plików, który ma być zamieniony
    'sRplWith - na co ma być zamieniony
    'nie podanie powyższych argumentów, spowoduje, że przekopiują się stare nazwy
    'jeśli w jakimś pliku sRplWhat nie występuje, to oczywiście nic się nie dzieje, przekopiowywana jest stara nazwa
    'jeśli sRplWith jest pusty (a sRplWhat nie), to zostaniemy zapytani, czy chcemy usunąć dany fragment (nie zastępując go niczym)
    
    
    Dim bExists1, bExists2 As Boolean
    Dim file, sName As String
    
    'sprawdzenie czy lokalizacje istnieją
    bExists1 = (Len(Dir(sPath1, vbDirectory)) <> 0)
    bExists2 = (Len(Dir(sPath2, vbDirectory)) <> 0)
    
    If Not bExists1 Then
        MsgBox "Brak ścieżki " & sPath1
    End If
    
    If Not bExists2 Then
        MkDir sPath2
    Else
        If MsgBox("Ścieżka " & sPath2 & " już istnieje! Czy chcesz przekopiować do niego pliki?", vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If
    
    'dodanie backslashy, jeśli do tej pory nie było
    If Right(sPath1, 1) <> "/" Then
        sPath1 = sPath1 & "/"
    End If
    If Right(sPath2, 1) <> "/" Then
        sPath2 = sPath2 & "/"
    End If
    
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    file = Dir(sPath1, vbNormal)
    While (file <> "")
        'kopiowanie plików
        FSO.copyfile sPath1 & file, sPath2, False
        
        'podmiana nazw, jeśli jest na co
        If sRplWhat <> "" And sRplWith <> "" Then
        sName = Replace(file, sRplWhat, sRplWith)
        Name sPath2 & file As sPath2 & sName
        ElseIf sRplWhat <> "" And sRplWith = "" Then
            If MsgBox("Czy chcesz usunąć z nazw plików fragment """ & sRplWhat & """ i nie podsawić nic w zamian?", vbYesNo) = vbYes Then
                sName = Replace(file, sRplWhat, "")
                Name sPath2 & file As sPath2 & sName
            End If
        End If
        file = Dir
    Wend
    
End Sub

'wymnaża komórki przez zadaną wartość
Public Sub Multiply()

    Dim c As Range

    For Each c In Selection
        c.Value = c.Value * ThisWorkbook.Sheets("PANEL").Range("multiplier").Value
    Next

End Sub

'dzieli komórki przez zadaną wartość
Public Sub Divide()

    Dim c As Range

    For Each c In Selection
        c.Value = c.Value / ThisWorkbook.Sheets("PANEL").Range("multiplier").Value
    Next

End Sub

Function LastRow(sh As Worksheet)
    On Error Resume Next
    LastRow = sh.Cells.Find(What:="*", _
                            After:=sh.Range("A1"), _
                            LookAt:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).Row
    On Error GoTo 0
End Function


Function LastCol(sh As Worksheet)
    On Error Resume Next
    LastCol = sh.Cells.Find(What:="*", _
                            After:=sh.Range("A1"), _
                            LookAt:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByColumns, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).Column
    On Error GoTo 0
End Function

Sub change_sign()
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
    
    For Each s In Selection
        If Left(s.Formula, 1) = "=" Then
            s.Formula = "=-(" & Right(s.Formula, Len(s.Formula) - 1) & ")"
        Else
            s.Value = -s.Value
        End If
    Next
    
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
End Sub

Sub BreakLinks( _
    ByRef w As Workbook)
Application.DisplayAlerts = False

    Dim Links As Variant
    Links = w.LinkSources(Type:=xlLinkTypeExcelLinks)
    For i = 1 To UBound(Links)
    ActiveWorkbook.BreakLink _
        Name:=Links(i), _
        Type:=xlLinkTypeExcelLinks
    Next i
    
Application.DisplayAlerts = True
End Sub
    
Sub change_filter(ByVal bChange, ByRef sh As Worksheet)

    If Not (sh.AutoFilterMode And bChange) Then
        If bChange Then
            sh.Range("a1").AutoFilter
        Else
            sh.AutoFilterMode = False
        End If
    End If
End Sub


Public Sub SQL2Range( _
    ByRef rng As Range, _
    ByVal SQL As String, _
    ByVal sCon As String)
Application.ScreenUpdating = False
    
    Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim QT As Excel.QueryTable
    
    cn.Open sCon
    cn.CommandTimeout = 0
    rs.Open SQL, cn, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    rng.Activate
    rng.CopyFromRecordset rs
    
    rs.Close
    cn.Close
Application.ScreenUpdating = True
End Sub

Public Sub Force_Recalc()
    
    ThisWorkbook.Sheets("Kalkulator").Cells.Replace What:="=", Replacement:="=", LookAt:=xlPart, SearchOrder _
        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False

End Sub