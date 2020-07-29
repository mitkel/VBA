Option Explicit

Const Sciezka_do_szablonow As String = ""
Const ALCOPath As String = ""
Const tmpPath As String = ""
Const FXPath As String = ""

Sub AlcoRap()
    Dim olApp As Outlook.Application
    Set olApp = GetObject(, "Outlook.Application")
    Dim objMsg As Outlook.MailItem
    
    Dim dData As Date
    dData = DATES.PoprzedniRoboczy(Date)
    Dim sName As String
    sName = Sciezka_do_szablonow & "AlcoRap.oft"
    
    Set objMsg = olApp.CreateItemFromTemplate(sName)
    With objMsg
        .Subject = "ALM_portfolio_" & Format(dData, "ddmmyyyy")
        .Attachments.Add ALCOPath & "ALM_portfolio_" & Format(dData, "yyyymmdd") & ".xlsx"
        .Display
    End With
    
End Sub

Sub MarketData()
'To makro zapisuje dane market_data i wkleja je do zakladki Rynek w pliku ALCOnewAlter.xlsm

'Zapisywanie danych w pomocnicznym pliku market_data.xlsx
    Dim xlApp As Excel.Application
    Dim wData As Excel.Workbook
    Dim objAttachments As Outlook.Attachments

    Dim olApp As Outlook.Application
    Dim objMsg As Outlook.MailItem
    
    Dim objSelection As Outlook.Selection
    
    Set xlApp = CreateObject("Excel.Application")
    Set olApp = CreateObject("Outlook.Application")
    xlApp.ScreenUpdating = False
    
    xlApp.Visible = False
    Set objSelection = olApp.ActiveExplorer.Selection

    For Each objMsg In objSelection
    objMsg.Attachments.Item(1).SaveAsFile (tmpPath & "market_data.xls")
    Next
''' Przeklejanie danych do wlasciwego pliku

Dim x As Workbook
Dim y As Workbook

'## Open both workbooks first:
Set x = Workbooks.Open(tmpPath & "market_data.xls")
Set y = Workbooks.Open(ALCOPath & "ALCOnewAlter.xlsm")

'Now, copy what you want from x:
x.Sheets(1).Range("A:G").Copy

'Now, paste to y worksheet:
y.Sheets("Rynek").Range("E:K").PasteSpecial Paste:=xlPasteValues

'Close x:
x.Application.CutCopyMode = False
x.Close False
xlApp.ScreenUpdating = True
xlApp.Visible = True
End Sub
Sub stawki()

    Dim xlApp As Excel.Application
    Dim wRates As Excel.Workbook
    
    Dim olApp As Outlook.Application
    Dim objMsg As Outlook.MailItem
    
    Dim objSelection As Outlook.Selection
    Dim sPath As String, sBody As String
    Dim nRow As Integer, nPoz As Integer
    Dim dEon As Double, dPol As Double
    
    Set xlApp = CreateObject("Excel.Application")
    Set olApp = CreateObject("Outlook.Application")
    
xlApp.ScreenUpdating = False
    
    sPath = ""
    xlApp.Visible = False
    Set wRates = xlApp.Workbooks.Open(sPath & "rates_history.xlsx")
    nRow = wRates.Sheets(1).Range("a2").End(xlDown).Row + 1
    If DATES.PoprzedniRoboczy(Date) <= wRates.Sheets(1).Cells(nRow - 1, 1).Value Then
        MsgBox "Stawki już zostały uzupełnione", vbExclamation
        wRates.Close False
        Exit Sub
    End If
    
    Set objSelection = olApp.ActiveExplorer.Selection
    For Each objMsg In objSelection
        sBody = objMsg.Body
        If Right(objMsg.Subject, 5) = "EONIA" Then
            nPoz = InStr(sBody, "wynoszaca ")
            If nPoz = 0 Then MsgBox "Błąd, zmienił się format maila. Zatrzymaj makro przez Ctrl+Pause, żeby to sprawdzić."
            dEon = Val(Mid(sBody, nPoz + 10))
            wRates.Sheets(2).Cells(nRow, 1).Value = DATES.PoprzedniRoboczy(Date)
            wRates.Sheets(2).Cells(nRow, 2).Value = dEon
        ElseIf Right(objMsg.Subject, 7) = "POLONIA" Then
            nPoz = InStr(sBody, "wynoszaca: ")
            If nPoz = 0 Then MsgBox "Błąd, zmienił się format maila. Zatrzymaj makro przez Ctrl+Pause, żeby to sprawdzić."
            dPol = Val(Mid(sBody, nPoz + 10))
            wRates.Sheets(1).Cells(nRow, 1).Value = DATES.PoprzedniRoboczy(Date)
            wRates.Sheets(1).Cells(nRow, 2).Value = dPol
        End If
    Next
    wRates.Close True
    
    MsgBox "Pomyślnie zapisano stawki!", vbInformation
xlApp.ScreenUpdating = True
End Sub

Sub krystian()
    
    Dim olApp As Outlook.Application
    Dim objMsg As Outlook.MailItem
    Set olApp = CreateObject("Outlook.Application")
    
    Dim sName As String
    sName = Sciezka_do_szablonow & "Krystian.oft"
    
    Set objMsg = olApp.CreateItemFromTemplate(sName)
    objMsg.Display
    
End Sub
Sub kontrola_dane()
    
    Dim objMsg As MailItem
    Dim sBody As String
    Dim data_raportu As Date
    
    data_raportu = Month_end(DateAdd("m", -1, Date))
    Call AppendFromTextFile(sBody, tmpPath, "html_kontrola_dane.txt")
    sBody = Replace(sBody, "@@@", Format(data_raportu, "yyyymm"))
    
    Set objMsg = Application.CreateItem(olMailItem)
    With objMsg
        .To = ""
        .Subject = "Kontrola danych do raportu ALM portfolio za " & MonthName(Month(data_raportu), False)
        .HTMLBody = sBody
        .Display
    End With
    
End Sub

Sub kontrola_finanse()
    
    Dim objMsg As MailItem
    Dim sBody As String
    Dim data_raportu As Date
    
    data_raportu = Month_end(DateAdd("m", -1, Date))
    Call AppendFromTextFile(sBody, tmpPath, "html_kontrola_finanse.txt")
    sBody = Replace(sBody, "@@@", Format(data_raportu, "yyyymm"))
    sBody = Replace(sBody, "@#@", Format(data_raportu, "dd.mm.yyyy"))
    
    Set objMsg = Application.CreateItem(olMailItem)
    With objMsg
        .To = ""
        .CC = ""
        .Subject = "Wyniki kontroli pozycji ALM ze sprawozdaniem DFS za " & MonthName(Month(data_raportu), False)
        .HTMLBody = sBody
        .Display
    End With
    
End Sub

Sub FX_position()
    Dim objAttachments As Outlook.Attachments

    Dim olApp As Outlook.Application
    Dim objMsg As Outlook.MailItem
    
    Dim objSelection As Outlook.Selection
    Set olApp = CreateObject("Outlook.Application")
    Set objSelection = olApp.ActiveExplorer.Selection

    For Each objMsg In objSelection
        objMsg.Attachments.Item(1).SaveAsFile (FXPath & objMsg.Attachments.Item(1).FileName)
    Next
    
End Sub
