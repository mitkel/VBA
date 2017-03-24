Public Function CheckFileIsOpen(chkSumfile As String) As Boolean

    On Error Resume Next
    CheckFileIsOpen = (Workbooks(chkSumfile).Name = chkSumfile)
    On Error GoTo 0
    
End Function

Public Function CheckFileExists(chkSumfile As String) As Boolean

    Set fs = CreateObject("Scripting.FileSystemObject")   
    CheckFileExists = fs.fileexists(chkSumfile)
    
End Function

'....przypisuje danej zmiennej arkusz o zadanej ścieżce
Public Sub Assign_Wrkbk( _
    ByRef wbk As Variant, _
    ByVal sPath As String, _
    ByVal sName As String, _
    Optional ByVal bUpdate As Boolean = True, _
    Optional ByVal bReadOnly As Boolean = False)
    
    If User_functions.CheckFileIsOpen(sName) Then
        Set wbk = Workbooks(sName)
    Else
        Set wbk = Workbooks.Open(sPath & sName, UpdateLinks:=bUpdate, ReadOnly:=bReadOnly)
    End If

End Sub
Sub usun_funkcje()

Sub assign_wbk_dialogbox( _
    ByVal sPath As String, _
    ByRef wbk As Workbook)
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
    Application.FileDialog(msoFileDialogOpen).InitialFileName = sPath
    intChoice = Application.FileDialog(msoFileDialogOpen).Show
    
    If intChoice <> 0 Then
        plik = Application.FileDialog(msoFileDialogOpen).SelectedItems(1)
    End If
    
    Call User_functions.Assign_Wrkbk(wbk, Left(plik, InStrRev(plik, "\")), Right(plik, Len(plik) - InStrRev(plik, "\")))
End Sub

Sub save_all_workbooks()

    Dim w As Workbook
    
    For Each w In Workbooks
        w.Save
    Next

End Sub

Sub close_similar()
Application.EnableEvents = False
Application.ScreenUpdating = True

    Dim sName As String
    Dim bSave As Boolean
    
    
    bSave = ThisWorkbook.Sheets("PANEL").OLEObjects("zapisz_skoroszyt").Object.Value
    sName = ThisWorkbook.Sheets("PANEL").Range("c15").Value
    
    
    Dim w As Workbook
    For Each w In Workbooks
        If InStr(w.Name, sName) Then
            If bSave Then
                w.Save
            End If
            w.Close savechanges:=bSave
        End If
    Next

Application.ScreenUpdating = True
Application.EnableEvents = True
End Sub

' przypisuje zmienną do szablonu i zapisuje ja z nową nazwą
' path[h] - ścieżka do pliku zakończona "\"
' name[k] - nazwa pliku
Sub tmplate2report( _
    ByVal sPath1 As String, _
    ByVal sName1 As String, _
    ByVal sName2 As String, _
    ByRef wRap As Variant, _
    Optional ByVal sPath2 As String, _
    Optional ByVal bUpdate As Boolean = True, _
    Optional ByVal nFileFormat As XlFileFormat = xlOpenXMLWorkbook)
    
    Call User_functions.Assign_Wrkbk(wRap, sPath1, sName1, bUpdate)
    
    If sPath2 = "" Then
        wRap.SaveAs Filename:=sPath1 & sName2, FileFormat:=nFileFormat
    Else
        wRap.SaveAs Filename:=sPath2 & sName2, FileFormat:=nFileFormat
    End If
    
End Sub

Sub ExportToTextFile(FName As String, _
    Sep As String, SelectionOnly As Boolean, _
    AppendData As Boolean, _
    Optional ByVal decSep As String = ".")

Application.DecimalSeparator = decSep

	Dim WholeLine As String
	Dim FNum As Integer
	Dim RowNdx As Long
	Dim ColNdx As Integer
	Dim StartRow As Long
	Dim EndRow As Long
	Dim StartCol As Integer
	Dim EndCol As Integer
	Dim CellValue As String


Application.ScreenUpdating = False
	On Error GoTo EndMacro:
	FNum = FreeFile

	If SelectionOnly = True Then
	    With Selection
	        StartRow = .Cells(1).Row
	        StartCol = .Cells(1).Column
	        EndRow = .Cells(.Cells.Count).Row
	        EndCol = .Cells(.Cells.Count).Column
	    End With
	Else
	    With ActiveSheet.UsedRange
	        StartRow = .Cells(1).Row
	        StartCol = .Cells(1).Column
	        EndRow = .Cells(.Cells.Count).Row
	        EndCol = .Cells(.Cells.Count).Column
	    End With
	End If

	If AppendData = True Then
	    Open FName For Append Access Write As #FNum
	Else
	    Open FName For Output Access Write As #FNum
	End If

	For RowNdx = StartRow To EndRow
	    WholeLine = ""
	    For ColNdx = StartCol To EndCol
	        If Cells(RowNdx, ColNdx).Value = "" Then
	            CellValue = Chr(34) & Chr(34)
	        Else
	           CellValue = Cells(RowNdx, ColNdx).Text
	        End If
	        WholeLine = WholeLine & CellValue & Sep
	    Next ColNdx
	    WholeLine = Left(WholeLine, Len(WholeLine) - Len(Sep))
	    Print #FNum, WholeLine
	Next RowNdx

EndMacro:
	On Error GoTo 0
Application.ScreenUpdating = True
	Close #FNum

Application.DecimalSeparator = ","
End Sub

'dołacza do zadaniego strina tekst z zadanego pliku tekstowego
Sub AppendFromTextFile(
	ByRef sEnd As String, _
	ByVal sPath As String, _
	ByVal sName As String, _
	Optional ByVal sSep As String = "")

    Dim myFile As String
    Dim sLine As String
    
    myFile = sPath & sName
    Open myFile For Input As #1
    
    Do Until EOF(1)
        Line Input #1, sLine
        sEnd = sEnd & sLine & sSep
    Loop
    
    Close #1
End Sub