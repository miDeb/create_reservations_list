Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1
Sub reservierungsliste()
Dim blatt As Object
aktuel = ActiveSheet.Name
blattname = aktuel + "-Liste"
existiertblatt = False

For Each blatt In ThisWorkbook.Sheets
    If blatt.Name = blattname Then
        existiertblatt = True
        Exit For
    End If
Next blatt

If existiertblatt = False Then
    ThisWorkbook.Worksheets.Add.Name = aktuel + "-Liste"
Else
    MsgBox "Dieses Blatt gibt es schon"
End If

Worksheets(blattname).Select
Range("A1:E1000").Select
Selection.ClearContents
Worksheets(blattname).Cells(1, 1) = "Alphabetische Liste der Besucher"

rei = 2
spal = 1
For Spalte = 4 To 18
For reihe = 9 To 30

If Worksheets(aktuel).Cells(reihe, Spalte).Value <> "" Then
    Worksheets(blattname).Cells(rei, spal) = Worksheets(aktuel).Cells(reihe, Spalte)
    Worksheets(blattname).Cells(rei, spal + 1) = "Reihe - " + Format(reihe - 8, "00")
    Worksheets(blattname).Cells(rei, spal + 2) = "Nummer - " + Trim(Str(Spalte - 3))
    rei = rei + 1
End If

Next reihe
Next Spalte
'Stop
Worksheets(blattname).Select
Range("A1:D" + Trim(Str(rei - 1))).Select
Selection.Sort Key1:=Range("a1"), Order1:=xlAscending, Header:=xlYes, _
    OrderCustom:=1, MatchCase:=True, _
    Key2:=Range("b1"), Order1:=xlAscending, Header:=xlYes, _
    OrderCustom:=1, MatchCase:=True, Orientation:=xlTopToBottom, _
    DataOption1:=xlSortasnormal

namvorher = Worksheets(blattname).Cells(2, 1)
xx = 1
'Stop

For reihe = 3 To rei - 1
    namaktuell = Worksheets(blattname).Cells(reihe, 1)
    If namaktuell = namvorher Then
        Worksheets(blattname).Cells(reihe - 1, 4) = Str(xx)
        xx = xx + 1
        Worksheets(blattname).Cells(reihe, 4) = Str(xx)
    Else
        xx = 1
    End If
    namvorher = Worksheets(blattname).Cells(reihe, 1)
Next reihe
'Stop
For reihe = 2 To rei - 1
    If Worksheets(blattname).Cells(reihe, 4) <> 0 Then
        namaktuell = Worksheets(blattname).Cells(reihe, 1)
        Worksheets(blattname).Cells(reihe, 1) = namaktuell + " -" + Str(Worksheets(blattname).Cells(reihe, 4))
    End If
Next reihe
Range("D2:D1000").Select
Selection.ClearContents
Cells(1, 1).Select
End Sub






