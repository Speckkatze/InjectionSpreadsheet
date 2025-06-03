Sub AddNewVial
 	Dim oDoc As Object, oSheet As Object
    oDoc = ThisComponent
    oSheet = oDoc.CurrentController.ActiveSheet

    Dim row As Long
    row = 1

    Do While row < 1000 ' if you think you need more than a thousand vials I really think you should talk to some sort of specialist, even if its one for garbage disposal
        Dim isEmpty As Boolean
        isEmpty = True
        For col = 15 To 22
            If oSheet.getCellByPosition(col, row).Type <> com.sun.star.table.CellContentType.EMPTY Then
                isEmpty = False
                Exit For
            End If
        Next
        If isEmpty Then Exit Do
        row = row + 1
    Loop

    ' Load and show the dialog
    DialogLibraries.LoadLibrary("Standard")
    Dim oDialog As Object
    oDialog = CreateUnoDialog(DialogLibraries.Standard.VialInputDialog)

    If oDialog.execute() = 1 Then
        Dim txtCode, txtDensity, txtVolume, txtDate, txtConcentration, txtAPI, txtSource As String
        txtCode = oDialog.getControl("txtCode").Text
        txtDensity = oDialog.getControl("txtDensity").Text
        txtVolume = oDialog.getControl("txtVolume").Text
        txtDate = oDialog.getControl("txtDate").Text
        txtConcentration = oDialog.getControl("txtConcentration").Text
        txtAPI = oDialog.getControl("txtAPI").Text
        txtSource = oDialog.getControl("txtSource").Text

        ' Fill columns B, C, D
        oSheet.getCellByPosition(15, row).String = txtCode ' Column P
        oSheet.getCellByPosition(16, row).Value = txtDensity ' Column Q
        oSheet.getCellByPosition(17, row).Value = txtVolume ' Column R
        oSheet.getCellByPosition(18, row).String = txtDate ' Column S
        oSheet.getCellByPosition(19, row).Value = txtConcentration ' Column T
        oSheet.getCellByPosition(20, row).String = txtAPI ' Column U
        oSheet.getCellByPosition(21, row).String = txtSource ' Column V
        
        Dim r As String : r = (row + 1)       
        
        oSheet.getCellByPosition(22, row).Formula = "=SUMIF(D:D; P" & (row + 1) & "; O:O)"

    End If

    oDialog.dispose()

End Sub




Sub AutoFillNextEmptyRow
    Dim oDoc As Object, oSheet As Object
    oDoc = ThisComponent
    oSheet = oDoc.CurrentController.ActiveSheet

    Dim row As Long
    row = 1

    ' Find the next row where A to O are all empty
    Do While row < 1000 ' if you feel like you will live longer than 200 years with weekly injections, increase this
        Dim isEmpty As Boolean
        isEmpty = True
        For col = 0 To 14 ' Columns A to O
            If oSheet.getCellByPosition(col, row).Type <> com.sun.star.table.CellContentType.EMPTY Then
                isEmpty = False
                Exit For
            End If
        Next
        If isEmpty Then Exit Do
        row = row + 1
    Loop

    ' Load and show the dialog
    DialogLibraries.LoadLibrary("Standard")
    Dim oDialog As Object
    oDialog = CreateUnoDialog(DialogLibraries.Standard.RowInputDialog)

    If oDialog.execute() = 1 Then
        Dim txtNumber, txtDate, txtTime, txtCode, txtDrawn, txtVialBefore, txtVialAfter, txtSNMassBefore, txtSNMassDrawn, txtSNMassAfter As String
        txtNumber = oDialog.getControl("txtNumber").Text
        txtDate = oDialog.getControl("txtDate").Text
        txtTime = oDialog.getControl("txtTime").Text
        txtCode = oDialog.getControl("txtCode").Text
        txtDrawn = oDialog.getControl("txtDrawn").Text
        txtVialBefore = oDialog.getControl("txtVialBefore").Text
        txtVialAfter = oDialog.getControl("txtVialAfter").Text
        txtSNMassBefore = oDialog.getControl("txtSNMassBefore").Text
        txtSNMassDrawn = oDialog.getControl("txtSNMassDrawn").Text
        txtSNMassAfter = oDialog.getControl("txtSNMassAfter").Text

        ' Fill columns B, C, D
        oSheet.getCellByPosition(0, row).Value = txtNumber ' Column A
        oSheet.getCellByPosition(1, row).String = txtDate ' Column B
        oSheet.getCellByPosition(2, row).String = txtTime ' Column C
        oSheet.getCellByPosition(3, row).String = txtCode ' Column D
        oSheet.getCellByPosition(4, row).Value = txtVialBefore ' Column E
        oSheet.getCellByPosition(5, row).Value = txtVialAfter ' Column F
        oSheet.getCellByPosition(6, row).Value = txtDrawn ' Column G
        oSheet.getCellByPosition(7, row).Value = txtSNMassBefore ' Column H
        oSheet.getCellByPosition(8, row).Value = txtSNMassDrawn ' Column I
        oSheet.getCellByPosition(9, row).Value = txtSNMassAfter ' Column J

        ' Fill formulas in K to O
        'oSheet.getCellByPosition(10, row).Formula = "=IF((I" & (row+1) & " - J" & (row+1) & ") * IFNA(VLOOKUP(D" & (row+1) & "; P:Q; 2; 0); 0) * 0,001 = 0; ""; (I" & (row+1) & " - J" & (row+1) & ") * IFNA(VLOOKUP(D" & (row+1) & "; P:Q; 2; 0); 0) * 0,001)"
        'oSheet.getCellByPosition(11, row).Formula = "=IF((I" & (row+1) & " - J" & (row+1) & ") * IFNA(VLOOKUP(D" & (row+1) & "; P:Q; 2; 0); 0) * 0,001 * IFNA(VLOOKUP(D" & (row+1) & "; P:T; 5; 0); 0) = 0; ""; (I" & (row+1) & " - J" & (row+1) & ") * IFNA(VLOOKUP(D" & (row+1) & "; P:Q; 2; 0); 0) * 0,001 * IFNA(VLOOKUP(D" & (row+1) & "; P:T; 5; 0); 0))"
        'oSheet.getCellByPosition(12, row).Formula = "=IF((J" & (row+1) & " - H" & (row+1) & ") * IFNA(VLOOKUP(D" & (row+1) & "; P:Q; 2; 0); 0) * 0,001 = 0; ""; (J" & (row+1) & " - H" & (row+1) & ") * IFNA(VLOOKUP(D" & (row+1) & "; P:Q; 2; 0); 0) * 0,001)"
        'oSheet.getCellByPosition(13, row).Formula = "=IF((J" & (row+1) & " - H" & (row+1) & ") * IFNA(VLOOKUP(D" & (row+1) & "; P:Q; 2; 0); 0) * 0,001 * IFNA(VLOOKUP(D" & (row+1) & "; P:T; 5; 0); 0) = 0; ""; (J" & (row+1) & " - H" & (row+1) & ") * IFNA(VLOOKUP(D" & (row+1) & "; P:Q; 2; 0); 0) * 0,001 * IFNA(VLOOKUP(D" & (row+1) & "; P:T; 5; 0); 0))"
        'oSheet.getCellByPosition(14, row).Formula = "=IF(N(M" & (row+1) & ") + N(K" & (row+1) & ") = 0; ""; N(M" & (row+1) & ") + N(K" & (row+1) & "))"
        
        Dim r As String : r = (row + 1)

		oSheet.getCellByPosition(10, row).Formula = "=IF((I" & r & " - J" & r & ") * VLOOKUP(D" & r & "; P:Q; 2; 0) * 0.001 = 0; """ & """; (I" & r & " - J" & r & ") * VLOOKUP(D" & r & "; P:Q; 2; 0) * 0.001)"
		
		oSheet.getCellByPosition(11, row).Formula = "=IF((I" & r & " - J" & r & ") * VLOOKUP(D" & r & "; P:Q; 2; 0) * 0.001 * VLOOKUP(D" & r & "; P:T; 5; 0) = 0; """ & """; (I" & r & " - J" & r & ") * VLOOKUP(D" & r & "; P:Q; 2; 0) * 0.001 * VLOOKUP(D" & r & "; P:T; 5; 0))"
		
		oSheet.getCellByPosition(12, row).Formula = "=IF((J" & r & " - H" & r & ") * VLOOKUP(D" & r & "; P:Q; 2; 0) * 0.001 = 0; """ & """; (J" & r & " - H" & r & ") * VLOOKUP(D" & r & "; P:Q; 2; 0) * 0.001)"
		
		oSheet.getCellByPosition(13, row).Formula = "=IF((J" & r & " - H" & r & ") * VLOOKUP(D" & r & "; P:Q; 2; 0) * 0.001 * VLOOKUP(D" & r & "; P:T; 5; 0) = 0; """ & """; (J" & r & " - H" & r & ") * VLOOKUP(D" & r & "; P:Q; 2; 0) * 0.001 * VLOOKUP(D" & r & "; P:T; 5; 0))"
		
		oSheet.getCellByPosition(14, row).Formula = "=IF(N(M" & r & ") + N(K" & r & ") = 0; """ & """; N(M" & r & ") + N(K" & r & "))"
    End If

    oDialog.dispose()
End Sub
