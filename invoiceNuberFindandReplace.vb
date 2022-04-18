REM  *****  BASIC  *****
' Sets up Global var's
' Also when running this code always click inside of the Sub Main
Global currentDoc as Object
Global currentSheet as Object
Global outputCell as Object
Global invoiceCell as Object
Global inputCell as Object

Sub Main
	currentDoc = ThisComponent ' I get what doc your working in
	currentSheet = currentDoc.sheets(0) ' I get the first sheet in the doc

	Dim invoiceNumber as String
	Dim lastCellID as Integer
	Dim rowID as Integer
	Dim whatToReplace as String
	Dim newInvoiceNumber As String
	
	' tell me(lastCellID) where the last cell is
	lastCellID = 63
	' tell me(rowID) which row to look at
	rowID = 13
	' Tell me(whatToReplace) what to replace
	whatToReplace = "1884"
	
	' I search and replace the invoice numbers to be the right ones
	for i = 4 To lastCellID Step 1
		inputCell = currentSheet.getCellByPosition(rowID, i)
		invoiceCell = currentSheet.getCellByPosition(0, i)
		outputCell = currentSheet.getCellByPosition(rowID, i)

		invoiceNumber = invoiceCell.value
		newInvoiceNumber = Replace(inputCell.formula, whatToReplace, invoiceNumber)
		outputCell.formula = newInvoiceNumber
	Next
	
End Sub

'TODO: Not used and needs to be fixed so it only checks one collum
Function GetRange(cellName as Variant) as Integer
	Dim Cur as Object
	Dim Range as Object
	Cur = currentSheet.createCursorByRange(currentSheet.getCellRangeByName(cellName))
	Cur.gotoEndOfUsedArea(True)
	Range = currentSheet.getCellRangeByName(Cur.AbsoluteName)
	GetRange = Range.RangeAddress.EndRow
End Function