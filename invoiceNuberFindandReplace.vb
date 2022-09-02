REM  *****  BASIC  *****
' Sets up Global variables
' Also when running this code always click inside of the Sub Main
' new test here.
Global currentDoc as Object
Global currentSheet as Object
Global outputCell as Object
Global invoiceCell as Object
Global inputCell as Object

Sub Main
	currentDoc = ThisComponent ' I get what doc your working in
	currentSheet = currentDoc.sheets(0) ' I get the first sheet in the doc

	Dim invoiceNumber as String
	Dim whatToReplace as String
	Dim newInvoiceNumber As String

	whatToReplace = "1884"

	updateInviceNumber(whatToReplace)
	
End Sub

Function updateInviceNumber(whatToReplace as String)
	Dim selectedCell as Object ' I get the current user selected cell
	Dim cellAddress as Object ' I get the current cell address
	Dim isCellEmpty as boolean ' I check if the cell is empty
	Dim counter as Integer ' I count how many times the loop runs
	
	selectedCell = ThisComponent.getCurrentSelection()
	cellAddress = selectedCell.CellAddress
	isCellEmpty = False
	counter = 0
	
	While isCellEmpty = False
		currentCell = currentSheet.getCellByPosition(cellAddress.column, cellAddress.row + counter) ' I get the current cell for the loop.
		If currentCell.Formula = "" Then
			isCellEmpty = True
		Else
			inputCell = currentSheet.getCellByPosition(cellAddress.column, cellAddress.row + counter) ' I get the input cell that has what needs to be changed.
			invoiceCell = currentSheet.getCellByPosition(0, cellAddress.row + counter) ' I get the invoice numbers from the first column of the sheet.
			outputCell = currentSheet.getCellByPosition(cellAddress.column, cellAddress.row + counter) ' I set the output cell for the new formula.
			invoiceNumber = invoiceCell.value
			newInvoiceNumber = Replace(inputCell.formula, whatToReplace, invoiceNumber) ' I find and replace the invoice number in the formula.
			outputCell.formula = newInvoiceNumber ' I set the output cell to have the updated formula.
			counter = counter + 1
		End If		
	Wend	
End Function
