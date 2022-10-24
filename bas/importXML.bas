REM  *****  BASIC  *****
Option Explicit


Global GemptyPlhdr As Integer	'Use empty placeholder #EMPTY# when importing a file
Private Const MNAME As String = "Import XML"

Private ColumnNames() As Variant
Private ColumnsOK As Boolean
Private nIO(3) As Long 'Number of used nodes for each non-empty IO node type
Private Rpos(3) As Long	'Current row position for writing
Private sheetsIO(3) As Object 'Sheets for each IO node type
Private oSheetObj 'Spreadheet object
Private XMLn As String 'XML file name without path, used for dialogs
Private IOtypes() As Variant 'Array ("DI", "AI", "DO", "AO")
Private oCheckBox(3) 'Dialog elements
Private oNameExistsLabel '"Name exists" label object
Private Overwrite As Boolean 'Should existing sheets be overwritten?
Private oOWCheckBox '"Overwrite" checkbox object
Private oOKbutt '"OK" button object
Private oSheetName 'Sheet name prefix
Private oDialogLib 'Dialog library


REM Returns True if sheet exists
Private Function SheetExists(sName As String) As Boolean
	SheetExists = oSheetObj.Sheets.hasByName(sName)
End Function

REM Advanced button click event
Function AdvancedEvent(oEvent)
	Dim oModule, oDlg, oHoffs, oemptyPl
	Dim val as Integer
	oModule = oDialogLib.getByName("ImportAdv")
	oDlg = CreateUnoDialog(oModule)
	oHoffs = oDlg.GetControl("OffsetHText")
	oHoffs.Text = GfirstCol
	oemptyPl = oDlg.GetControl("ReplaceEmptyCheckBox")
	oemptyPl.State = GemptyPlhdr
	DialogSetPosition(oDlg, oEvent.Source.Context, 1)
	If oDlg.execute = 0 Then
		Exit Function
	End If
	val = oHoffs.Text
	If val > 0 Then
		GfirstCol = val
	End If
	GemptyPlhdr = oemptyPl.State
End Function

REM Dialog element event handler
Sub UpdDialog(oEvent)
	Dim i As Integer
	Dim nCheck As Integer 'Number of selected nodes
	Dim sName As String 'IO sheet name
	Dim ShExists As Boolean
	For i = 0 To 3
		If oCheckBox(i).State = 1 Then nCheck = nCheck + 1 'Count the number of selected sheets
		sName = oSheetName.Text & "_" & IOtypes(i)
		If SheetExists(sName) Then
			ShExists = True
			oNameExistsLabel.Visible = True
			If oOWCheckBox.State = 0 Then oOKbutt.Enable = False
		End If
	Next i
	Select Case oEvent.source.getModel().Name
		Case "DICheckBox", "AICheckBox", "DOCheckBox", "AOCheckBox"
			nCheck = 0 'Count selected nodes
			For i = 0 To 3
				If oCheckBox(i).State = 1 Then nCheck = nCheck + 1
			Next i
			If nCheck < 1 Then
				oOKbutt.Enable = False 'None of IO nodes selected
'				MsgBox "None of IO nodes selected"
			Else
				If (NOT ShExists) Then
'					MsgBox "Sheet doesn't exist"
					oOKbutt.Enable = True 'One or more of IO nodes selected
				ElseIf (ShExists AND oOWCheckBox.State = 1) Then
'					MsgBox "Sheet exists AND oOWCheckBox.State=1"
					oOKbutt.Enable = True
				End If
			End If

		Case "SheetPrefix"
			oNameExistsLabel.Visible = False
			oOWCheckBox.Enable = False
			oOWCheckBox.State = 0
			If nCheck > 0 Then oOKbutt.Enable = True
			For i = 0 To 3
				sName = oSheetName.Text & "_" & IOtypes(i)
				oCheckBox(i).Label = sName & " - " & nIO(i) & " item(s)"
				If SheetExists(sName) Then
					ShExists = True
					oNameExistsLabel.Visible = True
					oOWCheckBox.Enable = True
					oOWCheckBox.State = 0
					oOKbutt.Enable = False
				End If
			Next i

		Case "OWCheckBox"
			Overwrite = oOWCheckBox.State = 1
			If nCheck > 0 Then oOKbutt.Enable = Overwrite
	End Select
	Overwrite = oOWCheckBox.State = 1
End Sub

REM Insert new name into Column array
Sub InsertCols(offs As Integer, ixIO As Integer, newName As String)
	Dim i as Integer
	Dim last as Integer
	Dim colNames as Variant
	colNames = ColumnNames(ixIO)
	last = UBound(colNames()) + 1
	ReDim Preserve colNames(last)

	if offs < last Then 'Shift elements towards the end of the array
		For i = last to offs + 1 Step -1
			colNames(i) = colNames(i - 1)
		Next i
	End If
	colNames(offs) = newName
	ColumnNames(ixIO) = colNames
End Sub

REM Check if there is an attribute name that needs to be added to the Column array
Sub UpdateCols(attrList() As String, ixIO As Integer)
	Dim i As Integer, a As Integer
	Dim offset As Integer 'Index of the attribute name in array, -1 if not found
	Dim insAt As Integer 'Offset to insert a new element in Column array

	For a = 0 to UBound(attrList()) 'Check if attribute name exists in Column array
		offset = IndexInArray(attrList(a), ColumnNames(ixIO))
		If offset < 0 Then 'Attribute not found
			insAt = UBound(ColumnNames(ixIO)) + 1 'Set initial insert offset beyond the end of the array
			For i = a + 1 to UBound(attrList()) 'Check if the next attribute name exists in array
				offset = IndexInArray(attrList(i), ColumnNames(ixIO)) 'Set the insert offset before next found attribute
				If offset >= 0 Then 'Attribute found
					insAt = offset
					Exit For
				End If
			Next i
			InsertCols(insAt, ixIO, attrList(a)) 'Get space for the new element
		End If
	Next a
End Sub

Sub CreateCols(cName As String, oAttList As Object)
	Dim i As Integer
	Dim ixIO As Integer

	ixIO = IndexInArray(cName, IOtypes())
	If ixIO < 0 Then Exit Sub
	If oAttList.Length > 0 Then
		Dim attrList(oAttList.Length - 1) As String
		nIO(ixIO) = nIO(ixIO) + 1
		For i = 0 to oAttList.Length - 1
			attrList(i) = oAttList.getNameByIndex(i)
		Next i
		If UBound(ColumnNames(ixIO)) < 0 Then
			ColumnNames(ixIO) = attrList
		Else
			UpdateCols(attrList(), ixIO)
		End If
	End If
End Sub

REM Check if value is Integer and copy to Cell
Sub CopyValue(vValue As Variant, oCell as Object)
	Dim fval As Double
	On Local Error Goto copyString

	fval = CDbl(vValue)	'Convert to Double, On Error handles overflow
	If fval = vValue Then
		oCell.Value = vValue
	Else
		copyString:
		oCell.String = vValue
	End If
End Sub

REM Populate a sheet row with IO node attribute values
Sub GetValues(sNode As String, oAttList As Object)
	Dim nAttrib As Integer 'Number of attributes of sNode
	Dim i As Integer, c As Integer
	Dim ixIO As Integer

	ixIO = IndexInArray(sNode, IOtypes())
	If ixIO < 0 Then Exit Sub
	If IsNull(sheetsIO(ixIO)) Then Exit Sub

	nAttrib = oAttList.getLength()
	Dim Row(UBound(ColumnNames(ixIO)))

	For i = 0 to nAttrib - 1
		c = IndexInArray(oAttList.getNameByIndex(i), ColumnNames(ixIO))
		If c >= 0 Then
			Row(c) = oAttList.getValueByIndex(i)
			If Row(c) = "" Then
				If GemptyPlhdr > 0 Then
					Row(c) = EMPTYKEY
				End If
			End If
		Else 'Attribute name not found in Columns array, this is not supposed to happen
			MsgBox "Attribute: " & oAttList.getNameByIndex(i) & " is not found in Columns()", MB_ICONSTOP, MNAME
		End If
	Next i
	For c = 0 To UBound(Row())
		If Not IsEmpty(Row(c)) Then
			CopyValue(Row(c), sheetsIO(ixIO).getCellByPosition(c + GfirstCol - 1, Rpos(ixIO)))
		End If
	Next c
	Rpos(ixIO) = Rpos(ixIO) + 1
End Sub

Sub ReadXmlFromInputStream(oInputStream)
	Dim oSaxParser As Object
	Dim oDocEventsHandler As Object
	Dim oInputSource As Object
	' Create a Sax Xml parser.
	oSaxParser = createUnoService("com.sun.star.xml.sax.Parser")
	' Create a document event handler object.
	' As methods of this object are called, Basic arranges
	' for global routines (see below) to be called.
	oDocEventsHandler = CreateDocumentHandler()
	' Plug our event handler into the parser.
	' As the parser reads an Xml document, it calls methods
	' of the object, and hence global subroutines below
	' to notify them of what it is seeing within the Xml document.
	oSaxParser.setDocumentHandler(oDocEventsHandler)
	' Create an InputSource structure.
	oInputSource = createUnoStruct("com.sun.star.xml.sax.InputSource")
	oInputSource.aInputStream = oInputStream ' plug in the input stream

	' Now parse the document.
	' This reads in the entire document.
	' Methods of the oDocEventsHandler object are called as
	' the document is scanned.
	oSaxParser.parseStream(oInputSource)
End Sub

Sub ReadXmlFromUrl(fName)
	Dim oInputStream As Object
	oInputStream = oSimpleFileAccess.openFileRead(fName)
	ReadXmlFromInputStream(oInputStream)
	oInputStream.closeInput()
End Sub

'==================================================
'  Xml Sax document handler.
'==================================================
' Global variables used by our document handler.
'
' Once the Sax parser has given us a document locator,
' the glLocatorSet variable is set to True,
' and the goLocator contains the locator object.
'
' The methods of the locator object has cool methods
' which can tell you where within the current Xml document
' being parsed that the current Sax event occured.
' The locator object implements com.sun.star.xml.sax.XLocator.
'
Private goLocator As Object
Private glLocatorSet As Boolean
	' This creates an object which implements the interface
	' com.sun.star.xml.sax.XDocumentHandler.
	' The doucment handler is returned as the function result.
Function CreateDocumentHandler()
	' Use the CreateUnoListener function of Basic.
	' Basic creates and returns an object that implements a particular interface.
	' When methods of that object are called,
	' Basic will call global Basic functions whose names are the same
	' as the methods, but prefixed with a certian prefix.
	Dim oDocHandler As Object
	oDocHandler = CreateUnoListener("DocHandler_", "com.sun.star.xml.sax.XDocumentHandler")
	glLocatorSet = False
	CreateDocumentHandler() = oDocHandler
End Function

'==================================================
'  Methods of our document handler call these
'  global functions.
'  These methods look strangely similar to
'  a SAX event handler. ;-)
'  These global routines are called by the Sax parser
'  as it reads in an XML document.
'  These subroutines must be named with a prefix that is
'  followed by the event name of the com.sun.star.xml.sax.XDocumentHandler interface.
'==================================================
Function DocHandler_startDocument()
End Function

Function DocHandler_endDocument()
	ColumnsOK = True
End Function

Function DocHandler_startElement(sName As String, oAttList As com.sun.star.xml.sax.XAttributeList)
	If ColumnsOK Then
		GetValues(sName, oAttList)
	Else
		CreateCols(sName, oAttList)
	End If
End Function

Function DocHandler_endElement(sName As String)
End Function

Function DocHandler_characters(sChars As String)
End Function

Function DocHandler_ignorableWhitespace(sWhitespace As String)
End Function

Function DocHandler_processingInstruction(cTarget As String, cData As String)
End Function

Function DocHandler_setDocumentLocator(oLocator As com.sun.star.xml.sax.XLocator)
	' Save the locator object in a global variable.
	' The locator object has valuable methods that we can
	' call to determine
	goLocator = oLocator
	glLocatorSet = True
End Function

Sub Main
	Dim fName As String 'XML file name
	Dim i As Integer, t As Integer
	Dim oModule, oDlg 'Dialog objects
'	Dim oSheetName 'Sheet name prefix
	Dim oOptionButton1, oOptionButton2
	Dim oCell 'Cell object
	Dim oRange
	Dim clear As Boolean

	GlobalInit()
	IOtypes = Array("DI", "AI", "DO", "AO")
	ColumnNames = Array(Array(), Array(), Array(), Array())
	oSheetObj = ThisComponent
	If GhomeURL = "" Then ResolveFilepath("")

	fName = OpenFile()
	If fName = "" Then
'		MsgBox "No file selected", MB_ICONINFORMATION, MNAME
		Exit Sub
	End If

	XMLn = FileNameoutofPath(fName)
	ReadXmlFromUrl(fName) 'First run to get column headers (attributes of the IO nodes)

	DialogLibraries.LoadLibrary("XML_lib")
	oDialogLib = DialogLibraries.getByName("XML_lib")
	oModule = oDialogLib.getByName("ImportDialog")
	oDlg = CreateUnoDialog(oModule)
	oSheetName = oDlg.GetControl("SheetPrefix")
	oCheckBox(0) = oDlg.GetControl("DICheckBox")
	oCheckBox(1) = oDlg.GetControl("AICheckBox")
	oCheckBox(2) = oDlg.GetControl("DOCheckBox")
	oCheckBox(3) = oDlg.GetControl("AOCheckBox")
	oNameExistsLabel = oDlg.GetControl("NameExists")
	oOWCheckBox = oDlg.GetControl("OWCheckBox")
	oOKbutt = oDlg.GetControl("ImportButt")
	oDlg.Title = XMLn & " - " & MNAME
	DialogSetPosition(oDlg, ThisComponent.getCurrentController().Frame.getComponentWindow(), 0)
	oSheetName.Text = GetFileNameWithoutExtension(XMLn)
	oNameExistsLabel.Visible = False
	oOWCheckBox.Enable = False
	oOKbutt.Enable = True
	For t = 0 to 3
		REM Update labels of checkboxes according the sheet name prefix (default -- XML file name wo extension)
		oCheckBox(t).Label = oSheetName.Text & "_" & IOtypes(t) & " - " & nIO(t) & " item(s)"
		If nIO(t) > 0 Then 'Set checkboxes for non-empty IO nodes
			oCheckBox(t).State = 1
		Else
			oCheckBox(t).Enable = False
		End If
		If SheetExists(oSheetName.Text & "_" & IOtypes(t)) Then
			oNameExistsLabel.Visible = True
			oOWCheckBox.Enable = True
			oOWCheckBox.State = 0
			oOKbutt.Enable = False
		End If
	Next t

	If oDlg.execute = 0 Then
'		MsgBox "Operation canceled", MB_ICONINFORMATION, MNAME
		Exit Sub
	End If

	For t = 0 To 3
		If oCheckBox(t).State = 1 Then
			Rpos(t) = 1
			If SheetExists(oSheetName.Text & "_" & IOtypes(t)) Then
				clear = True
			Else
				oSheetObj.Sheets.insertNewByName(oSheetName.Text & "_" & IOtypes(t), oSheetObj.Sheets.Count + 1)  'Add new sheet at the end of the document
				clear = False
			End If
			sheetsIO(t) = oSheetObj.Sheets.getByName(oSheetName.Text & "_" & IOtypes(t))

			If clear Then  'Clear all cell contents
				sheetsIO(t).ClearContents(511)	'com.sun.star.sheet.CellFlags
			End If

			For i = 0 To UBound(ColumnNames(t))
				oCell = sheetsIO(t).getCellByPosition(i + GfirstCol - 1, 0) 'Start from B1
				oCell.String = ColumnNames(t)(i)
				'oCell.CharWeight = com.sun.star.awt.FontWeight.BOLD
			Next i
		End If
	Next t

	oSheetObj.lockControllers()
	'oSheetObj.enableAutomaticCalculation(False)
	ReadXmlFromUrl(fName) 'Second run to fill column data (values of the IO node attributes)
	'Set optimal width of the columns
	For t = 3 To 0 Step -1 'Process IO sheets in reverse direction to finish on the first used sheet
		If Not IsNull(sheetsIO(t)) Then
			oRange = sheetsIO(t).getCellRangeByPosition(1, 0, UBound(ColumnNames(t)) + 1, 0)
			oRange.Columns.OptimalWidth = True
			oSheetObj.CurrentController.ActiveSheet = sheetsIO(t)
			oSheetObj.CurrentController.freezeAtPosition(0, 1) 'Freeze the first row
		End If
	Next t
	oSheetObj.unlockControllers()
	'oSheetObj.enableAutomaticCalculation(True)
	MsgBox "Import complete!" & CHR$(13) & ConvertFromURL(fName), MB_ICONINFORMATION, MNAME
End Sub

