REM  *****  BASIC  *****
Option Explicit

Private Const MNAME As String = "Export XML"
Private Const CTABLE As String = "Table"
Private Const VERSIONCONTROL As String = "VersionControl"
Private Const CCONF As String = "conf"
Private Const CDATE As String = "date"
Private Const CTIME As String = "time"
Private Const GDEBUG As Boolean = False
Private HtmlEntities As Variant
Private IOtypes() As Variant	' Array ("DI", "AI", "DO", "AO")
Private oDialogLib As Object	' Dialog library
Private oOptionButton1, oOptionButton2	' Radio buttons
Private oFileName1, oFileName2	' XML file text box and label
Private oNameExistsLabel		'"Name exists" label object
Private oOWCheckBox				'"Overwrite" checkbox object
Private oOKbutt					'"OK" button object
Private sOutputFile As String	' filename
Private NScontainer() As Variant	' Namespace definitions found in document

REM Advanced button click event
Function AdvancedEvent(oEvent)
	Dim oModule, oDlg, oHoffs
	Dim val as Integer
	oModule = oDialogLib.getByName("ExportAdv")
	oDlg = CreateUnoDialog(oModule)
	oHoffs = oDlg.GetControl("OffsetHText")
	oHoffs.Text = GfirstCol
	DialogSetPosition(oDlg, oEvent.Source.Context, 1)
	If oDlg.execute = 0 Then
		Exit Function
	End If
	val = oHoffs.Text
	If val > 0 Then
		GfirstCol = val
	End If
End Function

REM File name (text) modified event
Function NameModifyEvent(oEvent)
	Dim sPath As String
	If oOptionButton2.State Then
		sPath = ResolveFilepath(oFileName2.Text)
		If oSimpleFileAccess.exists(sPath) Then
			oNameExistsLabel.Visible = True
			oOWCheckBox.Enable = True
			If oOWCheckBox.State = 0 Then oOKbutt.Enable = False
			Exit Function
		End If
	End If
	oNameExistsLabel.Visible = False
	oOWCheckBox.Enable = False
	oOKbutt.Enable = True
End Function

REM Overwrite checkbox event
Function OverwriteCbEvent(oEvent)
	If oOWCheckBox.State = 0 Then
		NameModifyEvent(Nothing)
	Else
		oOKbutt.Enable = True
	End If
End Function

REM Browse file button click event
Function BrowseEvent(oEvent)
	Dim fName As String 'XML file name
	fName = OpenFile()
	If fName <> "" Then
		oOptionButton2.State = True
		oFileName2.Text = ConvertFromUrl(fName)
	End If
End Function

REM Decode UTF-8 character
Function DecodeUTF8(sBuff As String, utfVal As Long, bData As Integer, charnum As Integer) As Boolean
	Dim b As Integer
	If charnum > 0 Then	' Parsing an UTF-8 character
		b = 256 + bData
		If (b AND 192) = 128 Then
			utfVal = utfVal * 64
			utfVal = utfVal + (b AND 63)
			charnum = charnum - 1
		Else
			sBuff = "Invalid byte (0x" & Hex$(b) & ") of the UTF-8 character in:"
			Goto ErrorHandler
		End If
		If charnum = 0 Then
			sBuff = sBuff & CHR$(utfVal)
		End If
	Else
		If bData > 0 Then ' ASCII character
			sBuff = sBuff & CHR$(bData)
		Else ' UTF-8 character
			b = 256 + bData
			If (b AND 224) = 192 Then
				utfVal = (b AND 31)
				charnum = 1
			ElseIf (b AND 240) = 224 Then
				utfVal = (b AND 15)
				charnum = 2
			ElseIf (b AND 248) = 240 Then
				utfVal = (b AND 7)
				charnum = 3
			Else
				sBuff = "Invalid first byte (0x" & Hex$(b) & ") of the UTF-8 character in:"
				Goto ErrorHandler
			End If
		End If
	End If

	Exit Function
	ErrorHandler:
	MsgBox sBuff & CHR$(13) & sOutputFile, MB_ICONSTOP, MNAME
	DecodeUTF8 = True
End Function

REM Update configuration version if XML file has VersionControl node
Function ConfigVersion(oRoot As Object) As Boolean
	Dim oNodeList As Object
	Dim oAttr As Object
	Dim fVersNum As Single
	Dim sValue As String
	On Error Goto ErrorHandler

	oNodeList = oRoot.getElementsByTagName(VERSIONCONTROL)
	If IsNull(oNodeList) Then Exit Function
	Select Case (oNodeList.Length)
	Case 0
	Case 1
		oAttr = oNodeList.item(0).getAttributeNode(CCONF)
		If Not IsNull(oAttr) Then
			fVersNum = oAttr.getValue()
			If (fVersNum > 0 AND fVersNum < 9999) Then
				fVersNum = fVersNum + 1
				sValue = Format(fVersNum, "0.00")
				sValue = join(split(sValue, ","), ".")	' Make sure there is no ','
				oAttr.setValue(sValue)
			End If
		End If
		oAttr = oNodeList.item(0).getAttributeNode(CDATE)
		If Not IsNull(oAttr) Then
			oAttr.setValue(Format(Date, "yyyy-mm-dd"))
		End If
		oAttr = oNodeList.item(0).getAttributeNode(CTIME)
		If Not IsNull(oAttr) Then
			oAttr.setValue(Format(Time, "HH:MM:SS"))
		End If
	Case Else
		MsgBox oNodeList.Length & " <" & VERSIONCONTROL & "> nodes found instead of 1 in file:" & CHR$(13) & sOutputFile, MB_ICONSTOP, MNAME
		Goto AfterError
	End Select

	Exit Function
	ErrorHandler:
	MsgBox Error$ & CHR$(13) & sOutputFile, MB_ICONSTOP, MNAME
	AfterError:
	ConfigVersion = True
End Function

REM Get all namespaces defined in the document by checking 'xmlns' attributes
Function GetNamespaces() As Boolean
	Dim i As Long, size As Long, utfVal As Long
	Dim oInputStream As Object
	Dim bData() As Byte
	Dim sBuff As String, prefix As String, sName As String, sPropVal As String
	Dim state As Integer, ucharnum As Integer

	oInputStream = oSimpleFileAccess.openFileRead(sOutputFile)
	size = oInputStream.Length
	For i = 0 To size - 1
		oInputStream.readBytes(bData, 1)
		Select Case (state)
		Case 0
			If bData(0) = 60 Then	' Less than '<'
				state = state + 1
			End If
		Case 1
			Select Case (bData(0))
			Case 62					' Greater than '>'
				state = 0
			Case 33					' Exclamation mark '!'
				state = 3
			Case 9,10,13,32,47,63	' Ignore TAB,LF,CR,SPACE, '/', '?'
				state = state + 1
			Case 34,38,39,60,61		' Can't have double qoute, '&', single qoute, '<', '='
				sBuff = "A-Z"
				Goto InvalidChar
			Case Else
				StartName:
				ucharnum = 0
				sName = ""
				state = 8
				Goto ParseName
			End Select
		Case 2
			Select Case (bData(0))
			Case 62					' Greater than '>'
				state = 0
			Case 9,10,13,32			' Ignore TAB,LF,CR,SPACE
			Case 33,34,38,39,47,60,61,63 ' Can't have '!', double qoute, '&', single qoute, '/', '<', '=', '?'
				sBuff = "A-Z"
				Goto InvalidChar
			Case Else
				Goto StartName
			End Select
		Case 3, 4
			If bData(0) = 45 Then	' Minus '-'
				state = state + 1
			Else
			 	sBuff = "-"
				Goto InvalidChar
			End If
		Case 5
			If bData(0) = 45 Then state = state + 1	' Minus '-'
		Case 6
			If bData(0) = 45 Then 	' Minus '-'
				state = state = state + 1
			Else
				state = state = state - 1
			End If
		Case 7
			If bData(0) = 62 Then	' Greater than '>'
				state = 0
			Else
				sBuff = ">"
				InvalidChar:
				If bData(0) > 0 Then
			 		sBuff = "Invalid symbol '" & CHR$(bData(0)) & "' found, expected '" & sBuff & "' in:"
			 	Else
			 		sBuff =	"Invalid UTF-8 byte (0x" & Hex$(256 + bData(0)) & ") found, expected '" & sBuff & "' in:"
			 	End If
				MsgBox sBuff & CHR$(13) & sOutputFile, MB_ICONSTOP, MNAME
				GetNamespaces = True
				Goto nscomplete
			End If
		Case 8
			ParseName:
			Select Case (bData(0))
			Case 9,10,13,32			' TAB,LF,CR,WHITESPACE means end of name and looking for '='
				state = state + 1
			Case 61					' '=' means end of name
				state = 10
			Case 62					' '>'
				state = 0
			Case 47,63				' '/', '?'
				state = 7
			Case 33,34,38,39,60		' Can't have '!', double qoute, '&', single qoute, '<'
				sBuff = "="
				Goto InvalidChar
			Case Else
				If DecodeUTF8(sName, utfVal, bData(0), ucharnum) Then Goto nscomplete
			End Select
		Case 9
			Select Case (bData(0))
			Case 61					' '=' attribute value will follow
				state = state + 1
			Case 34					' '"' means attribute value starts now
				state = 11
			Case 62					' '>'
				state = 0
			Case 47,63				' '/', '?'
				state = 7
			Case 9,10,13,32			' Ignore TAB,LF,CR,SPACE
			Case 33,38,39,60		' Can't have '!', '&', single qoute, '<'
				sBuff = """"
				Goto InvalidChar
			Case Else
				Goto StartName
			End Select
		Case 10
			Select Case (bData(0))
			Case 34					' '"' means attribute value starts now
				ucharnum = 0
				sPropVal = ""
				state = state + 1
			Case 9,10,13,32			' Ignore TAB,LF,CR,SPACE
			Case Else
			 	sBuff = """"
				Goto InvalidChar
			End Select
		Case 11
			Select Case (bData(0))
			Case 34					' '"' means attribute value ends now
				If Left(sName, 5) = "xmlns" Then
					If Mid(sName, 6, 1) = ":" Then
						prefix = Right(sName, Len(sName) - 6)
					Else
						prefix = ""
					End If
					Dim oldsize As Integer
					oldsize = UBound(NScontainer)
					ReDim Preserve NScontainer(oldsize + 1)
					NScontainer(oldsize + 1) = Array(prefix, sPropVal)
				End If
				state = state + 1
			Case Else
				If DecodeUTF8(sPropVal, utfVal, bData(0), ucharnum) Then Goto nscomplete
			End Select
		Case 12
			Select Case (bData(0))
			Case 62					' Greater than '>'
				If UBound(NScontainer) >= 0 Then Goto nscomplete
				state = 0
			Case 47,63				' '/', '?'
				state = 7
			Case 9,10,13,32			' Ignore TAB,LF,CR,SPACE
			Case 33,34,38,39,60,61	' Can't have '!', double qoute, '&', single qoute, '<', '='
				sBuff = ">"
				Goto InvalidChar
			Case Else
				Goto StartName
			End Select
		End Select
	Next i
	nscomplete:
	oInputStream.closeInput()
End Function

REM Create a new XML file
Function NewXML(oDocBuilder As Object, oDocument As Object, tabNodes() As Object, sComment As String) As Boolean
	Dim t As Integer
	Dim oRoot As Object
	Dim oComment As Object
	On Error Goto ErrorHandler

	oDocument = oDocBuilder.newDocument()
	sComment = "XML generated from sheets:" & sComment
	oComment = oDocument.createComment(sComment)
	oDocument.appendChild(oComment)
	oRoot = oDocument.createElement("objects")
	oDocument.appendChild(oRoot)
	For t = 0 to 3	' Create nodes for each selected sheet
		tabNodes(t) = oDocument.createElement(IOtypes(t) & CTABLE)
		oRoot.appendChild(tabNodes(t))
	Next t

	Exit Function
	ErrorHandler:
	MsgBox Error$ & CHR$(13) & sOutputFile, MB_ICONSTOP, MNAME
	AfterError:
	NewXML = True
End Function

REM Parse XML file and remove existing <DI>, <AI>, <DO>, <AO> nodes
Function ParseXML(oDocBuilder As Object, oDocument As Object, tabNodes() As Object) As Boolean
	Dim t As Integer, i As Integer
	Dim oRoot As Object
	Dim oNodeList As Object
	On Error Goto ErrorHandler

	oDocument = oDocBuilder.parseURI(sOutputFile)
	If IsNull(oDocument) Then Goto AfterError
	oRoot = oDocument.getDocumentElement()
	If ConfigVersion(oRoot) Then Goto AfterError
	If oRoot.NamespaceURI <> "" Then	' Look for other namespaces only if main namespace defined
		If GetNamespaces() Then Goto AfterError
	End If

	For t = 0 to 3	' Check existing Table nodes
		oNodeList = oRoot.getElementsByTagName(IOtypes(t) & CTABLE)
		If Not IsNull(oNodeList) Then
			Select Case (oNodeList.Length)
			Case 0
			Case 1
				tabNodes(t) = oNodeList.item(0)
				oNodeList = tabNodes(t).getChildNodes()
				Do While oNodeList.Length > 0	' Delete existing <DI>, <AI>, <DO>, <AO> nodes
					tabNodes(t).removeChild(oNodeList.item(0))
				Loop
			Case Else
				MsgBox oNodeList.Length & " <" & IOtypes(t) & CTABLE & "> nodes found instead of 1 in file:" & CHR$(13) & sOutputFile, MB_ICONSTOP, MNAME
				Goto AfterError
			End Select
		End If
	Next t
	For t = 0 to 3	' Create missing Table nodes
		If IsNull(tabNodes(t)) Then
			tabNodes(t) = oDocument.createElement(IOtypes(t) & CTABLE)
			For i = t + 1 to 3
				If Not IsNull(tabNodes(i)) Then
					oRoot.insertBefore(tabNodes(t), tabNodes(i))
					Goto continue
				End If
			Next i
			oRoot.appendChild(tabNodes(t))
		End If
		continue:
	Next t

	Exit Function
	ErrorHandler:
	MsgBox Error$ & CHR$(13) & sOutputFile, MB_ICONSTOP, MNAME
	AfterError:
	ParseXML = True
End Function

REM Check if string has a special character that needs to be replaced with XML entity
REM String may already contain XML entities e.g. '&lt;' which will be exluded
Private Function CheckEntity(start As Integer, s As String, key As String) As Integer
	Dim offs As Integer, tail As Integer, namp As Integer
	skipent:
	offs = InStr(start, s, key)
	If (key = "&" AND offs > 0) Then
		tail = InStr(offs + 3, s, ";")	'Assuming XML entity has at least 2 chars and terminates with semicolon ';'
		If tail > 0 Then
			namp = InStr(offs + 1, s, "&")	'Offset of the next ampersand, if it exists
			If (namp = 0 OR namp > tail) Then	' Offset of the next ampersand is after the semicolon
				start = tail + 1
				Goto skipent
			End If
		End If
	End If
	CheckEntity = offs
End Function

REM Substitute special characters '&', ''' (single quote) '"' (double quote) '<' and '>' with XML entities
Private Function ToEntity(s As String) As String
	Dim offs As Integer, i As Integer, l As Integer, from As Integer
	l = Len(s)
	For i = 0 to UBound(HtmlEntities(0))
		from = 1
		offs = CheckEntity(from, s, HtmlEntities(0)(i))
		While offs <> 0
			from = offs + 1
			s = Left(s, offs - 1) & "&" & HtmlEntities(1)(i) & ";" & Right(s, l - offs)
			l = Len(s)
			offs = CheckEntity(from, s, HtmlEntities(0)(i))
		Wend
	Next i
	ToEntity = s
End Function

REM Populate Row array with roffset cell contents. Cols -- number of columns to use
Sub GetRow(oSheet, Row, roffset, Cols) 'Sheet object, row value array, row number to read, number of columns to read
	Dim c As Integer
	ReDim Row(Cols - 1)
	For c = 0 to Cols - 1
		Row(c) = ToEntity(Trim(oSheet.getCellByPosition(c + GfirstCol - 1, roffset + 1).string))
	Next c
End Sub

REM Count number of rows by checking column "B"
Sub CountRows(oSheet, Rows)
	Dim r As Long
	Dim sValue As String
	Rows = 0
	sValue = Trim(oSheet.getCellByPosition(GfirstCol - 1, 0).string) 'Check cell B1
	If sValue = "" Then Exit Sub
	REM Get number of rows of column 1 (B1:B1048576)
	r = 0 'Begin with B2
	While sValue <> ""
		r = r + 1
		sValue = Trim(oSheet.getCellByPosition(GfirstCol - 1, r).string)
	Wend
	Rows = r - 1
End Sub

REM Copy contents of the first row (B1,C1...256) to Head array
REM Cols will containg number of processed columns (0 to 256)
Sub GetHeaders(oSheet, Head, Cols)
	Dim c As Integer
	Dim sValue As String
	Cols = 0
	sValue = Trim(oSheet.getCellByPosition(GfirstCol - 1, 0).string) 'Check cell B1
	If sValue = "" Then Exit Sub
	ReDim Head(255)
	For c = 0 to UBound(Head())	'Begin with B1
		sValue = Trim(oSheet.getCellByPosition(c + GfirstCol - 1, 0).string)
		If sValue = "" Then Exit For
		Head(c) = sValue
	Next c
	Cols = c
	If Cols > 0 Then
		ReDim Preserve Head(Cols - 1)
	Else
		ReDim Head()
	End If
End	Sub

REM Fill Sheets array with names of all sheets in the current document.
REM Fill IOSheets array with names of selected DI/AI/DO/AO sheets.
Sub GetSheets(sSelName, IOSheets, Sheets)
	Dim i As Integer, count As Integer
	count = ThisComponent.Sheets().Count
	For i = 0 to count - 1
		Sheets(i) = ThisComponent.Sheets(i).Name
	Next i
	ReDim Preserve Sheets(count - 1)
	If PartStringInArray(IOtypes(), UCase(Right(sSelName, 2)), 1) < 0 Then
		'MsgBox "Current sheet name doesn't have suffix DI", MB_ICONEXCLAMATION, MNAME
		Exit Sub
	End If
	sSelName = RTrimStr(sSelName, Right(sSelName, 3))
	For i = 0 to 3
		If ThisComponent.Sheets().hasByName(sSelName & "_" & IOtypes(i)) Then
			IOSheets(i) = sSelName & "_" & IOtypes(i)
		End If
	Next i
End Sub

REM Returns string of all attributes and their values
Function AttString(oAttList) As String
	Dim oAtt As Object		' Attribute node
	Dim sValue As String	' Attribute value
	Dim sAtts As String		' String containing all attributes and values
	Dim i%

	If Not IsNull(oAttList) Then
		For i = 0 To oAttList.Length - 1
			oAtt = oAttList.item(i)
			sValue = oAtt.getNodeValue()
			If sValue <> "" Then
				If sValue = EMPTYKEY Then sValue = ""
				sAtts = sAtts & " "
				If oAtt.Prefix <> "" Then sAtts = sAtts & oAtt.Prefix & ":"
				sAtts = sAtts & oAtt.getNodeName() & "=""" & sValue & """"
			End If
		Next i
	End If
	AttString = sAtts
End Function

REM Add namespace attributes to the root element
Function NamespaceString(oRoot As Object)
	Dim i As Integer
	Dim oAttList As Object	' Attribute list
	Dim oAttr As Object		' Attribute node
	Dim sAtts As String		' String containing all attributes and values

	If oRoot.NamespaceURI <> "" Then
		sAtts = " xmlns=""" & oRoot.NamespaceURI & """"
	End If
	If UBound(NScontainer) >= 0 Then
		For i = 0 To UBound(NScontainer)
			If NScontainer(i)(0) <> "" Then
				sAtts = sAtts & " xmlns:" & NScontainer(i)(0) & "=""" & NScontainer(i)(1) & """"
			End If
		Next i
	End If
	NamespaceString = sAtts & AttString(oRoot.getAttributes)
End Function

REM Checks recursively, if there is some child node with content, be it text,
REM comment or attribute value other than "". Returns True, if some content
REM is found in the tree, otherwise False.
REM It is called by PrintDom, before the actual element start tag is written,
REM so PrintDom can abstain from printing an empty element.
Function HasContent(oElementList) As Boolean
	Dim oChild 'A single child node of oElementList
	Dim oAttList As Object 'The attributes of oChild
	Dim oChildren 'The child nodes of oChild
	Dim i%, j%
	For i = 0 To oElementList.getLength - 1
		oChild = oElementList.item(i)
		If oChild.hasAttributes Then
			oAttList = oChild.getAttributes()
			For j = 0 To oAttList.Length - 1
				If oAttList.item(j).getNodeValue <> "" Then
					HasContent = True
					Exit Function
				End If
			Next
		End If
		If oChild.getNodeType = com.sun.star.xml.dom.NodeType.TEXT_NODE Then
			If oChild.getNodeValue <> "" Then
				HasContent = True
				Exit Function
			End If
		Elseif oChild.getNodeType = com.sun.star.xml.dom.NodeType.COMMENT_NODE Then
			If oChild.getNodeValue <> "" Then
				HasContent = True
				Exit Function
			End If
		Else
			oChildren = oChild.getChildNodes()
			If oChildren.getLength <> 0 Then
				REM Start the recursion
				If HasContent(oChildren) Then
					HasContent = True
					Exit Function
				End If
			End If
		End If
	Next
	HasContent = False
End Function

REM Writes the elements of a DOM tree recursively line after line into a
REM text file. Mark to start with iLevel 0.
REM Indents lower levels by accumulating iIndent spaces (or sICh), except for text nodes:
REM they are written in the same line directly after the element start tag
REM and are directly followed by the element end tag.
REM It is assumed that there are either one text node or one or more other
REM child nodes to an element, if any.
REM Optional sICh parameter sets character used for indentation. If skipped, space (" ") will be used
Sub PrintDom(oNode, oStream, iLevel As Integer, iIndent As Integer, Optional sICh As String)
	Dim oElementChildren
	Dim oChild
	Dim sLine As String
	Dim sAtt As String
	Dim sIndent As String
	Dim i As Long, iLen As Long
	Dim sNodeName As String
	sNodeName = oNode.getNodeName
	If IsMissing(sICh) Then
		sIndent = String(iLevel * iIndent, " ")
	Else
		sIndent = String(iLevel * iIndent, sICh)
	End If
	REM Only comments and elements are treated.
	If oNode.getNodeType = com.sun.star.xml.dom.NodeType.COMMENT_NODE Then
		sLine = sIndent & "<!--" & oNode.getNodeValue() & "-->"
		oStream.writeString(sLine & CHR$(10))
	Elseif oNode.getNodeType = com.sun.star.xml.dom.NodeType.ELEMENT_NODE Then
		If iLevel = 0 Then
			sAtt = NamespaceString(oNode)
		Else
			sAtt = AttString(oNode.getAttributes)
		End If
		REM Check, if the element has data. Otherwise the element is skipped.
		If oNode.hasChildNodes() OR sAtt <> "" Then
			oElementChildren = oNode.getChildNodes()
			If HasContent(oElementChildren) Then
				sLine = sIndent & "<" & sNodeName & sAtt & ">" 'Start tag line
				iLen = oElementChildren.getLength()
				If iLen = 1 Then
					REM Test for text node, assuming that there are no other
					REM sibling nodes besides a text node.
					oChild = oElementChildren.item(0)
					If oChild.getNodeType = com.sun.star.xml.dom.NodeType.TEXT_NODE Then
						sLine = sLine & oChild.getNodeValue() & "</" & sNodeName & ">"
						REM Write the line: start tag plus text value plus end tag.
						oStream.writeString(sLine & CHR$(10))
						Exit Sub
					End If
				End If
				REM At this point there are child elements other than text nodes.
				REM Write the start tag line.
				oStream.writeString(sLine & CHR$(10))
				For i = 0 To iLen - 1
					REM Start the recursion, increment the indentation level
					PrintDom(oElementChildren.item(i), oStream, iLevel + 1, iIndent, CHR$(9))
				Next
				sLine = sIndent & "</" & sNodeName & ">" 'End tag line
				REM Write the end tag line.
				oStream.writeString(sLine & CHR$(10))
			Else
				REM There are no child elements to be written.
				REM If there are attributes, the short notation is used.
				REM If there are not attributes, empty element tag is written.
				sLine = sIndent & "<" & sNodeName & sAtt & "/>"
				REM Write the element in a short notation line.
				oStream.writeString(sLine & CHR$(10))
			End If
		End If
	End If
End Sub

REM Writes the structure of the current DOM document
REM into a valid and nicely formatted XML text file.
REM This subroutine and the called subroutines and functions are not
REM document specific and needn't be part of the document library.
REM They can be stored in the "My Macros" library.
Function WriteDomToFile(oDOM, sFilePath As String) As Boolean
	Dim oOutputStream 'Stream returned from SimpleFileAccess
	Dim oTextOutput 'TextOutputStream service
	Dim oNodes 'List of child nodes of the root node
	Dim i As Long
	Dim iIndentLevel As Integer 'Indentation level
	Dim iIndentSpaces As Integer 'Number of spaces added to each indentation level
	Dim sFileURL As String
	On Error Goto ErrorHandler

	iIndentSpaces = 1
	REM Set the output stream
	sFileURL = ConvertToURL(sFilePath)
	With oSimpleFileAccess
		If .exists(sFileURL) Then .kill(sFileURL)
		oOutputStream = .openFileWrite(sFileURL)
	End With
	oTextOutput = createUnoService("com.sun.star.io.TextOutputStream")
	With oTextOutput
		.OutputStream = oOutputStream
		.setEncoding("UTF-8")
		REM The first line is a processing instruction. It usually isn't part of
		REM the DOM tree. So we write it separately.
		.WriteString("<?xml version=""1.0"" encoding=""UTF-8""?>" & CHR$(10))
		REM A DOM tree can consist of zero, one or more child nodes on the
		REM root level. The child nodes are treated hierarchically.
		oNodes = oDOM.getChildNodes
		For i = 0 To oNodes.getLength - 1
			PrintDom(oNodes.item(i), oTextOutput, iIndentLevel, iIndentSpaces, CHR$(9))
		Next
		.closeOutput
	End With
	oOutputStream.closeOutput

	Exit Function
	ErrorHandler:
	MsgBox Error$ & CHR$(13) & sOutputFile, MB_ICONSTOP, MNAME
	WriteDomToFile = True
End Function

Sub Main
	GlobalScope.BasicLibraries.LoadLibrary("Tools")
	Dim oSheet
	Dim IOSheets(3) As String		' Sheets used for generating xml file (DI, AI, DO, AO and their descriptions)
	Dim Sheets(255) As String		' List of all sheet names
	Dim sSelName As String			' Name of the current sheet
	Dim Head() As String			' Headers (The text content from cell B1 to the first empty cell of row 1)
	Dim Row() As String				' Text contents of the cells of a row
	Dim oModule, oDlg				' Dialog object
	Dim oComboBox(3), oCheckBox(3)	' Dialog elements
	Dim i As Integer, t As Integer, c As Integer
	Dim r As Long
	Dim Cols As Integer, Rows As Long 'Number of columns and rows
	Dim UsedIOs As Integer			' Number of IO sheets to be used
	Dim oDocBuilder As Object		' DocumentBuilder interface
	Dim oDocument As Object			' DOM Document object
	Dim tabNodes(3) As Object		' <DITable>, <AITable>, ...
	Dim oElement As Object			' <DI>, <AI>, <DO>, <AO>
	Dim sComment As String

	GlobalInit()
	IOtypes = Array("DI", "AI", "DO", "AO")
	HtmlEntities = Array(Array("&", CHR$(34), "'", "<", ">"), Array("amp", "quot", "apos", "lt", "gt"))

	oSheet = ThisComponent.CurrentController.ActiveSheet
	sSelName = oSheet.Name
	GetSheets(sSelName, IOSheets, Sheets)

	DialogLibraries.LoadLibrary("XML_lib")
	oDialogLib = DialogLibraries.getByName("XML_lib")
	oModule = oDialogLib.getByName("ExportDialog")
	oDlg = CreateUnoDialog(oModule)
	oComboBox(0) = oDlg.GetControl("DIComboBox")
	oComboBox(1) = oDlg.GetControl("AIComboBox")
	oComboBox(2) = oDlg.GetControl("DOComboBox")
	oComboBox(3) = oDlg.GetControl("AOComboBox")
	oCheckBox(0) = oDlg.GetControl("DICheckBox")
	oCheckBox(1) = oDlg.GetControl("AICheckBox")
	oCheckBox(2) = oDlg.GetControl("DOCheckBox")
	oCheckBox(3) = oDlg.GetControl("AOCheckBox")
	oFileName1 = oDlg.GetControl("FileName1")
	oFileName2 = oDlg.GetControl("FileName2")
	oOptionButton1 = oDlg.GetControl("OptionButton1")
	oOptionButton2 = oDlg.GetControl("OptionButton2")
	oNameExistsLabel = oDlg.GetControl("NameExists")
	oOWCheckBox = oDlg.GetControl("OWCheckBox")
	oOKbutt = oDlg.GetControl("ExportButt")
	'oDlg.Title = "Export to XML file"
	DialogSetPosition(oDlg, ThisComponent.getCurrentController().Frame.getComponentWindow(), 0)
	REM Fill entries to the DIComboBox
	For t = 0 to 3
		for i = 0 to UBound(Sheets())
			oComboBox(t).additem(Sheets(i), i)
			If UCase(Sheets(i)) = UCase(IOSheets(t)) Then
				oComboBox(t).Text = oComboBox(t).Items(i)
				oCheckBox(t).State = 1
			End If
		Next i
	Next t
	If ThisComponent.url = "" Then
		MsgBox "Current document is not saved", MB_ICONEXCLAMATION, MNAME
		Exit Sub
	End If
	oFileName2.Text = ResolveFilepath(sSelName & ".xml")

	If oDlg.execute = 0 Then
		'MsgBox "Export canceled", MB_ICONINFORMATION, MNAME
		Exit Sub
	End If

	UsedIOs = 0 'Number of currently used IO sheets
	For t = 0 to 3
		If oCheckBox(t).State = 1 Then
			IOSheets(t) = Trim(oComboBox(t).Text)
			REM Check if sheet name is selected
			If IOSheets(t) = "" Then
				MsgBox IOtypes(t) & " selected, but sheet name missing", MB_ICONEXCLAMATION, MNAME
				Exit Sub
			End If
			REM Check if sheet exists
			if Not ThisComponent.Sheets().hasByName(IOSheets(t)) Then
				MsgBox "Sheet """ & IOSheets(t) & """ doesn't exist", MB_ICONEXCLAMATION, MNAME
				Exit Sub
			End If
			UsedIOs = UsedIOs + 1
			sComment = sComment & " " & IOtypes(t) & "=""" & IOSheets(t) & """"
		Else
			IOSheets(t) = ""
		End If
	Next t
	If UsedIOs = 0 Then
		MsgBox "No sheets selected", MB_ICONEXCLAMATION, MNAME
		Exit Sub
	End If

	oDocBuilder = createUnoService("com.sun.star.xml.dom.DocumentBuilder")

	If oOptionButton2.State Then
		sOutputFile = ResolveFilepath(oFileName2.Text)
		If oSimpleFileAccess.exists(sOutputFile) Then
			'If MsgBox("Confirm to overwrite:" & CHR$(13) & sOutputFile, MB_ICONQUESTION + MB_OKCANCEL, MNAME) <> IDOK Then Exit Sub
			If ParseXML(oDocBuilder, oDocument, tabNodes()) Then
				'MsgBox "Error occured while trying to parse file:" & CHR$(13) & sOutputFile, MB_ICONSTOP, MNAME
				Exit Sub
			End If
		End If
	Else
		sOutputFile = ResolveFilepath(oFileName1.Text)
	End If

	If IsNull(oDocument) Then	' Create new XML object
		If NewXML(oDocBuilder, oDocument, tabNodes(), sComment) Then
			'MsgBox "Error occured while trying to parse file:" & CHR$(13) & sOutputFile, MB_ICONSTOP, MNAME
			Exit Sub
		End If
	End If

	For t = 0 to 3	' Create nodes for each selected sheet
		If IOSheets(t) <> "" Then
			oSheet = ThisComponent.Sheets().getByName(IOSheets(t)) 'Get sheet
			GetHeaders(oSheet, Head, Cols) 'Get headers of the sheet
			CountRows(oSheet, Rows)
			If (Cols = 0 OR Rows = 0) Then Goto emptynode
			For r = 0 to Rows - 1
				oElement = oDocument.createElement(IOtypes(t))
				tabNodes(t).appendChild(oElement)
				GetRow(oSheet, Row, r, Cols)
				For c = 0 to Cols - 1
					oElement.setAttribute(Head(c), Row(c))
				Next c
			Next r
		Else
			emptynode:
			oElement = oDocument.createElement("")
			tabNodes(t).appendChild(oElement)
		End If
	Next t
	If WriteDomToFile(oDocument, sOutputFile) Then Exit Sub
	MsgBox "Export complete!" & CHR$(13) & sOutputFile, MB_ICONINFORMATION, MNAME
End Sub

