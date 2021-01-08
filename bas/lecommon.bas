REM  *****  BASIC  *****
Option Explicit

Global Ginitialized As Boolean		' Global variable have been initilized
Global GfirstCol As Integer			' Default first column is "B"
Global GhomeURL As String			' The initial path for OpenFile function
Public Const EMPTYKEY As String = "#EMPTY#"	'Empty attribute placeholder
Public oSimpleFileAccess As Object


REM Initialize global variables
Public Function GlobalInit()
	If Not Ginitialized Then
		Ginitialized = 1
		GfirstCol = 2	' Default first column is "B"
		GlobalScope.BasicLibraries.LoadLibrary("Tools")
	End If
	oSimpleFileAccess = createUnoService("com.sun.star.ucb.SimpleFileAccess")
End Function

REM Position the dialog in the center of the parent object
Public Function DialogSetPosition(oDlg, oParent, rel As Boolean)
	Dim xPos As Integer, yPos As Integer 'Position of the dialog
	Dim psParent	'Dimensions of the parent component
	Dim sizeDlg		'Size of the dialog

	sizeDlg = oDlg.getSize()
	psParent = oParent.getPosSize()
	If rel Then
		psParent.X = 0
		psParent.Y = 0
	End If

	xPos = ((psParent.Width/2) - (sizeDlg.Width/2) + psParent.X) : If xPos < 1 then xPos = 1
	yPos = ((psParent.Height/2) - (sizeDlg.Height/2) + psParent.Y) : If yPos < 1 then yPos = 1
	oDlg.setPosSize(xPos, yPos, sizeDlg.Width, sizeDlg.Height, 3) 'Set the dialog position without changing it's size
End Function

REM Add current directory path to filename if necessary
REM Initialize Global Path to the current document's path (if saved) or User's home directory
Public Function ResolveFilepath(sPath As String)
	Dim separator As String, homePath As String
	separator = GetPathSeparator()
	sPath = Trim(sPath)
	If InStr(1, sPath, separator) = 0 Then
		If ThisComponent.URL = "" Then
			If GetGUIType() = 1 Then ' Initialize file path
				homePath = Environ("USERPROFILE") ' Windows
			Else
				homePath = Environ("HOME") ' Unix (Linux, MacOS,...)
			End If
			homePath = homePath & separator & sPath
		Else
			homePath = ConvertFromURL(ThisComponent.URL)
		End If
		sPath = ConvertFromURL(DirectoryNameoutofPath(homePath, separator)) & separator & sPath
	End If
	If GhomeURL = "" Then GhomeURL = ConvertToURL(DirectoryNameoutofPath(sPath, separator))
	ResolveFilepath = sPath
End Function

REM Launch file open dialog and return path of the selected file
Public Function OpenFile() as String
	Dim oDialog as Object
	Dim filePath As String
	Dim filters(1) as String

	filters(0) = "*.xml"
	filters(1) = "*.*"
	oDialog = CreateUnoService("com.sun.star.ui.dialogs.FilePicker")
	AddFiltersToDialog(filters(), oDialog)

	If oSimpleFileAccess.Exists(GhomeURL) Then
		oDialog.SetDisplayDirectory(GhomeURL)
	End If

	If oDialog.Execute() = 1 Then
		filePath = oDialog.Files(0)
		GhomeURL = DirectoryNameoutofPath(filePath, "/") ' URL is always separated by '/'
		OpenFile = filePath
	End If
	oDialog.Dispose()
End Function
