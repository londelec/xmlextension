
package com.londelec.addon.leandcaddon;

import com.sun.star.beans.XPropertySet;
import com.sun.star.container.XIndexAccess;
import com.sun.star.container.XNamed;
import com.sun.star.frame.XController;
import com.sun.star.frame.XDesktop;
import com.sun.star.frame.XFrame;
import com.sun.star.frame.XStorable;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.sheet.XCellAddressable;
import com.sun.star.sheet.XCellRangeAddressable;
import com.sun.star.sheet.XSheetCellCursor;
import com.sun.star.sheet.XSheetCellRange;
import com.sun.star.sheet.XSheetOperation;
import com.sun.star.sheet.XSpreadsheet;
import com.sun.star.sheet.XSpreadsheetDocument;
import com.sun.star.sheet.XSpreadsheetView;
import com.sun.star.sheet.XSpreadsheets;
import com.sun.star.sheet.XUsedAreaCursor;
import com.sun.star.table.BorderLine;
import com.sun.star.table.CellAddress;
import com.sun.star.table.CellContentType;
import com.sun.star.table.CellRangeAddress;
import com.sun.star.table.TableBorder;
import com.sun.star.table.XCell;
import com.sun.star.table.XCellRange;
import com.sun.star.text.XText;
import com.sun.star.text.XTextCursor;
import com.sun.star.ucb.XFileIdentifierConverter;
import com.sun.star.uno.RuntimeException;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import com.sun.star.util.NumberFormat;
import com.sun.star.util.XNumberFormatTypes;
import com.sun.star.util.XNumberFormatsSupplier;
import java.text.DecimalFormat;

/**
 * This is a helper class for the spreadsheet.
 * @version 1.0.0
 * @author MP
 */
public class SpreadsheetDocHelper
{
    private XComponentContext  mxContext;
    private XFrame mxFrame;
    private XMultiComponentFactory  mxServiceManager;
    private XSpreadsheetDocument mxDocument;

    
    /**
     * Initialize instance of 'SpreadsheetDocHelper' class.
     * @param XComponentContext context current context.
     * @param XFrame frame current frame.
     */
    public SpreadsheetDocHelper(XComponentContext context, XFrame frame)
    {
        try
        {
            mxContext = context;
            mxFrame = frame;
            mxServiceManager = context.getServiceManager();
            mxDocument = initDocument();
        }
        catch (Exception ex)
        {
            System.err.println( "Error: Couldn't initialize spreadsheet helpers\nException Message = " + ex.getMessage());
        }
    }

    /** Returns the service manager of the connected office.
        @return  XMultiComponentFactory interface of the service manager. */
    public XMultiComponentFactory getServiceManager()
    {
        return mxServiceManager;
    }

    /** Returns the component context of the connected office
        @return  XComponentContext interface of the context. */
    public XComponentContext getContext()
    {
        return mxContext;
    }

    /** Returns the whole spreadsheet document.
        @return  XSpreadsheetDocument interface of the document. */
    public XSpreadsheetDocument getDocument()
    {
        return mxDocument;
    }

    /** Returns the spreadsheet with the specified index (0-based).
        @param nIndex  The index of the sheet.
        @return  XSpreadsheet interface of the sheet. */
    public XSpreadsheet getSpreadsheet( int nIndex )
    {
        // Collection of sheets
        XSpreadsheets xSheets = mxDocument.getSheets();
        XSpreadsheet xSheet = null;
        try
        {
            XIndexAccess xSheetsIA = (XIndexAccess)UnoRuntime.queryInterface(XIndexAccess.class, xSheets );
            xSheet = (XSpreadsheet) UnoRuntime.queryInterface(XSpreadsheet.class, xSheetsIA.getByIndex(nIndex));
        }
        catch (Exception ex)
        {
            System.err.println( "Error: caught exception in getSpreadsheet()!\nException Message = " + ex.getMessage());
        }
        
        return xSheet;
    }
    
    /** Returns the spreadsheet with the specified name.
        @param sheetName  The name of the sheet.
        @return  XSpreadsheet interface of the sheet. */
    public XSpreadsheet getSpreadsheet( String sheetName )
    {
        // Collection of sheets
        XSpreadsheets xSheets = mxDocument.getSheets();
        XSpreadsheet xSheet = null;
        try
        {
            if(xSheets.hasByName(sheetName))
            {
                xSheet = UnoRuntime.queryInterface(XSpreadsheet.class, xSheets.getByName(sheetName));
            }
        }
        catch (Exception ex)
        {
            System.err.println( "Error: caught exception in getSpreadsheet()!\nException Message = " + ex.getMessage());
        }
        
        return xSheet;
    }
    
    /** Returns the currently active spreadsheet.
         @return  XSpreadsheet interface of the sheet. */
    public XSpreadsheet getActiveSpreadsheet()
    {
        XSpreadsheet xSheet = null;
        try
        {
            XController controller = mxFrame.getController();
            XSpreadsheetView spreadsheetView = UnoRuntime.queryInterface(XSpreadsheetView.class, controller);
            xSheet = spreadsheetView.getActiveSheet();
        }
        catch (Exception ex)
        {
            System.err.println( "Error: caught exception in getActiveSpreadsheet()!\nException Message = " + ex.getMessage());
        }
        
        return xSheet;
    }

    /**
     * Get name of currently active spreadsheet.
     * @return Name of currently active spreadsheet.
     */
    public String getActiveSpreadsheetName()
    {
        XSpreadsheet xSheet = getActiveSpreadsheet();
        return getSpreadsheetName(xSheet);
    }
    
    /**
     * Get all spreadsheet names in current workbook.
     * @return Collection of all spreadsheet names in current workbook.
     */
    public String[] getAllSpreadsheetNames()
    {
        // Collection of sheets
        XSpreadsheets xSheets = mxDocument.getSheets();
        String[] sheetNames = {};
        try
        {
            sheetNames = xSheets.getElementNames();
        }
        catch (Exception ex)
        {
            System.err.println( "Error: caught exception in getAllSpreadsheetNames()!\nException Message = " + ex.getMessage());
        }
        
        return sheetNames;
    }
    
    /**
     * Set active spreadsheet.
     * @param xSheet  The XSpreadsheet interface of the spreadsheet.
     */
    public void setActiveSpreadsheet (XSpreadsheet xSheet)
    {
        XSpreadsheet xActiveSheet = null;
        try
        {
            XController controller = mxFrame.getController();
            XSpreadsheetView spreadsheetView = UnoRuntime.queryInterface(XSpreadsheetView.class, controller);
            spreadsheetView.setActiveSheet(xSheet);
        }
        catch (Exception ex)
        {
            System.err.println( "Error: caught exception in getSheetName()!\nException Message = " + ex.getMessage());
        }
    }    

    /** Inserts a new empty spreadsheet with the specified name.
        @param aName  The name of the new sheet.
        @return  The XSpreadsheet interface of the new sheet. */
    public XSpreadsheet insertSpreadsheet( String aName )
    {
        // Collection of sheets
        XSpreadsheets xSheets = mxDocument.getSheets();
        XSpreadsheet xSheet = null;
        try
        {
            XIndexAccess xSheetsIA = UnoRuntime.queryInterface(XIndexAccess.class, xSheets );
            xSheets.insertNewByName( aName, (short)xSheetsIA.getCount());
            xSheet = (XSpreadsheet)UnoRuntime.queryInterface(XSpreadsheet.class, xSheets.getByName( aName ));
        }
        catch (Exception ex)
        {
            System.err.println( "Error: caught exception in insertSpreadsheet()!\nException Message = " + ex.getMessage());
        }
        
        return xSheet;
    }

    /** Inserts a new empty spreadsheet with the specified name.
        @param aName  The name of the new sheet.
        @param nIndex  The insertion index.
        @return  The XSpreadsheet interface of the new sheet. */
    public XSpreadsheet insertSpreadsheet( String aName, short nIndex )
    {
        // Collection of sheets
        XSpreadsheets xSheets = mxDocument.getSheets();
        XSpreadsheet xSheet = null;
        try
        {
            xSheets.insertNewByName( aName, nIndex );
            xSheet = (XSpreadsheet)UnoRuntime.queryInterface(XSpreadsheet.class, xSheets.getByName( aName ));
        }
        catch (Exception ex)
        {
            System.err.println( "Error: caught exception in insertSpreadsheet()!\nException Message = " + ex.getMessage());
        }
        
        return xSheet;
    }

    /**
     * Get name of spreadsheet.
     * @param xSheet  The XSpreadsheet interface of the spreadsheet.
     * @return name of spreadsheet.
     */
    public String getSpreadsheetName (XSpreadsheet xSheet)
    {
        String sheetName = null;
        
        try
        {
            XNamed xNamed = UnoRuntime.queryInterface(XNamed.class, xSheet);
            sheetName = xNamed.getName();
        }
        catch (Exception ex)
        {
            System.err.println( "Error: caught exception in getSheetName()!\nException Message = " + ex.getMessage());
        }
        
        return sheetName;
    }    
    
    /**
     * Clear all used area in spreadsheet.
     * @param xSheet  The XSpreadsheet interface of the spreadsheet.
     */
    public void clearSpreadsheet (XSpreadsheet xSheet) throws RuntimeException, Exception
    {
        try
        {
            XSheetCellCursor sheetCellCursor = xSheet.createCursor();
            XUsedAreaCursor sheetUsedCursor = UnoRuntime.queryInterface(XUsedAreaCursor.class, sheetCellCursor);
            sheetUsedCursor.gotoStartOfUsedArea(false);
            sheetUsedCursor.gotoEndOfUsedArea(true);

            XSheetOperation sheetOperation = UnoRuntime.queryInterface(XSheetOperation.class, sheetCellCursor);
            sheetOperation.clearContents(1+2+4+8+16+32+64+128+256+512);//clear all contents
        }
        catch (Exception ex)
        {
            System.err.println( "Error: caught exception in clearSpreadsheet()!\nException Message = " + ex.getMessage());
        }
    }
    
    /**
     * Get system file path of currently opened document.
     * @return current file path as String
     */
    public String getCurrentFilePath() {
        XStorable xStore = UnoRuntime.queryInterface (XStorable.class, mxDocument); 
        String currentFileURL = xStore.getLocation();
        return convertFromURL(currentFileURL);
    }
    
    /**
     * Converts a (file) URL to a file path in system dependent notation.
     * @param String fileURL - file location in URL format.
     * @return file path in system dependent notation.
     */
    public String convertFromURL(String fileURL) {
        String systemPath = null;
        try
        {
            Object fileContentProvider = mxServiceManager.createInstanceWithContext("com.sun.star.ucb.FileContentProvider", mxContext); 
            
            XFileIdentifierConverter xFileConverter =UnoRuntime.queryInterface(XFileIdentifierConverter.class, fileContentProvider);
            systemPath = xFileConverter.getSystemPathFromFileURL(fileURL);                  
        }
        catch (Exception ex)
        {
            System.err.println( "Error: caught exception in convertFromURL()!\nException Message = " + ex.getMessage());
        }
        
        return systemPath;
    }
// ________________________________________________________________
// Methods to fill values into cells.

    /** Writes a double value into a spreadsheet.
        @param xSheet  The XSpreadsheet interface of the spreadsheet.
        @param aCellName  The address of the cell (or a named range).
        @param fValue  The value to write into the cell. */
    public void setValue( XSpreadsheet xSheet, String aCellName, double fValue ) 
        throws RuntimeException, Exception
    {
        xSheet.getCellRangeByName( aCellName ).getCellByPosition( 0, 0 ).setValue( fValue );
    }

    /** Writes a formula into a spreadsheet.
        @param xSheet  The XSpreadsheet interface of the spreadsheet.
        @param aCellName  The address of the cell (or a named range).
        @param aFormula  The formula to write into the cell. */
    public void setFormula( XSpreadsheet xSheet, String aCellName, String aFormula ) 
        throws RuntimeException, Exception
    {
        xSheet.getCellRangeByName( aCellName ).getCellByPosition( 0, 0 ).setFormula( aFormula );
    }
    
    /** Writes a string into a cell.
        @param xCell XCell 
        @param stringValue  The string to write into the cell. */
    public void setString(XCell xCell, String stringValue)
    {
        try
        {
            XText xText = UnoRuntime.queryInterface(XText.class, xCell);
            XTextCursor xTextCursor = xText.createTextCursor();

            xText.insertString(xTextCursor, stringValue, false);
        }
        catch (Exception ex)
        {
            System.err.println( "Error: caught exception in setString()!\nException Message = " + ex.getMessage());
        }
    }
    
    /**
     * Get value from cell as string.
     * @param xCell XCell with value to get.
     */
    public String getString(XCell xCell)
    {
        String cellValue = null;
        CellContentType contentType = xCell.getType();
        
        if(contentType == CellContentType.TEXT || contentType == CellContentType.VALUE || contentType == CellContentType.FORMULA)
        {
            XText xText = UnoRuntime.queryInterface(XText.class, xCell);
            cellValue = xText.getString();
        }
        
        if(contentType == CellContentType.EMPTY)
        {
            cellValue = "";
        }
        
        return cellValue;
    }

    /** Writes a date with standard date format into a spreadsheet.
        @param xSheet  The XSpreadsheet interface of the spreadsheet.
        @param aCellName  The address of the cell (or a named range).
        @param nDay  The day of the date.
        @param nMonth  The month of the date.
        @param nYear  The year of the date. */
    public void setDate( XSpreadsheet xSheet, String aCellName, int nDay, int nMonth, int nYear ) 
        throws RuntimeException, Exception
    {
        // Set the date value.
        XCell xCell = xSheet.getCellRangeByName( aCellName ).getCellByPosition( 0, 0 );
        String aDateStr = nMonth + "/" + nDay + "/" + nYear;
        xCell.setFormula( aDateStr );

        // Set standard date format.
        XNumberFormatsSupplier xFormatsSupplier = (XNumberFormatsSupplier) UnoRuntime.queryInterface(
                XNumberFormatsSupplier.class, getDocument() );
        XNumberFormatTypes xFormatTypes = (XNumberFormatTypes) UnoRuntime.queryInterface(
                XNumberFormatTypes.class, xFormatsSupplier.getNumberFormats() );
        int nFormat = xFormatTypes.getStandardFormat( NumberFormat.DATE, new com.sun.star.lang.Locale() );

        XPropertySet xPropSet = (XPropertySet) UnoRuntime.queryInterface( XPropertySet.class, xCell );
        xPropSet.setPropertyValue( "NumberFormat", new Integer( nFormat ) );
    }

    /** Draws a colored border around the range and writes the headline in the
        first cell.
        @param xSheet  The XSpreadsheet interface of the spreadsheet.
        @param aRange  The address of the cell range (or a named range).
        @param aHeadline  The headline text. */
    public void prepareRange( XSpreadsheet xSheet, String aRange, String aHeadline ) 
        throws RuntimeException, Exception
    {
        // draw border
        XCellRange xCellRange = xSheet.getCellRangeByName( aRange );
        XPropertySet xPropSet = (XPropertySet) UnoRuntime.queryInterface( XPropertySet.class, xCellRange );
        BorderLine aLine = new BorderLine();
        aLine.Color = 0x99CCFF;
        aLine.InnerLineWidth = aLine.LineDistance = 0;
        aLine.OuterLineWidth = 100;
        TableBorder aBorder = new TableBorder();
        aBorder.TopLine = aBorder.BottomLine = aBorder.LeftLine = aBorder.RightLine = aLine;
        aBorder.IsTopLineValid = aBorder.IsBottomLineValid = true;
        aBorder.IsLeftLineValid = aBorder.IsRightLineValid = true;
        xPropSet.setPropertyValue( "TableBorder", aBorder );

        // draw headline
        XCellRangeAddressable xAddr = (XCellRangeAddressable)
            UnoRuntime.queryInterface( XCellRangeAddressable.class, xCellRange );
        CellRangeAddress aAddr = xAddr.getRangeAddress();

        xCellRange = xSheet.getCellRangeByPosition(
            aAddr.StartColumn, aAddr.StartRow, aAddr.EndColumn, aAddr.StartRow );
        xPropSet = (XPropertySet)
            UnoRuntime.queryInterface( XPropertySet.class, xCellRange );
        xPropSet.setPropertyValue( "CellBackColor", new Integer( 0x99CCFF ) );
        // write headline
        XCell xCell = xCellRange.getCellByPosition( 0, 0 );
        xCell.setFormula( aHeadline );
        xPropSet = (XPropertySet)
            UnoRuntime.queryInterface( XPropertySet.class, xCell );
        xPropSet.setPropertyValue( "CharColor", new Integer( 0x003399 ) );
        xPropSet.setPropertyValue( "CharWeight", new Float( com.sun.star.awt.FontWeight.BOLD ) );
    }

// ________________________________________________________________
// Methods to create cell addresses and range addresses.

    /** Creates a CellAddress and initializes it
        with the given range.
        @param xSheet  The XSpreadsheet interface of the spreadsheet.
        @param aCell  The address of the cell (or a named cell). */
    public CellAddress createCellAddress(
            XSpreadsheet xSheet,
            String aCell ) throws RuntimeException, Exception
    {
        XCellAddressable xAddr = (XCellAddressable)
            UnoRuntime.queryInterface( XCellAddressable.class,
                xSheet.getCellRangeByName( aCell ).getCellByPosition( 0, 0 ) );
        return xAddr.getCellAddress();
    }

    /** Creates a CellRangeAddress and initializes
        it with the given range.
        @param xSheet  The XSpreadsheet interface of the spreadsheet.
        @param aRange  The address of the cell range (or a named range). */
    public CellRangeAddress createCellRangeAddress(
            XSpreadsheet xSheet, String aRange )
    {
        XCellRangeAddressable xAddr = (XCellRangeAddressable)
            UnoRuntime.queryInterface( XCellRangeAddressable.class,
                xSheet.getCellRangeByName( aRange ) );
        return xAddr.getRangeAddress();
    }

// ________________________________________________________________
// Methods to convert cell addresses and range addresses to strings.

    /** Returns the text address of the cell.
        @param nColumn  The column index.
        @param nRow  The row index.
        @return  A string containing the cell address. */
    public String getCellAddressString( int nColumn, int nRow )
    {
        String aStr = "";
        if (nColumn > 25)
        {
            aStr += (char) ('A' + nColumn / 26 - 1);
        }
        
        aStr += (char) ('A' + nColumn % 26);
        aStr += (nRow + 1);
        return aStr;
    }

    /** Returns the text address of the cell range.
        @param aCellRange  The cell range address.
        @return  A string containing the cell range address. */
    public String getCellRangeAddressString(
            CellRangeAddress aCellRange )
    {
        return
            getCellAddressString( aCellRange.StartColumn, aCellRange.StartRow )
            + ":"
            + getCellAddressString( aCellRange.EndColumn, aCellRange.EndRow );
    }

    /** Returns the text address of the cell range.
        @param xCellRange  The XSheetCellRange interface of the cell range.
        @param bWithSheet  true = Include sheet name.
        @return  A string containing the cell range address. */
    public String getCellRangeAddressString(
            XSheetCellRange xCellRange,
            boolean bWithSheet )
    {
        String aStr = "";
        if (bWithSheet)
        {
            XSpreadsheet xSheet = xCellRange.getSpreadsheet();
            com.sun.star.container.XNamed xNamed = (com.sun.star.container.XNamed)
                UnoRuntime.queryInterface( com.sun.star.container.XNamed.class, xSheet );
            aStr += xNamed.getName() + ".";
        }
        
        XCellRangeAddressable xAddr = (XCellRangeAddressable)
            UnoRuntime.queryInterface( XCellRangeAddressable.class, xCellRange );
        aStr += getCellRangeAddressString( xAddr.getRangeAddress() );
        return aStr;
    }

    /** Returns a list of addresses of all cell ranges contained in the collection.
        @param xRangesIA  The XIndexAccess interface of the collection.
        @return  A string containing the cell range address list. */
    public String getCellRangeListString(
            com.sun.star.container.XIndexAccess xRangesIA ) throws RuntimeException, Exception
    {
        String aStr = "";
        int nCount = xRangesIA.getCount();
        for (int nIndex = 0; nIndex < nCount; ++nIndex)
        {
            if (nIndex > 0)
            {
                aStr += " ";
            }
            
            Object aRangeObj = xRangesIA.getByIndex( nIndex );
            XSheetCellRange xCellRange = (XSheetCellRange)
                UnoRuntime.queryInterface( XSheetCellRange.class, aRangeObj );
            aStr += getCellRangeAddressString( xCellRange, false );
        }
        return aStr;
    }


    /** Get current spreadsheet document.
        @return  The XSpreadsheetDocument interface of the document. */
    private XSpreadsheetDocument initDocument()
            throws RuntimeException, Exception
    {
        Object desktop = mxServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", mxContext); 
          
        XDesktop xDesktop = (XDesktop)UnoRuntime.queryInterface(com.sun.star.frame.XDesktop.class, desktop); 
          
        XComponent document = xDesktop.getCurrentComponent();
 
        return UnoRuntime.queryInterface(XSpreadsheetDocument.class, document);
    }
}