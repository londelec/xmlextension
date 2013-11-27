package com.londelec.addon.leandcaddon;

import com.sun.star.awt.PushButtonType;
import com.sun.star.awt.XCheckBox;
import com.sun.star.awt.XComboBox;
import com.sun.star.awt.XDialog;
import com.sun.star.frame.XFrame;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.sheet.XSpreadsheet;
import com.sun.star.table.XCell;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import java.io.BufferedWriter;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.io.StringReader;
import java.io.StringWriter;
import java.io.Writer;
import java.util.Arrays;
import javax.xml.stream.XMLOutputFactory;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.XMLStreamWriter;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Source;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.stream.StreamResult;
import javax.xml.transform.stream.StreamSource;

/**
 * All addon commands, that can be chosen from menus.
 * 
 * @version 1.1.0
 * @author MP
 */
public class LeandcCommands extends SpreadsheetDocHelper
{
    private XComponentContext  mxContext;
    private XMultiComponentFactory mxServiceManager;
    
    /**
     * Initialize instance of 'LeandcCommands' class.
     * @param XComponentContext context current context.
     * @param XFrame frame current frame.
     */
    public LeandcCommands(XComponentContext context, XFrame frame)
    {
        super(context, frame);
        this.mxContext = context;
        this.mxServiceManager = mxContext.getServiceManager();
    }

    /**
     * Generate XML File with DI AI DO objects from chosen spreadsheets.
     */
    public void GenerateDataObjectXml()
    {
        UnoDialog generateXmlDialog = null;
            
        String[] spreadsheetList;
        String activeSpreadsheetName;
        
        try
        {
            File currentFile = new File(getCurrentFilePath());
                
            if("".equals(currentFile.getPath()))
            {
                UnoMessageBox infoBox = new UnoMessageBox(mxContext, mxServiceManager);
                infoBox.showError("Error", "Save spreadsheet before generating xml.");
                return;
            }
            
            generateXmlDialog = new UnoDialog(mxContext, mxServiceManager);
            
            generateXmlDialog.initialize( new String[] {"Height", "Moveable", "Name", "PositionX", "PositionY", "Step", "TabIndex","Title","Width"},
                new Object[] { new Integer(110), Boolean.TRUE, "GenerateXmlDialog", new Integer(102),new Integer(41), new Integer(0), new Short((short) 0), "Choose sheets to use for xml generation", new Integer(130)});
            
            spreadsheetList = getAllSpreadsheetNames();
            activeSpreadsheetName = getActiveSpreadsheetName();

            String diSheetName = getSheetName("DI", activeSpreadsheetName, spreadsheetList);
            String aiSheetName = getSheetName("AI", activeSpreadsheetName, spreadsheetList);
            String doSheetName = getSheetName("DO", activeSpreadsheetName, spreadsheetList);
            String aoSheetName = getSheetName("AO", activeSpreadsheetName, spreadsheetList);
            
            XCheckBox diCheckbox = generateXmlDialog.insertCheckBox(generateXmlDialog, 10, 10, 20, 8, "~DI", Boolean.FALSE, 1);
            XCheckBox aiCheckbox = generateXmlDialog.insertCheckBox(generateXmlDialog, 10, 30, 20, 8, "~AI", Boolean.FALSE, 1);
            XCheckBox doCheckbox = generateXmlDialog.insertCheckBox(generateXmlDialog, 10, 50, 20, 8, "~DO", Boolean.FALSE, 1);
            XCheckBox aoCheckbox = generateXmlDialog.insertCheckBox(generateXmlDialog, 10, 70, 20, 8, "~AO", Boolean.FALSE, 1);
            
            XComboBox diComboBox = generateXmlDialog.insertComboBox(generateXmlDialog, 30, 10, 90, 12, spreadsheetList, diSheetName);
            XComboBox aiComboBox = generateXmlDialog.insertComboBox(generateXmlDialog, 30, 30, 90, 12, spreadsheetList, aiSheetName);
            XComboBox doComboBox = generateXmlDialog.insertComboBox(generateXmlDialog, 30, 50, 90, 12, spreadsheetList, doSheetName);
            XComboBox aoComboBox = generateXmlDialog.insertComboBox(generateXmlDialog, 30, 70, 90, 12, spreadsheetList, aoSheetName);
            
            generateXmlDialog.insertButton(generateXmlDialog, 80, 90, 40, 14, "~Generate", PushButtonType.OK, "buttonGenerateDataObjectXml");
            
            generateXmlDialog.createWindowPeer();
            generateXmlDialog.xDialog = UnoRuntime.queryInterface(XDialog.class, generateXmlDialog.mxDialogControl);
            
            if(generateXmlDialog.executeDialog() == 1)
            {
                GenerateDataObjectXmlImpl(
                    (generateXmlDialog.getCheckBoxState(diCheckbox) == 1),
                    (generateXmlDialog.getCheckBoxState(aiCheckbox) == 1),
                    (generateXmlDialog.getCheckBoxState(doCheckbox) == 1),
                    (generateXmlDialog.getCheckBoxState(aoCheckbox) == 1),
                    generateXmlDialog.getComboBoxSelectedValue(diComboBox),
                    generateXmlDialog.getComboBoxSelectedValue(aiComboBox),
                    generateXmlDialog.getComboBoxSelectedValue(doComboBox),
                    generateXmlDialog.getComboBoxSelectedValue(aoComboBox));
            }
            
        }
        catch( Exception e )
        {
            System.err.println( e + e.getMessage());
        }
        finally
        {
            //make sure always to dispose the component and free the memory!
            if (generateXmlDialog != null)
            {
                if (generateXmlDialog.mxComponent != null)
                {
                    generateXmlDialog.mxComponent.dispose();
                }
            }
        }
    }
    
    private void GenerateDataObjectXmlImpl(Boolean isDISelected, Boolean isAISelected, Boolean isDOSelected, Boolean isAOSelected,
        String diSheetName, String aiSheetName, String doSheetName, String aoSheetName) throws IOException, XMLStreamException {
        
        ByteArrayOutputStream xmlStream = null;
        XMLStreamWriter xmlWriter = null;
        Writer fileWriter  = null;
        
        if(isDISelected || isAISelected || isDOSelected || isAOSelected) {
            try {
                
                xmlStream = new ByteArrayOutputStream();
                
                XMLOutputFactory xmlOutputFactory =  XMLOutputFactory.newInstance();
                xmlWriter = xmlOutputFactory.createXMLStreamWriter(xmlStream, "utf-8");
                
                xmlWriter.writeStartDocument("utf-8", "1.0");
                
                xmlWriter.writeComment(
                    getCommentForDataObjectXmlGeneration(isDISelected, isAISelected, isDOSelected, isAOSelected, diSheetName, aiSheetName, doSheetName, aoSheetName));
                        
                xmlWriter.writeStartElement("objects");
                
                if(isDISelected) {
                    // DI table
                    xmlWriter.writeStartElement("DITable");
                    if(!"".equals(diSheetName)) {
                        ObjectSheetToXmlStream("DI", diSheetName, xmlWriter);
                    }
                    xmlWriter.writeEndElement();//DITable
                }
                
                if(isAISelected) {
                    // AI table
                    xmlWriter.writeStartElement("AITable");
                    if(!"".equals(aiSheetName)) {
                        ObjectSheetToXmlStream("AI", aiSheetName, xmlWriter);
                    }
                    xmlWriter.writeEndElement();//AITable
                }
 
                if(isDOSelected) {
                    // DO table
                    xmlWriter.writeStartElement("DOTable");
                    if(!"".equals(doSheetName)) {
                        ObjectSheetToXmlStream("DO", doSheetName, xmlWriter);
                    }
                    xmlWriter.writeEndElement();//DOTable
                }
                
                if(isAOSelected) {
                    // AO table
                    xmlWriter.writeStartElement("AOTable");
                    if(!"".equals(aoSheetName)) {
                        ObjectSheetToXmlStream("AO", aoSheetName, xmlWriter);
                    }
                    xmlWriter.writeEndElement();//AOTable
                }
                
                xmlWriter.writeEndElement();//generatedObjects
                
                String result = prettyFormat(xmlStream.toString("utf-8"), 8);
                
                File currentFile = new File(getCurrentFilePath());
                String currentFileName = currentFile.getName();
                String outputFilePath = currentFile.getPath().replace(currentFileName, "output.xml");

                fileWriter = new BufferedWriter(new OutputStreamWriter(
                    new FileOutputStream(outputFilePath), "utf-8"));

                fileWriter.write(result);

                UnoMessageBox infoBox = new UnoMessageBox(mxContext, mxServiceManager);
                infoBox.showInfo("Success", "Xml generated successfully!\nFile location: " + outputFilePath);

            } catch (Exception ex) {
                System.err.println( "Error: caught exception in GenerateDataObjectXmlImpl()!\nException Message = " + ex.getMessage());
            }
            finally {
                
                if(xmlStream != null) {
                    xmlStream.flush();
                    xmlStream.close();
                }
                
                if(xmlWriter != null) {
                    xmlWriter.flush();
                    xmlWriter.close();
                }
                
                if(fileWriter != null) {
                    fileWriter.flush();
                    fileWriter.close();
                }
            }
        }
    }
    
    private void ObjectSheetToXmlStream(String nodeName, String sheetName, XMLStreamWriter xmlWriter) throws RuntimeException, Exception
    {
        int columnIterator = 1;
        int rowIterator = 1;
        XCell headerCell, sourceCell;
        String cellValue, headerValue; 
        
        XSpreadsheet sourceSheet = getSpreadsheet(sheetName);
        
        if(sourceSheet != null) {
            
            sourceCell = sourceSheet.getCellByPosition(columnIterator, rowIterator);
            cellValue = getString(sourceCell);
            
            while(!"".equals(cellValue))
            {
                headerCell = sourceSheet.getCellByPosition(columnIterator, 0);
                headerValue = getString(headerCell);
                
                //create xml node
                xmlWriter.writeStartElement(nodeName);
                
                while(!"".equals(headerValue))
                {
                    if(!"".equals(cellValue))
                    {    
                        //add attribute to xml node
                        xmlWriter.writeAttribute(headerValue, cellValue);
                    }

                    columnIterator++;
                    headerCell = sourceSheet.getCellByPosition(columnIterator, 0);
                    headerValue = getString(headerCell);
                    
                    sourceCell = sourceSheet.getCellByPosition(columnIterator, rowIterator);
                    cellValue = getString(sourceCell);
                }

                xmlWriter.writeEndElement();//nodeName
                        
                columnIterator = 1;
                rowIterator++;
                sourceCell = sourceSheet.getCellByPosition(columnIterator, rowIterator);
                cellValue = getString(sourceCell);
            }
        }        
    }    
    
    private String getSheetName(String objectType, String activeSheetName, String[] sheetNamesList)
    {
        String result = "";
        
        if("DI".equals(objectType))
        {
            if(activeSheetName.endsWith("_DI"))
            {
                result = activeSheetName;
            }
            else if(activeSheetName.endsWith("_AI"))
            {
                String relatedSpreadsheetName = activeSheetName.replace("_AI", "_DI");
                result = Arrays.asList(sheetNamesList).contains(relatedSpreadsheetName) ? relatedSpreadsheetName : "";
            }
            else if(activeSheetName.endsWith("_DO"))
            {
                String relatedSpreadsheetName = activeSheetName.replace("_DO", "_DI");
                result = Arrays.asList(sheetNamesList).contains(relatedSpreadsheetName) ? relatedSpreadsheetName : "";
            }
            else if(activeSheetName.endsWith("_AO"))
            {
                String relatedSpreadsheetName = activeSheetName.replace("_AO", "_DI");
                result = Arrays.asList(sheetNamesList).contains(relatedSpreadsheetName) ? relatedSpreadsheetName : "";
            }
            else
            {
                result = "";
            }
        }
        
        if("AI".equals(objectType))
        {
            if(activeSheetName.endsWith("_AI"))
            {
                result = activeSheetName;
            }
            else if(activeSheetName.endsWith("_DI"))
            {
                String relatedSpreadsheetName = activeSheetName.replace("_DI", "_AI");
                result = Arrays.asList(sheetNamesList).contains(relatedSpreadsheetName) ? relatedSpreadsheetName : "";
            }
            else if(activeSheetName.endsWith("_DO"))
            {
                String relatedSpreadsheetName = activeSheetName.replace("_DO", "_AI");
                result = Arrays.asList(sheetNamesList).contains(relatedSpreadsheetName) ? relatedSpreadsheetName : "";
            }
            else if(activeSheetName.endsWith("_AO"))
            {
                String relatedSpreadsheetName = activeSheetName.replace("_AO", "_AI");
                result = Arrays.asList(sheetNamesList).contains(relatedSpreadsheetName) ? relatedSpreadsheetName : "";
            }
            else
            {
                result = "";
            }
        }
        
        if("DO".equals(objectType))
        {
            if(activeSheetName.endsWith("_DO"))
            {
                result = activeSheetName;
            }
            else if(activeSheetName.endsWith("_DI"))
            {
                String relatedSpreadsheetName = activeSheetName.replace("_DI", "_DO");
                result = Arrays.asList(sheetNamesList).contains(relatedSpreadsheetName) ? relatedSpreadsheetName : "";
            }
            else if(activeSheetName.endsWith("_AI"))
            {
                String relatedSpreadsheetName = activeSheetName.replace("_AI", "_DO");
                result = Arrays.asList(sheetNamesList).contains(relatedSpreadsheetName) ? relatedSpreadsheetName : "";
            }
            else if(activeSheetName.endsWith("_AO"))
            {
                String relatedSpreadsheetName = activeSheetName.replace("_AO", "_DO");
                result = Arrays.asList(sheetNamesList).contains(relatedSpreadsheetName) ? relatedSpreadsheetName : "";
            }
            else
            {
                result = "";
            }
        }

        if("AO".equals(objectType))
        {
            if(activeSheetName.endsWith("_AO"))
            {
                result = activeSheetName;
            }
            else if(activeSheetName.endsWith("_DI"))
            {
                String relatedSpreadsheetName = activeSheetName.replace("_DI", "_AO");
                result = Arrays.asList(sheetNamesList).contains(relatedSpreadsheetName) ? relatedSpreadsheetName : "";
            }
            else if(activeSheetName.endsWith("_AI"))
            {
                String relatedSpreadsheetName = activeSheetName.replace("_AI", "_AO");
                result = Arrays.asList(sheetNamesList).contains(relatedSpreadsheetName) ? relatedSpreadsheetName : "";
            }
            else if(activeSheetName.endsWith("_DO"))
            {
                String relatedSpreadsheetName = activeSheetName.replace("_DO", "_AO");
                result = Arrays.asList(sheetNamesList).contains(relatedSpreadsheetName) ? relatedSpreadsheetName : "";
            }
            else
            {
                result = "";
            }
        }

        return result;
    }
    
    private String getCommentForDataObjectXmlGeneration(Boolean isDISelected, Boolean isAISelected, Boolean isDOSelected, Boolean isAOSelected, 
            String diSheetName, String aiSheetName, String doSheetName, String aoSheetName)
    {
        String comment = "";
        
        if(isDISelected && !"".equals(diSheetName))
        {
            comment += " DI=\"" + diSheetName + "\"";
        }
        
        if(isAISelected && !"".equals(aiSheetName))
        {
            comment += " AI=\"" + aiSheetName + "\"";
        }
        
        if(isDOSelected && !"".equals(doSheetName))
        {
            comment += " DO=\"" + doSheetName + "\"";
        }
        
        if(isAOSelected && !"".equals(aoSheetName))
        {
            comment += " AO=\"" + aoSheetName + "\"";
        }
        
        comment = "".equals(comment) ? "There where no sheets selected for XML generation." : ("XML generated from sheets:" + comment);

        return comment;
    }
    
    private String prettyFormat(String input, int indent) {
        try {
            Source xmlInput = new StreamSource(new StringReader(input));
            StringWriter stringWriter = new StringWriter();
            StreamResult xmlOutput = new StreamResult(stringWriter);
            TransformerFactory transformerFactory = TransformerFactory.newInstance();
            transformerFactory.setAttribute("indent-number", indent);
            Transformer transformer = transformerFactory.newTransformer(); 
            transformer.setOutputProperty(OutputKeys.INDENT, "yes");
            transformer.transform(xmlInput, xmlOutput);
            
            String result = xmlOutput.getWriter().toString();
            return result.replace("\n        ", "\n\t").replace("\t        ", "\t\t");
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }
}

