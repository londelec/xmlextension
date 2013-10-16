package com.londelec.addon.leandcaddon;

import com.sun.star.awt.ActionEvent;
import com.sun.star.awt.AdjustmentEvent;
import com.sun.star.awt.AdjustmentType;
import com.sun.star.awt.FocusChangeReason;
import com.sun.star.awt.FocusEvent;
import com.sun.star.awt.ItemEvent;
import com.sun.star.awt.KeyEvent;
import com.sun.star.awt.MouseEvent;
import com.sun.star.awt.PosSize;
import com.sun.star.awt.PushButtonType;
import com.sun.star.awt.Rectangle;
import com.sun.star.awt.SpinEvent;
import com.sun.star.awt.TextEvent;
import com.sun.star.awt.XActionListener;
import com.sun.star.awt.XAdjustmentListener;
import com.sun.star.awt.XButton;
import com.sun.star.awt.XCheckBox;
import com.sun.star.awt.XComboBox;
import com.sun.star.awt.XControl;
import com.sun.star.awt.XControlContainer;
import com.sun.star.awt.XControlModel;
import com.sun.star.awt.XDialog;
import com.sun.star.awt.XFixedText;
import com.sun.star.awt.XFocusListener;
import com.sun.star.awt.XItemListener;
import com.sun.star.awt.XKeyListener;
import com.sun.star.awt.XListBox;
import com.sun.star.awt.XMouseListener;
import com.sun.star.awt.XPointer;
import com.sun.star.awt.XReschedule;
import com.sun.star.awt.XSpinField;
import com.sun.star.awt.XSpinListener;
import com.sun.star.awt.XTextComponent;
import com.sun.star.awt.XTextListener;
import com.sun.star.awt.XToolkit;
import com.sun.star.awt.XTopWindow;
import com.sun.star.awt.XWindow;
import com.sun.star.awt.XWindowPeer;
import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.XMultiPropertySet;
import com.sun.star.beans.XPropertySet;
import com.sun.star.container.XNameAccess;
import com.sun.star.container.XNameContainer;
import com.sun.star.frame.XDesktop;
import com.sun.star.frame.XFrame;
import com.sun.star.frame.XModel;
import com.sun.star.lang.EventObject;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.lang.XMultiServiceFactory;
import com.sun.star.script.BasicErrorException;
import com.sun.star.text.XTextDocument;
import com.sun.star.ucb.XFileIdentifierConverter;
import com.sun.star.uno.AnyConverter;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import com.sun.star.util.XNumberFormats;
import com.sun.star.util.XNumberFormatsSupplier;

/**
 * Class to create and show dialogs.
 * 
 * @version 1.0.0
 * @author MP
 */
public class UnoDialog implements XTextListener, XSpinListener, XActionListener, XFocusListener, XMouseListener, XItemListener, XAdjustmentListener, XKeyListener {

    protected XComponentContext mxContext = null;
    protected XMultiComponentFactory mxMultiComponentFactory;
    protected XMultiServiceFactory mxMultiServiceFactory;
    protected XModel mxModel;
    protected XNameContainer mxDialogModelNameContainer;
    protected XControlContainer mxDialogContainer;
    protected XControl mxDialogControl;
    protected XDialog xDialog;
    protected XReschedule mxReschedule;
    protected XWindowPeer mxWindowPeer = null;
    protected XTopWindow mxTopWindow = null;
    protected XFrame mxFrame = null;
    protected XComponent mxComponent = null;

    /**
     * Creates a new instance of UnoDialog
     */
    public UnoDialog(XComponentContext _xContext, XMultiComponentFactory _xMCF) {
        mxContext = _xContext;
        mxMultiComponentFactory = _xMCF;
        createDialog(mxMultiComponentFactory);
    }

    /**
     * @param _sKeyName
     * @return
     */
    public XNameAccess getRegistryKeyContent(String _sKeyName) {
        try {
            Object oConfigProvider;
            PropertyValue[] aNodePath = new PropertyValue[1];
            oConfigProvider = mxMultiComponentFactory.createInstanceWithContext("com.sun.star.configuration.ConfigurationProvider", this.mxContext);
            aNodePath[0] = new PropertyValue();
            aNodePath[0].Name = "nodepath";
            aNodePath[0].Value = _sKeyName;
            XMultiServiceFactory xMSFConfig = (XMultiServiceFactory) UnoRuntime.queryInterface(XMultiServiceFactory.class, oConfigProvider);
            Object oNode = xMSFConfig.createInstanceWithArguments("com.sun.star.configuration.ConfigurationAccess", aNodePath);
            XNameAccess xNameAccess = (XNameAccess) UnoRuntime.queryInterface(XNameAccess.class, oNode);
            return xNameAccess;
        } catch (Exception exception) {
            exception.printStackTrace(System.err);
            return null;
        }
    }

    private void createDialog(XMultiComponentFactory _xMCF) {
        try {
            Object oDialogModel = _xMCF.createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", mxContext);

            // The XMultiServiceFactory of the dialogmodel is needed to instantiate the controls...
            mxMultiServiceFactory = UnoRuntime.queryInterface(XMultiServiceFactory.class, oDialogModel);

            // The named container is used to insert the created controls into...
            mxDialogModelNameContainer = UnoRuntime.queryInterface(XNameContainer.class, oDialogModel);

            // create the dialog...
            Object oUnoDialog = _xMCF.createInstanceWithContext("com.sun.star.awt.UnoControlDialog", mxContext);
            mxDialogControl = UnoRuntime.queryInterface(XControl.class, oUnoDialog);

            // The scope of the control container is public...
            mxDialogContainer = UnoRuntime.queryInterface(XControlContainer.class, oUnoDialog);

            mxTopWindow = UnoRuntime.queryInterface(XTopWindow.class, mxDialogContainer);

            // link the dialog and its model...
            XControlModel xControlModel = UnoRuntime.queryInterface(XControlModel.class, oDialogModel);
            mxDialogControl.setModel(xControlModel);

        } catch (Exception exception) {
            exception.printStackTrace(System.err);
        }
    }

    public short executeDialog() throws BasicErrorException {
        if (mxWindowPeer == null) {
            createWindowPeer();
        }
        xDialog = (XDialog) UnoRuntime.queryInterface(XDialog.class, mxDialogControl);
        mxComponent = (XComponent) UnoRuntime.queryInterface(XComponent.class, mxDialogControl);
        return xDialog.execute();
    }

    public void initialize(String[] PropertyNames, Object[] PropertyValues) throws BasicErrorException {
        try {
            XMultiPropertySet xMultiPropertySet = UnoRuntime.queryInterface(XMultiPropertySet.class, mxDialogModelNameContainer);
            xMultiPropertySet.setPropertyValues(PropertyNames, PropertyValues);
        } catch (Exception ex) {
            ex.printStackTrace(System.err);
        }
    }

    /**
     * Creates a peer for this dialog, using the active OO frame as the parent
     * window.
     *
     * @return
     * @throws java.lang.Exception
     */
    public XWindowPeer createWindowPeer() throws BasicErrorException {
        return createWindowPeer(null);
    }

    /**
     * create a peer for this dialog, using the given peer as a parent.
     *
     * @param parentPeer
     * @return
     * @throws java.lang.Exception
     */
    public XWindowPeer createWindowPeer(XWindowPeer _xWindowParentPeer) throws BasicErrorException {
        try {
            if (_xWindowParentPeer == null) {
                XWindow xWindow = (XWindow) UnoRuntime.queryInterface(XWindow.class, mxDialogContainer);
                xWindow.setVisible(false);
                Object tk = mxMultiComponentFactory.createInstanceWithContext("com.sun.star.awt.Toolkit", mxContext);
                XToolkit xToolkit = (XToolkit) UnoRuntime.queryInterface(XToolkit.class, tk);
                mxReschedule = (XReschedule) UnoRuntime.queryInterface(XReschedule.class, xToolkit);
                mxDialogControl.createPeer(xToolkit, _xWindowParentPeer);
                mxWindowPeer = mxDialogControl.getPeer();
                return mxWindowPeer;
            }
        } catch (Exception exception) {
            exception.printStackTrace(System.err);
        }
        return null;
    }

    public void calculateDialogPosition(XWindow _xWindow) {
        Rectangle aFramePosSize = mxModel.getCurrentController().getFrame().getComponentWindow().getPosSize();
        Rectangle CurPosSize = _xWindow.getPosSize();
        int WindowHeight = aFramePosSize.Height;
        int WindowWidth = aFramePosSize.Width;
        int DialogWidth = CurPosSize.Width;
        int DialogHeight = CurPosSize.Height;
        int iXPos = ((WindowWidth / 2) - (DialogWidth / 2));
        int iYPos = ((WindowHeight / 2) - (DialogHeight / 2));
        _xWindow.setPosSize(iXPos, iYPos, DialogWidth, DialogHeight, PosSize.POS);
    }

    public void endExecute() {
        xDialog.endExecute();
    }

    public Object insertControlModel(String ServiceName, String sName, String[] sProperties, Object[] sValues) throws com.sun.star.script.BasicErrorException {
        try {
            Object oControlModel = mxMultiServiceFactory.createInstance(ServiceName);
            XMultiPropertySet xControlMultiPropertySet = (XMultiPropertySet) UnoRuntime.queryInterface(XMultiPropertySet.class, oControlModel);
            xControlMultiPropertySet.setPropertyValues(sProperties, sValues);
            mxDialogModelNameContainer.insertByName(sName, oControlModel);
            return oControlModel;
        } catch (com.sun.star.uno.Exception exception) {
            exception.printStackTrace(System.err);
            return null;
        }
    }

    public XFixedText insertFixedText(XMouseListener _xMouseListener, int _nPosX, int _nPosY, int _nWidth, int _nStep, String _sLabel) {
        XFixedText xFixedText = null;
        try {
            // create a unique name by means of an own implementation...
            String sName = createUniqueName(mxDialogModelNameContainer, "Label");

            // create a controlmodel at the multiservicefactory of the dialog model...
            Object oFTModel = mxMultiServiceFactory.createInstance("com.sun.star.awt.UnoControlFixedTextModel");
            XMultiPropertySet xFTModelMPSet = UnoRuntime.queryInterface(XMultiPropertySet.class, oFTModel);
            
            // Set the properties at the model - keep in mind to pass the property names in alphabetical order!
            xFTModelMPSet.setPropertyValues(
                    new String[]{"Height", "Name", "PositionX", "PositionY", "Step", "Width"},
                    new Object[]{new Integer(8), sName, new Integer(_nPosX), new Integer(_nPosY), new Integer(_nStep), new Integer(_nWidth)});
            
            // add the model to the NameContainer of the dialog model
            mxDialogModelNameContainer.insertByName(sName, oFTModel);

            // The following property may also be set with XMultiPropertySet but we
            // use the XPropertySet interface merely for reasons of demonstration
            XPropertySet xFTPSet = UnoRuntime.queryInterface(XPropertySet.class, oFTModel);
            xFTPSet.setPropertyValue("Label", _sLabel);

            // reference the control by the Name
            XControl xFTControl = mxDialogContainer.getControl(sName);
            xFixedText = UnoRuntime.queryInterface(XFixedText.class, xFTControl);
            XWindow xWindow = UnoRuntime.queryInterface(XWindow.class, xFTControl);
            xWindow.addMouseListener(_xMouseListener);
            
        } catch (Exception ex) {
            /* perform individual exception handling here.
             * Possible exception types are:
             * com.sun.star.lang.IllegalArgumentException,
             * com.sun.star.lang.WrappedTargetException,
             * com.sun.star.container.ElementExistException,
             * com.sun.star.beans.PropertyVetoException,
             * com.sun.star.beans.UnknownPropertyException,
             * com.sun.star.uno.Exception
             */
            ex.printStackTrace(System.err);
        }
        
        return xFixedText;
    }

    /**
     * makes a String unique by appending a numerical suffix
     *
     * @param _xElementContainer the com.sun.star.container.XNameAccess
     * container that the new Element is going to be inserted to
     * @param _sElementName the StemName of the Element
     */
    public static String createUniqueName(XNameAccess _xElementContainer, String _sElementName) {
        boolean bElementexists = true;
        int i = 1;
        String sIncSuffix = "";
        String BaseName = _sElementName;
        while (bElementexists) {
            bElementexists = _xElementContainer.hasByName(_sElementName);
            if (bElementexists) {
                i += 1;
                _sElementName = BaseName + Integer.toString(i);
            }
        }
        return _sElementName;
    }

    /**
     * Create checkBox and add to current dialog.
     * @param XItemListener _xItemListener
     * @param int posX
     * @param int posY
     * @param int height
     * @param width
     * @param String label
     * @param Boolean isTristate
     * @param int state
     * @return created instance of XCheckBox.
     */
    public XCheckBox insertCheckBox(XItemListener _xItemListener, 
        int posX, int posY, int width, int height, String label, Boolean isTristate, int state) {
        XCheckBox xCheckBox = null;
        try {
            // create a unique name by means of an own implementation...
            String sName = createUniqueName(mxDialogModelNameContainer, "CheckBox");

            // create a controlmodel at the multiservicefactory of the dialog model...
            Object oCBModel = mxMultiServiceFactory.createInstance("com.sun.star.awt.UnoControlCheckBoxModel");

            // Set the properties at the model - keep in mind to pass the property names in alphabetical order!
            XMultiPropertySet xCBMPSet = UnoRuntime.queryInterface(XMultiPropertySet.class, oCBModel);
            xCBMPSet.setPropertyValues(
                    new String[]{"Height", "Label", "Name", "State", "PositionX", "PositionY", "TriState", "Width"},
                    new Object[]{new Integer(height), label, sName, new Short((short) state), new Integer(posX), new Integer(posY), isTristate, new Integer(width)});

            // add the model to the NameContainer of the dialog model
            mxDialogModelNameContainer.insertByName(sName, oCBModel);
            XControl xCBControl = mxDialogContainer.getControl(sName);
            xCheckBox = UnoRuntime.queryInterface(XCheckBox.class, xCBControl);
            
            // An ActionListener will be notified on the activation of the button...
            xCheckBox.addItemListener(_xItemListener);
            
        } catch (Exception ex) {
            /* perform individual exception handling here.
             * Possible exception types are:
             * com.sun.star.lang.IllegalArgumentException,
             * com.sun.star.lang.WrappedTargetException,
             * com.sun.star.container.ElementExistException,
             * com.sun.star.beans.PropertyVetoException,
             * com.sun.star.beans.UnknownPropertyException,
             * com.sun.star.uno.Exception
             */
            ex.printStackTrace(System.err);
        }
        
        return xCheckBox;
    }

    /**
     * Get current state of specified XCheckBox.
     * @param XCheckBox xCheckBox
     * @return short value of XCheckBox current state.
     */
    public short getCheckBoxState(XCheckBox xCheckBox)
    {
        return xCheckBox.getState();
    }

    public void insertRadioButtonGroup(short _nTabIndex, int _nPosX, int _nPosY, int _nWidth) {
        try {
            // create a unique name by means of an own implementation...
            String sName = createUniqueName(mxDialogModelNameContainer, "OptionButton");

            // create a controlmodel at the multiservicefactory of the dialog model...
            Object oRBModel = mxMultiServiceFactory.createInstance("com.sun.star.awt.UnoControlRadioButtonModel");
            XMultiPropertySet xRBMPSet = (XMultiPropertySet) UnoRuntime.queryInterface(XMultiPropertySet.class, oRBModel);
            // Set the properties at the model - keep in mind to pass the property names in alphabetical order!
            xRBMPSet.setPropertyValues(
                    new String[]{"Height", "Label", "Name", "PositionX", "PositionY", "State", "TabIndex", "Width"},
                    new Object[]{new Integer(8), "~First Option", sName, new Integer(_nPosX), new Integer(_nPosY), new Short((short) 1), new Short(_nTabIndex++), new Integer(_nWidth)});
            // add the model to the NameContainer of the dialog model
            mxDialogModelNameContainer.insertByName(sName, oRBModel);

            // create a unique name by means of an own implementation...
            sName = createUniqueName(mxDialogModelNameContainer, "OptionButton");

            oRBModel = mxMultiServiceFactory.createInstance("com.sun.star.awt.UnoControlRadioButtonModel");
            xRBMPSet = (XMultiPropertySet) UnoRuntime.queryInterface(XMultiPropertySet.class, oRBModel);
            // Set the properties at the model - keep in mind to pass the property names in alphabetical order!
            xRBMPSet.setPropertyValues(
                    new String[]{"Height", "Label", "Name", "PositionX", "PositionY", "TabIndex", "Width"},
                    new Object[]{new Integer(8), "~Second Option", sName, new Integer(130), new Integer(214), new Short(_nTabIndex), new Integer(150)});
            // add the model to the NameContainer of the dialog model
            mxDialogModelNameContainer.insertByName(sName, oRBModel);
        } catch (com.sun.star.uno.Exception ex) {
            /* perform individual exception handling here.
             * Possible exception types are:
             * com.sun.star.lang.IllegalArgumentException,
             * com.sun.star.lang.WrappedTargetException,
             * com.sun.star.container.ElementExistException,
             * com.sun.star.beans.PropertyVetoException,
             * com.sun.star.beans.UnknownPropertyException,
             * com.sun.star.uno.Exception
             */
            ex.printStackTrace(System.err);
        }
    }

    public XListBox insertListBox(int _nPosX, int _nPosY, int _nWidth, int _nStep, String[] _sStringItemList) {
        XListBox xListBox = null;
        try {
            // create a unique name by means of an own implementation...
            String sName = createUniqueName(mxDialogModelNameContainer, "ListBox");

            // create a controlmodel at the multiservicefactory of the dialog model...
            Object oListBoxModel = mxMultiServiceFactory.createInstance("com.sun.star.awt.UnoControlListBoxModel");
            XMultiPropertySet xLBModelMPSet = (XMultiPropertySet) UnoRuntime.queryInterface(XMultiPropertySet.class, oListBoxModel);
            // Set the properties at the model - keep in mind to pass the property names in alphabetical order!
            xLBModelMPSet.setPropertyValues(
                    new String[]{"Dropdown", "Height", "Name", "PositionX", "PositionY", "Step", "StringItemList", "Width"},
                    new Object[]{Boolean.TRUE, new Integer(12), sName, new Integer(_nPosX), new Integer(_nPosY), new Integer(_nStep), _sStringItemList, new Integer(_nWidth)});
            // The following property may also be set with XMultiPropertySet but we
            // use the XPropertySet interface merely for reasons of demonstration
            XPropertySet xLBModelPSet = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xLBModelMPSet);
            xLBModelPSet.setPropertyValue("MultiSelection", Boolean.TRUE);
            short[] nSelItems = new short[]{(short) 1, (short) 3};
            xLBModelPSet.setPropertyValue("SelectedItems", nSelItems);
            // add the model to the NameContainer of the dialog model
            mxDialogModelNameContainer.insertByName(sName, xLBModelMPSet);
            XControl xControl = mxDialogContainer.getControl(sName);
            // retrieve a ListBox that is more convenient to work with than the Model of the ListBox...
            xListBox = (XListBox) UnoRuntime.queryInterface(XListBox.class, xControl);
        } catch (com.sun.star.uno.Exception ex) {
            /* perform individual exception handling here.
             * Possible exception types are:
             * com.sun.star.lang.IllegalArgumentException,
             * com.sun.star.lang.WrappedTargetException,
             * com.sun.star.container.ElementExistException,
             * com.sun.star.beans.PropertyVetoException,
             * com.sun.star.beans.UnknownPropertyException,
             * com.sun.star.uno.Exception
             */
            ex.printStackTrace(System.err);
        }
        return xListBox;
    }

    /**
     * Create comboBox and add to current dialog.
     * @param XItemListener _xItemListener
     * @param int posX
     * @param int posY
     * @param int height
     * @param int width
     * @param String[] listItems
     * @return created instance of XComboBox.
     */
    public XComboBox insertComboBox(XItemListener _xItemListener, 
        int posX, int posY, int width, int height, String[] listItems, String text) {
        XComboBox xComboBox = null;
        try {
            // create a unique name by means of an own implementation...
            String sName = createUniqueName(mxDialogModelNameContainer, "ComboBox");

            // create a controlmodel at the multiservicefactory of the dialog model...
            Object oComboBoxModel = mxMultiServiceFactory.createInstance("com.sun.star.awt.UnoControlComboBoxModel");
            XMultiPropertySet xCbBModelMPSet = UnoRuntime.queryInterface(XMultiPropertySet.class, oComboBoxModel);
            
            // Set the properties at the model - keep in mind to pass the property names in alphabetical order!
            xCbBModelMPSet.setPropertyValues(
                    new String[]{"Dropdown", "Height", "Name", "PositionX", "PositionY", "ReadOnly", "StringItemList", "Text", "Width"},
                    new Object[]{Boolean.TRUE, new Integer(height), sName, new Integer(posX), new Integer(posY), Boolean.FALSE, listItems, text, new Integer(width)});

            // add the model to the NameContainer of the dialog model
            mxDialogModelNameContainer.insertByName(sName, xCbBModelMPSet);
            XControl xControl = mxDialogContainer.getControl(sName);

            // retrieve a ListBox that is more convenient to work with than the Model of the ListBox...
            xComboBox = UnoRuntime.queryInterface(XComboBox.class, xControl);
            
            // An ActionListener will be notified on the activation of the button...
            xComboBox.addItemListener(_xItemListener);
            
        } catch (com.sun.star.uno.Exception ex) {
            /* perform individual exception handling here.
             * Possible exception types are:
             * com.sun.star.lang.IllegalArgumentException,
             * com.sun.star.lang.WrappedTargetException,
             * com.sun.star.container.ElementExistException,
             * com.sun.star.beans.PropertyVetoException,
             * com.sun.star.beans.UnknownPropertyException,
             * com.sun.star.uno.Exception
             */
            ex.printStackTrace(System.err);
        }
        
        return xComboBox;
    }
    
    /**
     * Get selected value from specified XComboBox.
     * @param XComboBox xComboBox
     * @return XComboBox selected value as String
     */
    public String getComboBoxSelectedValue(XComboBox xComboBox)
    {
        XTextComponent textComponent = UnoRuntime.queryInterface(XTextComponent.class, xComboBox);

        return textComponent.getText();
    }

    public XPropertySet insertFormattedField(XSpinListener _xSpinListener, int _nPosX, int _nPosY, int _nWidth) {
        XPropertySet xFFModelPSet = null;
        try {
            // create a unique name by means of an own implementation...
            String sName = createUniqueName(mxDialogModelNameContainer, "FormattedField");

            // create a controlmodel at the multiservicefactory of the dialog model...
            Object oFFModel = mxMultiServiceFactory.createInstance("com.sun.star.awt.UnoControlFormattedFieldModel");
            XMultiPropertySet xFFModelMPSet = (XMultiPropertySet) UnoRuntime.queryInterface(XMultiPropertySet.class, oFFModel);
            // Set the properties at the model - keep in mind to pass the property names in alphabetical order!
            xFFModelMPSet.setPropertyValues(
                    new String[]{"EffectiveValue", "Height", "Name", "PositionX", "PositionY", "StrictFormat", "Spin", "Width"},
                    new Object[]{new Double(12348), new Integer(12), sName, new Integer(_nPosX), new Integer(_nPosY), Boolean.TRUE, Boolean.TRUE, new Integer(_nWidth)});

            xFFModelPSet = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, oFFModel);
            // to define a numberformat you always need a locale...
            com.sun.star.lang.Locale aLocale = new com.sun.star.lang.Locale();
            aLocale.Country = "US";
            aLocale.Language = "en";
            // this Format is only compliant to the english locale!
            String sFormatString = "NNNNMMMM DD, YYYY";

            // a NumberFormatsSupplier has to be created first "in the open countryside"...
            Object oNumberFormatsSupplier = mxMultiComponentFactory.createInstanceWithContext("com.sun.star.util.NumberFormatsSupplier", mxContext);
            XNumberFormatsSupplier xNumberFormatsSupplier = (XNumberFormatsSupplier) UnoRuntime.queryInterface(XNumberFormatsSupplier.class, oNumberFormatsSupplier);
            XNumberFormats xNumberFormats = xNumberFormatsSupplier.getNumberFormats();
            // is the numberformat already defined?
            int nFormatKey = xNumberFormats.queryKey(sFormatString, aLocale, true);
            if (nFormatKey == -1) {
                // if not then add it to the NumberFormatsSupplier
                nFormatKey = xNumberFormats.addNew(sFormatString, aLocale);
            }

            // The following property may also be set with XMultiPropertySet but we
            // use the XPropertySet interface merely for reasons of demonstration
            xFFModelPSet.setPropertyValue("FormatsSupplier", xNumberFormatsSupplier);
            xFFModelPSet.setPropertyValue("FormatKey", new Integer(nFormatKey));

            // The controlmodel is not really available until inserted to the Dialog container
            mxDialogModelNameContainer.insertByName(sName, oFFModel);

            // finally we add a Spin-Listener to the control
            XControl xFFControl = mxDialogContainer.getControl(sName);
            // add a SpinListener that is notified on each change of the controlvalue...
            XSpinField xSpinField = (XSpinField) UnoRuntime.queryInterface(XSpinField.class, xFFControl);
            xSpinField.addSpinListener(_xSpinListener);

        } catch (com.sun.star.uno.Exception ex) {
            /* perform individual exception handling here.
             * Possible exception types are:
             * com.sun.star.lang.IllegalArgumentException,
             * com.sun.star.lang.WrappedTargetException,
             * com.sun.star.container.ElementExistException,
             * com.sun.star.beans.PropertyVetoException,
             * com.sun.star.beans.UnknownPropertyException,
             * com.sun.star.uno.Exception
             */
            ex.printStackTrace(System.err);
        }
        return xFFModelPSet;
    }

    public XTextComponent insertFileControl(XTextListener textListener, String pathText, int posX, int posY, int width) {
        
        XTextComponent textComponent = null;
        
        try {
            // create a unique name by means of an own implementation...
            String fileControlName = createUniqueName(mxDialogModelNameContainer, "FileControl");

            String workPath = pathText;
            
            if("".equals(pathText)) {
                // retrieve the configured Work path...
                Object oPathSettings = mxMultiComponentFactory.createInstanceWithContext("com.sun.star.util.PathSettings", mxContext);
                XPropertySet xPropertySet = UnoRuntime.queryInterface(XPropertySet.class, oPathSettings);
                String systemWorkPath = (String) xPropertySet.getPropertyValue("Work");
                
                // convert the Url to a system path that is "human readable"...
                Object fileContentProvider = mxMultiComponentFactory.createInstanceWithContext("com.sun.star.ucb.FileContentProvider", mxContext);
                XFileIdentifierConverter xFileIdentifierConverter = UnoRuntime.queryInterface(XFileIdentifierConverter.class, fileContentProvider);
                workPath = xFileIdentifierConverter.getSystemPathFromFileURL(systemWorkPath);
            }

            // create a controlmodel at the multiservicefactory of the dialog model...
            Object fileControlModel = mxMultiServiceFactory.createInstance("com.sun.star.awt.UnoControlFileControlModel");

            // Set the properties at the model - keep in mind to pass the property names in alphabetical order!
            XMultiPropertySet fileControlModelMultiProperty = UnoRuntime.queryInterface(XMultiPropertySet.class, fileControlModel);
            fileControlModelMultiProperty.setPropertyValues(
                    new String[]{"Height", "Name", "PositionX", "PositionY", "Text", "Width"},
                    new Object[]{new Integer(14), fileControlName, new Integer(posX), new Integer(posY), workPath, new Integer(width)});

            // The controlmodel is not really available until inserted to the Dialog container
            mxDialogModelNameContainer.insertByName(fileControlName, fileControlModel);
            //XPropertySet fileControlModelProperty = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, fileControlModel);

            // add a textlistener that is notified on each change of the controlvalue...
            XControl fileControlControl = mxDialogContainer.getControl(fileControlName);
            textComponent = UnoRuntime.queryInterface(XTextComponent.class, fileControlControl);
            //XWindow xFCWindow = UnoRuntime.queryInterface(XWindow.class, fileControlControl);
            textComponent.addTextListener(textListener);
        } catch (com.sun.star.uno.Exception ex) {
            /* perform individual exception handling here.
             * Possible exception types are:
             * com.sun.star.lang.IllegalArgumentException,
             * com.sun.star.lang.WrappedTargetException,
             * com.sun.star.container.ElementExistException,
             * com.sun.star.beans.PropertyVetoException,
             * com.sun.star.beans.UnknownPropertyException,
             * com.sun.star.uno.Exception
             */
            ex.printStackTrace(System.err);
        }
        
        return textComponent;
    }

    public XButton insertButton(XActionListener _xActionListener, 
        int posX, int posY, int width, int height, String label, PushButtonType pushButtonType, String controlName) {
        XButton xButton = null;
        try {

            // create a controlmodel at the multiservicefactory of the dialog model...
            Object oButtonModel = mxMultiServiceFactory.createInstance("com.sun.star.awt.UnoControlButtonModel");
            XMultiPropertySet xButtonMPSet = UnoRuntime.queryInterface(XMultiPropertySet.class, oButtonModel);
            
            // Set the properties at the model - keep in mind to pass the property names in alphabetical order!
            xButtonMPSet.setPropertyValues(
                    new String[]{"Height", "Label", "Name", "PositionX", "PositionY", "PushButtonType", "Width"},
                    new Object[]{new Integer(height), label, controlName, new Integer(posX), new Integer(posY), (short) pushButtonType.getValue(), new Integer(width)});

            // add the model to the NameContainer of the dialog model
            mxDialogModelNameContainer.insertByName(controlName, oButtonModel);
            XControl xButtonControl = mxDialogContainer.getControl(controlName);
            xButton = UnoRuntime.queryInterface(XButton.class, xButtonControl);
            
            // An ActionListener will be notified on the activation of the button...
            xButton.addActionListener(_xActionListener);
            
        } catch (com.sun.star.uno.Exception ex) {
            /* perform individual exception handling here.
             * Possible exception types are:
             * com.sun.star.lang.IllegalArgumentException,
             * com.sun.star.lang.WrappedTargetException,
             * com.sun.star.container.ElementExistException,
             * com.sun.star.beans.PropertyVetoException,
             * com.sun.star.beans.UnknownPropertyException,
             * com.sun.star.uno.Exception
             */
            ex.printStackTrace(System.err);
        }
        return xButton;
    }

    /**
     * gets the WindowPeer of a frame
     *
     * @param _XTextDocument the instance of a textdocument
     * @return the windowpeer of the frame
     */
    public XWindowPeer getWindowPeer(XTextDocument _xTextDocument) {
        XModel xModel = (XModel) UnoRuntime.queryInterface(XModel.class, _xTextDocument);
        XFrame xFrame = xModel.getCurrentController().getFrame();
        XWindow xWindow = xFrame.getContainerWindow();
        XWindowPeer xWindowPeer = (XWindowPeer) UnoRuntime.queryInterface(XWindowPeer.class, xWindow);
        return xWindowPeer;
    }

    public XFrame getCurrentFrame() {
        XFrame xRetFrame = null;
        try {
            Object oDesktop = mxMultiComponentFactory.createInstanceWithContext("com.sun.star.frame.Desktop", mxContext);
            XDesktop xDesktop = (XDesktop) UnoRuntime.queryInterface(XDesktop.class, oDesktop);
            xRetFrame = xDesktop.getCurrentFrame();
        } catch (Exception ex) {
            System.err.println( "Error: caught exception in getCurrentFrame()!\nException Message = " + ex.getMessage());
        }
        return xRetFrame;
    }

    public void textChanged(TextEvent textEvent) {
        try {
            // get the control that has fired the event,
            XControl xControl = (XControl) UnoRuntime.queryInterface(XControl.class, textEvent.Source);
            XControlModel xControlModel = xControl.getModel();
            XPropertySet xPSet = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xControlModel);
            String sName = (String) xPSet.getPropertyValue("Name");
            // just in case the listener has been added to several controls,
            // we make sure we refer to the right one
            if (sName.equals("TextField1")) {
                String sText = (String) xPSet.getPropertyValue("Text");
                System.out.println(sText);
                // insert your code here to validate the text of the control...
            }
        } catch (com.sun.star.uno.Exception ex) {
            /* perform individual exception handling here.
             * Possible exception types are:
             * com.sun.star.lang.WrappedTargetException,
             * com.sun.star.beans.UnknownPropertyException,
             * com.sun.star.uno.Exception
             */
            ex.printStackTrace(System.err);
        }
    }

    public void up(SpinEvent spinEvent) {
        try {
            // get the control that has fired the event,
            XControl xControl = (XControl) UnoRuntime.queryInterface(XControl.class, spinEvent.Source);
            XControlModel xControlModel = xControl.getModel();
            XPropertySet xPSet = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xControlModel);
            String sName = (String) xPSet.getPropertyValue("Name");
            // just in case the listener has been added to several controls,
            // we make sure we refer to the right one
            if (sName.equals("FormattedField1")) {
                double fvalue = AnyConverter.toDouble(xPSet.getPropertyValue("EffectiveValue"));
                System.out.println("Controlvalue:  " + fvalue);
                // insert your code here to validate the value of the control...
            }
        } catch (com.sun.star.uno.Exception ex) {
            /* perform individual exception handling here.
             * Possible exception types are:
             * com.sun.star.lang.WrappedTargetException,
             * com.sun.star.beans.UnknownPropertyException,
             * com.sun.star.uno.Exception
             */
            ex.printStackTrace(System.err);
        }
    }

    public void down(SpinEvent spinEvent) {
    }

    public void last(SpinEvent spinEvent) {
    }

    public void first(SpinEvent spinEvent) {
    }

    public void disposing(EventObject rEventObject) {
    }

    public void actionPerformed(ActionEvent rEvent) {
        try {
            // get the control that has fired the event,
            XControl xControl = UnoRuntime.queryInterface(XControl.class, rEvent.Source);
            XControlModel xControlModel = xControl.getModel();
            XPropertySet xPSet = UnoRuntime.queryInterface(XPropertySet.class, xControlModel);
            String sName = (String) xPSet.getPropertyValue("Name");
            // just in case the listener has been added to several controls,
            // we make sure we refer to the right one
            if (sName.equals("ActionButton")) {
                //...
            }
        } catch (com.sun.star.uno.Exception ex) {
            /* perform individual exception handling here.
             * Possible exception types are:
             * com.sun.star.lang.WrappedTargetException,
             * com.sun.star.beans.UnknownPropertyException,
             * com.sun.star.uno.Exception
             */
            ex.printStackTrace(System.err);
        }
    }

    public void focusLost(FocusEvent _focusEvent) {
        short nFocusFlags = _focusEvent.FocusFlags;
        int nFocusChangeReason = nFocusFlags & FocusChangeReason.TAB;
        if (nFocusChangeReason == FocusChangeReason.TAB) {
            // get the window of the Window that has gained the Focus...
            // Note that the xWindow is just a representation of the controlwindow
            // but not of the control itself
            XWindow xWindow = (XWindow) UnoRuntime.queryInterface(XWindow.class, _focusEvent.NextFocus);
        }
    }

    public void focusGained(FocusEvent focusEvent) {
    }

    public void mouseReleased(MouseEvent mouseEvent) {
    }

    public void mousePressed(MouseEvent mouseEvent) {
    }

    public void mouseExited(MouseEvent mouseEvent) {
    }

    public void mouseEntered(MouseEvent _mouseEvent) {
        try {
            // retrieve the control that the event has been invoked at...
            XControl xControl = (XControl) UnoRuntime.queryInterface(XControl.class, _mouseEvent.Source);
            Object tk = mxMultiComponentFactory.createInstanceWithContext("com.sun.star.awt.Toolkit", mxContext);
            XToolkit xToolkit = (XToolkit) UnoRuntime.queryInterface(XToolkit.class, tk);
            // create the peer of the control by passing the windowpeer of the parent
            // in this case the windowpeer of the control
            xControl.createPeer(xToolkit, mxWindowPeer);
            // create a pointer object "in the open countryside" and set the type accordingly...
            Object oPointer = this.mxMultiComponentFactory.createInstanceWithContext("com.sun.star.awt.Pointer", this.mxContext);
            XPointer xPointer = (XPointer) UnoRuntime.queryInterface(XPointer.class, oPointer);
            xPointer.setType(com.sun.star.awt.SystemPointer.REFHAND);
            // finally set the created pointer at the windowpeer of the control
            xControl.getPeer().setPointer(xPointer);
        } catch (com.sun.star.uno.Exception ex) {
            throw new java.lang.RuntimeException("cannot happen...");
        }
    }

    public void itemStateChanged(ItemEvent itemEvent) {
        try {
            // retrieve the control that the event has been invoked at...
            XCheckBox xCheckBox = (XCheckBox) UnoRuntime.queryInterface(XCheckBox.class, itemEvent.Source);
            // retrieve the control that we want to disable or enable
            XControl xControl = (XControl) UnoRuntime.queryInterface(XControl.class, mxDialogContainer.getControl("CommandButton1"));
            XPropertySet xModelPropertySet = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xControl.getModel());
            short nState = xCheckBox.getState();
            boolean bdoEnable = true;
            switch (nState) {
                case 1:     // checked
                    bdoEnable = true;
                    break;
                case 0:     // not checked
                case 2:     // don't know
                    bdoEnable = false;
                    break;
            }
            // Alternatively we could have done it also this way:
            // bdoEnable = (nState == 1);
            xModelPropertySet.setPropertyValue("Enabled", new Boolean(bdoEnable));
        } catch (com.sun.star.uno.Exception ex) {
            /* perform individual exception handling here.
             * Possible exception types are:
             * com.sun.star.lang.IllegalArgumentException
             * com.sun.star.lang.WrappedTargetException,
             * com.sun.star.beans.UnknownPropertyException,
             * com.sun.star.beans.PropertyVetoException
             * com.sun.star.uno.Exception
             */
            ex.printStackTrace(System.err);
        }
    }

    public void adjustmentValueChanged(AdjustmentEvent _adjustmentEvent) {
        switch (_adjustmentEvent.Type.getValue()) {
            case AdjustmentType.ADJUST_ABS_value:
                System.out.println("The event has been triggered by dragging the thumb...");
                break;
            case AdjustmentType.ADJUST_LINE_value:
                System.out.println("The event has been triggered by a single line move..");
                break;
            case AdjustmentType.ADJUST_PAGE_value:
                System.out.println("The event has been triggered by a block move...");
                break;
        }
        System.out.println("The value of the scrollbar is: " + _adjustmentEvent.Value);
    }

    public void keyReleased(KeyEvent keyEvent) {
        int i = keyEvent.KeyChar;
        int n = keyEvent.KeyCode;
        int m = keyEvent.KeyFunc;
    }

    public void keyPressed(KeyEvent keyEvent) {
    }
}
