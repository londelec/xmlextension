package com.londelec.addon.leandcaddon;

import com.sun.star.awt.Rectangle;
import com.sun.star.awt.XMessageBox;
import com.sun.star.awt.XMessageBoxFactory;
import com.sun.star.awt.XWindow;
import com.sun.star.awt.XWindowPeer;
import com.sun.star.frame.XDesktop;
import com.sun.star.frame.XFrame;
import com.sun.star.frame.XModel;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;

/**
 * Class to create and show message boxes.
 * 
 * @version 1.0.0
 * @author MP
 */
public class UnoMessageBox  {

    protected XComponentContext mxContext = null;
    protected XMultiComponentFactory mxMultiComponentFactory;

    /** Creates a new instance of MessageBox */
    public UnoMessageBox(XComponentContext xContext, XMultiComponentFactory xMultiComponentFactory) {
        mxContext = xContext;
        mxMultiComponentFactory = xMultiComponentFactory;
    }

    /** Shows an info messageBox.
     *  @param title the title of the messageBox
     *  @param message the message of the messageBox
     */
    public void showInfo(String title, String message) throws Exception
    {
        try
        {
            XWindowPeer xWindowPeer = getWindowPeerOfFrame();
            if (xWindowPeer != null)
            {
                showMessageBox("infobox", xWindowPeer, title, message);
            }
            else
            {
                System.out.println("Could not retrieve current frame");
            }
        }
        catch (Exception ex) {
            System.err.println( "Error: caught exception in showInfo()!\nException Message = " + ex.getMessage());
        }
    }
    
    /** Shows an error messageBox.
     *  @param title the title of the messageBox
     *  @param message the message of the messageBox
     */
    public void showError(String title, String message) throws Exception
    {
        try
        {
            XWindowPeer xWindowPeer = getWindowPeerOfFrame();
            if (xWindowPeer != null)
            {
                showMessageBox("errorbox", xWindowPeer, title, message);
            }
            else
            {
                System.out.println("Could not retrieve current frame");
            }
        }
        catch (Exception ex) {
            System.err.println( "Error: caught exception in showError()!\nException Message = " + ex.getMessage());
        }
    }

    /** Shows an info messageBox.
     *  @param XWindowPeer parentWindowPeer the windowPeer of the parent window
     *  @param String title the title of the messageBox
     *  @param String message the message of the messageBox
     */
    private void showMessageBox(String messageType, XWindowPeer parentWindowPeer, String title, String message) {
        XComponent xComponent = null;
        try
        {
            Object toolkit = mxMultiComponentFactory.createInstanceWithContext("com.sun.star.awt.Toolkit", mxContext);
            XMessageBoxFactory xMessageBoxFactory = (XMessageBoxFactory) UnoRuntime.queryInterface(XMessageBoxFactory.class, toolkit);
            
            // rectangle may be empty if position is in the center of the parent peer
            Rectangle rectangle = new Rectangle();
            
            XMessageBox messageBox = xMessageBoxFactory
                .createMessageBox(parentWindowPeer, rectangle, messageType, com.sun.star.awt.MessageBoxButtons.BUTTONS_OK, title, message);
            
            xComponent = UnoRuntime.queryInterface(XComponent.class, messageBox);
            if (messageBox != null)
            {
                messageBox.execute();
            }
        }
        catch (Exception ex)
        {
            System.err.println( "Error: caught exception in showMessageBox()!\nException Message = " + ex.getMessage());
        }
        finally
        {
            //make sure always to dispose the component and free the memory!
            if (xComponent != null)
            {
                xComponent.dispose();
            }
        }
    }
    
    // helper method to get the window peer of a document
    private XWindowPeer getWindowPeerOfFrame() throws Exception
    {
        Object desktop = mxMultiComponentFactory.createInstanceWithContext("com.sun.star.frame.Desktop", mxContext);
        XDesktop xDesktop = UnoRuntime.queryInterface(com.sun.star.frame.XDesktop.class, desktop); 
        XComponent document = xDesktop.getCurrentComponent();

        XFrame frame = null;
        XWindowPeer xWindowPeer = null;

        if (document != null) {
            XModel model = UnoRuntime.queryInterface(XModel.class, document);
            frame = model.getCurrentController().getFrame();
        }

        if (frame != null){
            XWindow containerWindow = frame.getContainerWindow();
            if (containerWindow != null){
                xWindowPeer = UnoRuntime.queryInterface(XWindowPeer.class, containerWindow);
            }
        }

        return xWindowPeer;
    }
}
