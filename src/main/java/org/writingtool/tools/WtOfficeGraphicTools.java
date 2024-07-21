/* WritingTool, a LibreOffice Extension based on LanguageTool
 * Copyright (C) 2024 Fred Kruse (https://fk-es.de)
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin St, Fifth Floor, Boston, MA  02110-1301
 * USA
 */
package org.writingtool.tools;

import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.XPropertySet;
import com.sun.star.graphic.XGraphic;
import com.sun.star.graphic.XGraphicProvider;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.lang.XMultiServiceFactory;
import com.sun.star.text.XText;
import com.sun.star.text.XTextContent;
import com.sun.star.text.XTextCursor;
import com.sun.star.text.XTextDocument;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;

/**
 * Some tools to handle graphics in LibreOffice/OpenOffice document context
 * @since 1.0 (WritingTool)
 * @author Fred Kruse
 */
public class WtOfficeGraphicTools {
  
  private final static int SIZE_FACTOR = 20;

  public static void insertGraphic(String strUrl, int size, XComponent xComp, XComponentContext xContext) {
    try {
      // get the remote office service manager
      XMultiComponentFactory xMCF = xContext.getServiceManager();

      // Querying for the interface XTextDocument on the xcomponent
      XTextDocument xTextDoc = UnoRuntime.queryInterface(XTextDocument.class, xComp);

      // Querying for the interface XMultiServiceFactory on the xtextdocument
      XMultiServiceFactory xMSFDoc = UnoRuntime.queryInterface(XMultiServiceFactory.class, xTextDoc);

      Object oGraphic = null;
      try {
          // Creating the service GraphicObject
          oGraphic =xMSFDoc.createInstance("com.sun.star.text.TextGraphicObject");
      } catch (Exception e) {
          WtMessageHandler.printToLogFile("Could not create instance");
          WtMessageHandler.printException(e);
      }

      // Getting the text
      XText xText = xTextDoc.getText();

      // Getting the cursor on the document (current position)
//      com.sun.star.text.XTextCursor xTextCursor = xText.createTextCursor();
      WtViewCursorTools vCursor = new WtViewCursorTools(xComp);
      XTextCursor xTextCursor = vCursor.getTextCursorBeginn();

      // Querying for the interface XTextContent on the GraphicObject
      XTextContent xTextContent = UnoRuntime.queryInterface(XTextContent.class, oGraphic);

      // Printing information to the log file
      WtMessageHandler.printToLogFile("inserting graphic");
      try {
        // Inserting the content
        xText.insertTextContent(xTextCursor, xTextContent, true);
      } catch (Exception e) {
        WtMessageHandler.printToLogFile("Could not insert Content");
        WtMessageHandler.printException(e);
      }

      // Printing information to the log file
      WtMessageHandler.printToLogFile("adding graphic");

      // Querying for the interface XPropertySet on GraphicObject
     XPropertySet xPropSet = UnoRuntime.queryInterface(XPropertySet.class, oGraphic);
      try {
        // Creating a string for the graphic url
        java.io.File sourceFile = new java.io.File(strUrl);
        StringBuffer sUrl = new StringBuffer("file:///");
        sUrl.append(sourceFile.getCanonicalPath().replace('\\', '/'));
        WtMessageHandler.printToLogFile("insert graphic \"" + sUrl + "\"");

        XGraphicProvider xGraphicProvider = UnoRuntime.queryInterface(XGraphicProvider.class,
                xMCF.createInstanceWithContext("com.sun.star.graphic.GraphicProvider", xContext));

        PropertyValue[] aMediaProps = new PropertyValue[] { new PropertyValue() };
        aMediaProps[0].Name = "URL";
        aMediaProps[0].Value = sUrl.toString();

        XGraphic xGraphic = UnoRuntime.queryInterface(XGraphic.class, xGraphicProvider.queryGraphic(aMediaProps));
        
        // Setting the anchor type
        xPropSet.setPropertyValue("AnchorType",
                   com.sun.star.text.TextContentAnchorType.AT_PARAGRAPH );

        // Setting the graphic url
        xPropSet.setPropertyValue( "Graphic", xGraphic );

        // Setting the horizontal position
//        xPropSet.setPropertyValue( "HoriOrientPosition", Integer.valueOf( 5500 ) );

        // Setting the vertical position
//        xPropSet.setPropertyValue( "VertOrientPosition", Integer.valueOf( 4200 ) );

        // Setting the width
        xPropSet.setPropertyValue("Width", size * SIZE_FACTOR);

        // Setting the height
        xPropSet.setPropertyValue("Height", size * SIZE_FACTOR);
      } catch (Exception e) {
        WtMessageHandler.printToLogFile("Couldn't set property 'GraphicURL'");
        WtMessageHandler.printException(e);
      }
    }
    catch( Exception e ) {
      WtMessageHandler.printException(e);
    }
  }

}
