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

import java.io.File;

import com.sun.star.awt.Size;
import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.XPropertySet;
import com.sun.star.container.XNameContainer;
import com.sun.star.drawing.XShape;
import com.sun.star.graphic.XGraphic;
import com.sun.star.graphic.XGraphicProvider;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.lang.XMultiServiceFactory;
import com.sun.star.text.ControlCharacter;
import com.sun.star.text.TextContentAnchorType;
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

  public static void addImageLink(XTextDocument doc, XTextCursor cursor, String fnm, XComponentContext xContext) {
    addImageLink(doc, cursor, fnm, 0, 0, xContext);  // 0, 0 means use image's size as width & height
  }


  public static void addImageLink(XTextDocument doc, XTextCursor cursor, String fnm, int width, int height, XComponentContext xContext) { 
    try {

      // create TextContent for graphic
      if (xContext == null) {
        return;
      }
      XMultiComponentFactory xMCF = UnoRuntime.queryInterface(XMultiComponentFactory.class, xContext.getServiceManager());
      if (xMCF == null) {
        return;
      }
      Object cont = xMCF.createInstanceWithContext("com.sun.star.text.TextGraphicObject", xContext);
      if (cont == null) {
        return;
      }
      XTextContent tgo = UnoRuntime.queryInterface(XTextContent.class, cont);
      if (tgo == null) {
        WtMessageHandler.printToLogFile("Could not create a text graphic object");
        return;
      }

      // set anchor and URL properties
      String urlString = fnmToURL(fnm);
      XPropertySet props = UnoRuntime.queryInterface(XPropertySet.class, tgo);
      props.setPropertyValue("AnchorType", TextContentAnchorType.AS_CHARACTER);
      props.setPropertyValue("GraphicURL", urlString);

      // optionally set the width and height
      if ((width > 0) && (height > 0)) {
        props.setPropertyValue("Width", width);
        props.setPropertyValue("Height", height);
      }

      // append image to document, followed by a newline
        append(cursor, tgo);
        endLine(cursor);
    } catch (Throwable e) {
      WtMessageHandler.printException(e);
    }
  }  // end of addImageLink()
  
  public static void addImageShape(XTextDocument doc, XTextCursor cursor, String fnm, XComponent xComponent, XComponentContext xContext) {
    addImageShape(doc, cursor, fnm, 0, 0, xComponent, xContext);  /* 0, 0 means that the method must calculate the image's size */
  }

  public static void addImageShape(XTextDocument doc, XTextCursor cursor, String fnm, int width, int height, 
      XComponent xComponent, XComponentContext xContext) {
    try {
      WtMessageHandler.printToLogFile("addImageShape: " + fnm);
      Size imSize;
      if ((width > 0) && (height > 0))
        imSize = new Size(width, height);
      else {
        imSize = getSize100mm(fnm, xContext);
      if (imSize == null)
        WtMessageHandler.printToLogFile("image size == null");
        return;
      }
      WtMessageHandler.printToLogFile("image size: (" + imSize.Width + "/" + imSize.Height + ")");

      // create TextContent for the graphic shape
      if (xContext == null) {
        WtMessageHandler.printToLogFile("xContext == null");
        return;
      }
      XMultiServiceFactory xMSF = UnoRuntime.queryInterface(XMultiServiceFactory.class, xComponent);
      if (xMSF == null) {
        WtMessageHandler.printToLogFile("XMultiServiceFactory == null");
        return;
      }
      Object o = xMSF.createInstance("com.sun.star.drawing.GraphicObjectShape");
      XTextContent gos = UnoRuntime.queryInterface(XTextContent.class, o);
      if (gos == null) {
        WtMessageHandler.printToLogFile("Could not create a graphic shape");
        return;
      }
      
      // store the image's bitmap in the "GraphicURL" property
      String bitmap = getBitmap(fnm, xComponent);
      WtMessageHandler.printToLogFile("Bitmap size: " + bitmap.length());
      
      XPropertySet propSet =  UnoRuntime.queryInterface(XPropertySet.class, gos);
      propSet.setPropertyValue("GraphicURL", bitmap);
      
      // set the shape's size
      XShape xDrawShape = UnoRuntime.queryInterface(XShape.class, gos);
      xDrawShape.setSize(imSize);  // must be set, or image is tiny
      
      // insert image shape into the document, followed by newline
      append(cursor, gos);
      endLine(cursor);
    } catch(Throwable e) {
      WtMessageHandler.printToLogFile("Insert of \"" + fnm + "\" failed: " + e);
    }
  }  // end of addImageShape()
  
  public static XGraphic loadGraphicFile(String imFnm, XComponentContext xContext)
  {
    try {
      if (xContext == null) {
        WtMessageHandler.printToLogFile("xContext == null");
        return null;
      }
      XMultiComponentFactory xMCF = UnoRuntime.queryInterface(XMultiComponentFactory.class, xContext.getServiceManager());
      if (xMCF == null) {
        WtMessageHandler.printToLogFile("xMCF == null");
        return null;
      }
      Object provider = xMCF.createInstanceWithContext("com.sun.star.graphic.GraphicProvider", xContext);
      if (provider == null) {
        WtMessageHandler.printToLogFile("provider == null");
        return null;
      }
      XGraphicProvider gProvider = UnoRuntime.queryInterface(XGraphicProvider.class, provider);
      if (gProvider == null) {
        WtMessageHandler.printToLogFile("Graphic Provider could not be found");
        return null;
      }
  
      PropertyValue[] fileProps =  makeProps("URL", fnmToURL(imFnm));
      return gProvider.queryGraphic(fileProps);
    } catch(Throwable t) {
      WtMessageHandler.printException(t);
      return null;  
    }
  }  // end of loadGraphicFile()

  private static Size getSize100mm(String imFnm, XComponentContext xContext) throws Throwable
  {
    XGraphic graphic = loadGraphicFile(imFnm, xContext);
    if (graphic == null) {
      WtMessageHandler.printToLogFile("graphic == null");
      return null;
    }
    XPropertySet propSet =  UnoRuntime.queryInterface(XPropertySet.class, graphic);
    return (Size) propSet.getPropertyValue("Size100thMM");
  }  // end of getSize100mm()

  private static String fnmToURL(String fnm)
  // convert a file path to URL format
  {
     try {
       StringBuffer sb = null;
       String path = new File(fnm).getCanonicalPath();
       sb = new StringBuffer("file:///");
       sb.append(path.replace('\\', '/'));
       return sb.toString();
     }
     catch (Throwable e) {
       WtMessageHandler.printToLogFile("Could not access: " + fnm);
       WtMessageHandler.printException(e);
       return null;
     }
  } // end of fnmToURL()
  
  private static PropertyValue[] makeProps(String oName, Object oValue) throws Throwable
  {
    PropertyValue[] props = new PropertyValue[1];
    props[0] = new PropertyValue();
    props[0].Name = oName;
    props[0].Value = oValue;
    return props;
  }  // end of makeProps()

  private static String getBitmap(String fnm, XComponent xComponent)
  // load the graphic as a bitmap, and return it as a string
  {
    try {
      XMultiServiceFactory xMSF = UnoRuntime.queryInterface(XMultiServiceFactory.class, xComponent);
      if (xMSF == null) {
        WtMessageHandler.printToLogFile("XMultiServiceFactory == null");
        return null;
      }
      Object o = xMSF.createInstance("com.sun.star.drawing.BitmapTable");
      XNameContainer bitmapContainer =  UnoRuntime.queryInterface(XNameContainer.class, o);
      if (bitmapContainer == null) {
        WtMessageHandler.printToLogFile("Could not create bitmap container");
        return null;
      }
      // insert image into container
      File f = new File(fnm);
      if (!f.exists() || !f.isFile() || !f.canRead()) {
        WtMessageHandler.printToLogFile("Can' t open file: " + fnm);
        return null;
      }
      String picURL = fnmToURL(fnm);
      if (picURL == null) {
        WtMessageHandler.printToLogFile("Url is null for: " + fnm);
        return null;
      }
      bitmapContainer.insertByName(fnm, picURL);
              // use the filename as the name of the bitmap

      // return the bitmap as a string
      return new String((String) bitmapContainer.getByName(fnm));
    } catch(Throwable e) {
      WtMessageHandler.printToLogFile("Could not create a bitmap container for " + fnm);
      WtMessageHandler.printException(e);
      return null;
    }
  }  // end of getBitmap()

  private static int append(XTextCursor cursor, XTextContent textContent) throws IllegalArgumentException
  {
    XText xText = cursor.getText();
    xText.insertTextContent(cursor, textContent, false);
    cursor.gotoEnd(false);
    return getPosition(cursor);
  }
  
  public static int append(XTextCursor cursor, short ctrlChar) throws IllegalArgumentException
  {
    XText xText = cursor.getText();
    xText.insertControlCharacter(cursor, ctrlChar, false);
    cursor.gotoEnd(false);
    return getPosition(cursor);
  }

  public static int getPosition(XTextCursor cursor) {
    return (cursor.getText().getString()).length();
  }

  public static void endLine(XTextCursor cursor) throws IllegalArgumentException {
    append(cursor, ControlCharacter.LINE_BREAK);
  }

}
