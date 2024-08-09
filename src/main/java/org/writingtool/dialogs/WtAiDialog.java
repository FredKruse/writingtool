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
package org.writingtool.dialogs;

import java.awt.Color;
import java.awt.Container;
import java.awt.Dimension;
import java.awt.Font;
import java.awt.Frame;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.Image;
import java.awt.Insets;
import java.awt.Point;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.ItemEvent;
import java.awt.event.ItemListener;
import java.awt.event.KeyEvent;
import java.awt.event.KeyListener;
import java.awt.event.WindowEvent;
import java.awt.event.WindowFocusListener;
import java.awt.event.WindowListener;
import java.awt.image.BufferedImage;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.io.PrintWriter;
import java.net.URL;
import java.util.ArrayList;
import java.util.List;
import java.util.ResourceBundle;

import javax.imageio.ImageIO;
import javax.swing.BorderFactory;
import javax.swing.ImageIcon;
import javax.swing.JButton;
import javax.swing.JComboBox;
import javax.swing.JDialog;
import javax.swing.JFileChooser;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JProgressBar;
import javax.swing.JScrollPane;
import javax.swing.JSlider;
import javax.swing.JTabbedPane;
import javax.swing.JTextField;
import javax.swing.JTextPane;
import javax.swing.ToolTipManager;
import javax.swing.event.ChangeEvent;
import javax.swing.event.ChangeListener;

import org.writingtool.WtDocumentCache;
import org.writingtool.WtDocumentsHandler;
import org.writingtool.WtSingleDocument;
import org.writingtool.WtDocumentCache.TextParagraph;
import org.writingtool.WtDocumentsHandler.WaitDialogThread;
import org.writingtool.aisupport.WtAiParagraphChanging;
import org.writingtool.aisupport.WtAiRemote;
import org.writingtool.config.WtConfiguration;
import org.writingtool.tools.WtMessageHandler;
import org.writingtool.tools.WtOfficeDrawTools;
import org.writingtool.tools.WtOfficeGraphicTools;
import org.writingtool.tools.WtOfficeTools;
import org.writingtool.tools.WtViewCursorTools;
import org.writingtool.tools.WtOfficeTools.DocumentType;

import com.sun.star.lang.Locale;
import com.sun.star.lang.XComponent;

/**
 * Dialog to change paragraphs by AI
 * @since 6.5
 * @author Fred Kruse
 */
public class WtAiDialog extends Thread implements ActionListener {
  
  private final static float DEFAULT_TEMPERATURE = 0.7f;
  private final static int DEFAULT_STEP = 30;
  private final static String TEMP_IMAGE_FILE_NAME = "tmpImage.jpg";
  private final static String AI_INSTRUCTION_FILE_NAME = "LT_AI_Instructions.dat";
  private final static int MAX_INSTRUCTIONS = 40;
  private final static int SHIFT1 = 14;
  private final static int dialogWidth = 700;
  private final static int dialogHeight = 750;
  private final static int imageWidth = 512;
//  private final static int imageHeight = 256;

  private boolean debugMode = false;
  private boolean debugModeTm = false;
  
  private final ResourceBundle messages;
  private final JDialog dialog;
  private final Container contentPane;
  private final JButton help; 
  private final JButton close;
  private JProgressBar checkProgress;
  private final Image ltImage;
  
  private final JLabel instructionLabel;
  private final JLabel imgInstructionLabel;
  private final JComboBox<String> instruction;
  private final JLabel paragraphLabel;
  private final JTextPane paragraph;
  private final JLabel resultLabel;
  private final JTextPane result;
  private final JLabel temperatureLabel;
  private final JSlider temperatureSlider;
  private final JButton execute; 
  private final JButton copyResult; 
  private final JButton reset; 
  private final JButton clear; 
  private final JButton undo;
  private final JButton createImage;
  private final JButton overrideParagraph; 
  private final JButton addToParagraph;
  
  private final JTextField imgInstruction;
  private final JLabel excludeLabel;
  private final JTextField exclude;
  private final JLabel imageLabel;
//  private final JFrame imageFrame;
  private final JLabel imageFrame;
  private final JLabel stepLabel;
  private final JSlider stepSlider;
  
  private final JButton changeImage;
  private final JButton newImage;
  private final JButton removeImage;
  private final JButton saveImage;
  private final JButton insertImage;
  
  private final JTabbedPane mainPanel;
  private final JPanel mainImagePanel;
  private final JPanel mainTextPanel;

  private final WtAiParagraphChanging aiParent;
  private WtSingleDocument currentDocument;
  private WtDocumentsHandler documents;
  private WtConfiguration config;
  private DocumentType documentType;
  
  private int dialogX = -1;
  private int dialogY = -1;
  private List<String> instructionList = new ArrayList<>();
  private String saveText;
  private String saveResult;
  private String paraText;
  private String resultText;
  private String imgInstText;
  private String instText;
  private Locale locale;
  private boolean atWork = false;
  private boolean focusLost = false;
  private float temperature = DEFAULT_TEMPERATURE;
  private int seed = randomInteger();
  private int step = DEFAULT_STEP;
  private BufferedImage image;
  private String urlString;

  /**
   * the constructor of the class creates all elements of the dialog
   */
  public WtAiDialog(WtSingleDocument document, WaitDialogThread inf, 
      ResourceBundle messages, WtAiParagraphChanging aiParent) {
    this.messages = messages;
    this.aiParent = aiParent;
    documents = document.getMultiDocumentsHandler();
    config = documents.getConfiguration();
    long startTime = 0;
    if (debugModeTm) {
      startTime = System.currentTimeMillis();
    }
    ltImage = WtOfficeTools.getLtImage();
    if (!documents.isJavaLookAndFeelSet()) {
      documents.setJavaLookAndFeel();
    }
    
    currentDocument = document;
    documentType = document.getDocumentType();
    
    dialog = new JDialog();
    contentPane = dialog.getContentPane();
    instructionLabel = new JLabel(messages.getString("loAiDialogInstructionLabel") + ":");
    instruction = new JComboBox<String>();
    paragraphLabel = new JLabel(messages.getString("loAiDialogParagraphLabel") + ":");
    paragraph = new JTextPane();
    paragraph.setBorder(BorderFactory.createLineBorder(Color.gray));
    resultLabel = new JLabel(messages.getString("loAiDialogResultLabel") + ":");
    result = new JTextPane();
    result.setBorder(BorderFactory.createLineBorder(Color.gray));
    execute = new JButton (messages.getString("loAiDialogExecuteButton")); 
    copyResult = new JButton (messages.getString("loAiDialogcopyResultButton")); 
    reset = new JButton (messages.getString("loAiDialogResetButton")); 
    clear = new JButton (messages.getString("loAiDialogClearButton")); 
    undo = new JButton (messages.getString("loAiDialogUndoButton")); 
    createImage = new JButton (messages.getString("loAiDialogCreateImageButton")); 
    overrideParagraph = new JButton (messages.getString("loAiDialogOverrideButton")); 
    addToParagraph = new JButton (messages.getString("loAiDialogaddToButton")); 
    help = new JButton (messages.getString("loAiDialogHelpButton")); 
    close = new JButton (messages.getString("loAiDialogCloseButton")); 
    
    imgInstructionLabel = new JLabel(messages.getString("loAiDialogInstructionLabel") + ":");
    imgInstruction = new JTextField();
    excludeLabel = new JLabel(messages.getString("loAiDialogImgExcludeLabel") + ":");
    exclude = new JTextField();
    imageLabel = new JLabel(messages.getString("loAiDialogImgImageLabel") + ":");
    imageFrame = new JLabel();
    imageFrame.setSize(imageWidth, imageWidth);
    imageFrame.setBorder(BorderFactory.createLineBorder(Color.gray));
//    imageFrame = new JFrame();
//    imageFrame.setSize(imageWidth, imageHeight);
//    imageFrame.add(image);
    
    changeImage = new JButton (messages.getString("loAiDialogImgChangeImageButton"));
    newImage = new JButton (messages.getString("loAiDialogImgNewImageButton"));
    removeImage = new JButton (messages.getString("loAiDialogImgRemoveButton"));
    saveImage = new JButton (messages.getString("loAiDialogImgSaveButton"));
    insertImage = new JButton (messages.getString("loAiDialogImgInsertButton"));
    
    mainPanel = new JTabbedPane();
    mainImagePanel = new JPanel();
    mainTextPanel = new JPanel();
    
    checkProgress = new JProgressBar(0, 100);

    temperatureLabel = new JLabel(messages.getString("loAiDialogCreativityLabel") + ":");
    
    temperatureSlider = new JSlider(0, 100, (int)(DEFAULT_TEMPERATURE*100));

    stepLabel = new JLabel(messages.getString("loAiDialogImgStepLabel") + ":");
    
    stepSlider = new JSlider(0, 100, DEFAULT_STEP);
    
    checkProgress.setStringPainted(true);
    checkProgress.setIndeterminate(false);
    try {
      if (debugMode) {
        WtMessageHandler.printToLogFile("CheckDialog: LtCheckDialog: LtCheckDialog called");
      }

      
      if (dialog == null) {
        WtMessageHandler.printToLogFile("CheckDialog: LtCheckDialog: LtCheckDialog == null");
      }
      String dialogName = messages.getString("loAiDialogTitle");
      dialog.setName(dialogName);
      dialog.setTitle(dialogName + " (WritingTool " + WtOfficeTools.getWtInformation() + ")");
      dialog.setDefaultCloseOperation(JDialog.DISPOSE_ON_CLOSE);
      ((Frame) dialog.getOwner()).setIconImage(ltImage);

      Font dialogFont = instructionLabel.getFont();
      instructionLabel.setFont(dialogFont);
      imgInstructionLabel.setFont(dialogFont);

      instruction.setFont(dialogFont);
      instruction.setEditable(true);
      instructionList = readInstructions();
      for (String instr : instructionList) {
        instruction.addItem(instr);
      }
      if (!instructionList.isEmpty()) {
        instText = instruction.getItemAt(0);
      }
      if (debugModeTm) {
        long runTime = System.currentTimeMillis() - startTime;
          WtMessageHandler.printToLogFile("CheckDialog: Time to initialise Languages: " + runTime);
          startTime = System.currentTimeMillis();
      }
      if (inf.canceled()) {
        return;
      }

      paragraphLabel.setFont(dialogFont);
      paragraph.setFont(dialogFont);
      JScrollPane paragraphPane = new JScrollPane(paragraph);
      paragraphPane.setMinimumSize(new Dimension(0, 30));

      resultLabel.setFont(dialogFont);
      result.setFont(dialogFont);
      
      temperatureLabel.setFont(dialogFont);
      temperatureSlider.setMajorTickSpacing(10);
      temperatureSlider.setMinorTickSpacing(5);
      temperatureSlider.setPaintTicks(true);
      temperatureSlider.setSnapToTicks(true);
      temperatureSlider.addChangeListener(new ChangeListener( ) {
        @Override
        public void stateChanged(ChangeEvent e) {
          int value = temperatureSlider.getValue();
          temperature = (float) (value / 100.);
        }
      });

      stepLabel.setFont(dialogFont);
      stepSlider.setMajorTickSpacing(10);
      stepSlider.setMinorTickSpacing(5);
      stepSlider.setPaintTicks(true);
      stepSlider.addChangeListener(new ChangeListener( ) {
        @Override
        public void stateChanged(ChangeEvent e) {
          step = stepSlider.getValue();
        }
      });
      stepSlider.addKeyListener(new KeyListener() {
        @Override
        public void keyPressed(KeyEvent e) {
          if(e.getKeyCode() == KeyEvent.VK_ENTER) {
            createImage();
          }
        }
        @Override
        public void keyReleased(KeyEvent e) {
        }
        @Override
        public void keyTyped(KeyEvent e) {
        }
      });
      

      JScrollPane resultPane = new JScrollPane(result);
      resultPane.setMinimumSize(new Dimension(0, 30));
      
      if (debugModeTm) {
        long runTime = System.currentTimeMillis() - startTime;
          WtMessageHandler.printToLogFile("CheckDialog: Time to initialise suggestions, etc.: " + runTime);
          startTime = System.currentTimeMillis();
      }
      if (inf.canceled()) {
        return;
      }
      
      imgInstruction.addKeyListener(new KeyListener() {
        @Override
        public void keyPressed(KeyEvent e) {
          if(e.getKeyCode() == KeyEvent.VK_ENTER) {
            createImage();
          }
        }
        @Override
        public void keyReleased(KeyEvent e) {
        }
        @Override
        public void keyTyped(KeyEvent e) {
          String txt = imgInstruction.getText();
          if (imgInstText == null && !txt.isEmpty()) {
            imgInstText = txt;
            setButtonState(true);
          } else if (imgInstText != null && txt.isEmpty()) {
            imgInstText = null;
            setButtonState(true);
          }
        }
      });
      
      exclude.addKeyListener(new KeyListener() {
        @Override
        public void keyPressed(KeyEvent e) {
          if(e.getKeyCode() == KeyEvent.VK_ENTER) {
            createImage();
          }
        }
        @Override
        public void keyReleased(KeyEvent e) {
        }
        @Override
        public void keyTyped(KeyEvent e) {
        }
      });
      
      JTextField iText = (JTextField) instruction.getEditor().getEditorComponent();
      iText.setSelectionColor(imgInstruction.getSelectionColor());
      
      instruction.getEditor().getEditorComponent().addKeyListener(new KeyListener() {
        @Override
        public void keyPressed(KeyEvent e) {
          if(e.getKeyCode() == KeyEvent.VK_ENTER) {
            instText = (String) instruction.getEditor().getItem();
            if (instText != null) {
              createText();
            }
          }
        }
        @Override
        public void keyReleased(KeyEvent e) {
        }
        @Override
        public void keyTyped(KeyEvent e) {
          String txt = (String) instruction.getEditor().getItem();
//          WtMessageHandler.printToLogFile("keyTyped: Instruction: " + txt);
          if (instText == null && !txt.isEmpty()) {
            instText = txt;
            setButtonState(true);
          } else if (instText != null && txt.isEmpty()) {
            instText = null;
            setButtonState(true);
          }
        }
      });

      instruction.addItemListener(new ItemListener() {
        @Override
        public void itemStateChanged(ItemEvent e) {
//          WtMessageHandler.printToLogFile("itemStateChanged: Instruction: " + instruction.getSelectedItem());
          instText = (String) instruction.getSelectedItem();
          setButtonState(true);
        }
      });
      
      execute.setFont(dialogFont);
      execute.addActionListener(this);
      execute.setActionCommand("execute");
      
      copyResult.setFont(dialogFont);
      copyResult.addActionListener(this);
      copyResult.setActionCommand("copyResult");
      
      reset.setFont(dialogFont);
      reset.addActionListener(this);
      reset.setActionCommand("reset");
      
      clear.setFont(dialogFont);
      clear.addActionListener(this);
      clear.setActionCommand("clear");
      
      undo.setFont(dialogFont);
      undo.addActionListener(this);
      undo.setActionCommand("undo");
      
      createImage.setFont(dialogFont);
      createImage.addActionListener(this);
      createImage.setActionCommand("createImage");
      
      overrideParagraph.setFont(dialogFont);
      overrideParagraph.addActionListener(this);
      overrideParagraph.setActionCommand("overrideParagraph");
      
      addToParagraph.setFont(dialogFont);
      addToParagraph.addActionListener(this);
      addToParagraph.setActionCommand("addToParagraph");
      
      help.setFont(dialogFont);
      help.addActionListener(this);
      help.setActionCommand("help");
      
      close.setFont(dialogFont);
      close.addActionListener(this);
      close.setActionCommand("close");

      changeImage.setFont(dialogFont);
      changeImage.addActionListener(this);
      changeImage.setActionCommand("changeImage");
      
      newImage.setFont(dialogFont);
      newImage.addActionListener(this);
      newImage.setActionCommand("newImage");
      
      removeImage.setFont(dialogFont);
      removeImage.addActionListener(this);
      removeImage.setActionCommand("removeImage");
      
      saveImage.setFont(dialogFont);
      saveImage.addActionListener(this);
      saveImage.setActionCommand("saveImage");
      
      insertImage.setFont(dialogFont);
      insertImage.addActionListener(this);
      insertImage.setActionCommand("insertImage");
      
     dialog.addWindowFocusListener(new WindowFocusListener() {
        @Override
        public void windowGainedFocus(WindowEvent e) {
          if (focusLost) {
            try {
              Point p = dialog.getLocation();
              dialogX = p.x;
              dialogY = p.y;
              if (debugMode) {
                WtMessageHandler.printToLogFile("CheckDialog: LtCheckDialog: Window Focus gained: Event = " + e.paramString());
              }
              setButtonState(!atWork);
              currentDocument = getCurrentDocument();
              if (currentDocument == null) {
                closeDialog();
                return;
              }
              if (debugMode) {
                WtMessageHandler.printToLogFile("CheckDialog: LtCheckDialog: Window Focus gained: new docType = " + currentDocument.getDocumentType());
              }
              setText();
              focusLost = false;
            } catch (Throwable t) {
              WtMessageHandler.showError(t);
              closeDialog();
            }
          }
        }
        @Override
        public void windowLostFocus(WindowEvent e) {
          try {
            if (debugMode) {
              WtMessageHandler.printToLogFile("CheckDialog: LtCheckDialog: Window Focus lost: Event = " + e.paramString());
            }
            setButtonState(false);
            dialog.setEnabled(true);
            focusLost = true;
          } catch (Throwable t) {
            WtMessageHandler.showError(t);
            closeDialog();
          }
        }
      });

      dialog.addWindowListener(new WindowListener() {
        @Override
        public void windowOpened(WindowEvent e) {
        }
        @Override
        public void windowClosing(WindowEvent e) {
          closeDialog();
        }
        @Override
        public void windowClosed(WindowEvent e) {
        }
        @Override
        public void windowIconified(WindowEvent e) {
        }
        @Override
        public void windowDeiconified(WindowEvent e) {
        }
        @Override
        public void windowActivated(WindowEvent e) {
        }
        @Override
        public void windowDeactivated(WindowEvent e) {
        }
      });
      
      if (debugModeTm) {
        long runTime = System.currentTimeMillis() - startTime;
          WtMessageHandler.printToLogFile("CheckDialog: Time to initialise Buttons: " + runTime);
          startTime = System.currentTimeMillis();
      }
      if (inf.canceled()) {
        return;
      }
      
      //  Define Text panels

      //  Define 1. right panel
      JPanel rightPanel1 = new JPanel();
      rightPanel1.setLayout(new GridBagLayout());
      GridBagConstraints cons21 = new GridBagConstraints();
      cons21.insets = new Insets(2, 0, 2, 0);
      cons21.gridx = 0;
      cons21.gridy = 0;
      cons21.anchor = GridBagConstraints.NORTHWEST;
      cons21.fill = GridBagConstraints.BOTH;
      cons21.weightx = 1.0f;
      cons21.weighty = 0.0f;
      cons21.gridy++;
      rightPanel1.add(copyResult, cons21);
      cons21.gridy++;
      rightPanel1.add(reset, cons21);
      cons21.gridy++;
      rightPanel1.add(clear, cons21);

      //  Define 2. right panel
      JPanel rightPanel2 = new JPanel();
      rightPanel2.setLayout(new GridBagLayout());
      GridBagConstraints cons22 = new GridBagConstraints();
      cons22.insets = new Insets(2, 0, 2, 0);
      cons22.gridx = 0;
      cons22.gridy = 0;
      cons22.anchor = GridBagConstraints.NORTHWEST;
      cons22.fill = GridBagConstraints.BOTH;
      cons22.weightx = 1.0f;
      cons22.weighty = 0.0f;
      cons22.gridy++;
      cons22.gridy++;
      rightPanel2.add(undo, cons22);
      cons22.gridy++;
      rightPanel2.add(createImage, cons22);
      cons22.gridy++;
      rightPanel2.add(addToParagraph, cons22);
      cons22.gridy++;
      rightPanel2.add(overrideParagraph, cons22);
      
      //  Define main text panel
      mainTextPanel.setLayout(new GridBagLayout());
      GridBagConstraints cons1 = new GridBagConstraints();
      cons1.insets = new Insets(SHIFT1, 4, 4, 4);
      cons1.gridx = 0;
      cons1.gridy = 0;
      cons1.anchor = GridBagConstraints.NORTHWEST;
      cons1.fill = GridBagConstraints.BOTH;
      cons1.weightx = 1.0f;
      cons1.weighty = 0.0f;
      mainTextPanel.add(instructionLabel, cons1);
      cons1.insets = new Insets(4, 4, 4, 4);
      cons1.gridy++;
      mainTextPanel.add(instruction, cons1);
      cons1.weightx = 0.0f;
      cons1.gridx++;
      mainTextPanel.add(execute, cons1);
      cons1.insets = new Insets(SHIFT1, 4, 4, 4);
      cons1.gridx = 0;
      cons1.gridy++;
      cons1.weightx = 1.0f;
      cons1.weighty = 0.0f;
      mainTextPanel.add(paragraphLabel, cons1);
      cons1.insets = new Insets(4, 4, 4, 4);
      cons1.gridy++;
      cons1.weighty = 2.0f;
      mainTextPanel.add(paragraphPane, cons1);
      cons1.gridx++;
      cons1.weightx = 0.0f;
      cons1.weighty = 0.0f;
      mainTextPanel.add(rightPanel1, cons1);
      cons1.insets = new Insets(SHIFT1, 4, 4, 4);
      cons1.gridx = 0;
      cons1.gridy++;
      cons1.weightx = 1.0f;
      mainTextPanel.add(resultLabel, cons1);
      cons1.insets = new Insets(4, 4, 4, 4);
      cons1.gridy++;
      cons1.weighty = 2.0f;
      mainTextPanel.add(resultPane, cons1);
      cons1.gridx++;
      cons1.weightx = 0.0f;
      cons1.weighty = 0.0f;
      mainTextPanel.add(rightPanel2, cons1);
      cons1.insets = new Insets(SHIFT1, 4, 4, 4);
      cons1.gridx = 0;
      cons1.gridy++;
      mainTextPanel.add(temperatureLabel, cons1);
      cons1.insets = new Insets(4, 4, 4, 4);
      cons1.gridy++;
      mainTextPanel.add(temperatureSlider, cons1);
      
      //  Define Image panels

      //  Define 1. left panel
      JPanel leftPanel1 = new JPanel();
      leftPanel1.setLayout(new GridBagLayout());
      GridBagConstraints cons11 = new GridBagConstraints();
      cons11.gridx = 0;
      cons11.gridy = 0;
      cons11.anchor = GridBagConstraints.NORTHWEST;
      cons11.fill = GridBagConstraints.BOTH;
      cons11.weightx = 1.0f;
      cons11.weighty = 0.0f;
      cons11.insets = new Insets(SHIFT1, 0, 4, 0);
      cons11.gridy++;
      leftPanel1.add(imgInstructionLabel, cons11);
      cons11.insets = new Insets(4, 0, 4, 0);
      cons11.gridy++;
      leftPanel1.add(imgInstruction, cons11);
      cons11.insets = new Insets(SHIFT1, 0, 4, 0);
      cons11.gridy++;
      leftPanel1.add(excludeLabel, cons11);
      cons11.insets = new Insets(4, 0, 4, 0);
      cons11.gridy++;
      leftPanel1.add(exclude, cons11);

      //  Define 1. right panel
      rightPanel1 = new JPanel();
      rightPanel1.setLayout(new GridBagLayout());
      cons21 = new GridBagConstraints();
      cons21.insets = new Insets(2, 0, 2, 0);
      cons21.gridx = 0;
      cons21.gridy = 0;
      cons21.anchor = GridBagConstraints.NORTHWEST;
      cons21.fill = GridBagConstraints.BOTH;
      cons21.weightx = 1.0f;
      cons21.weighty = 0.0f;
      cons21.gridy++;
      rightPanel1.add(changeImage, cons21);
      cons21.gridy++;
      rightPanel1.add(newImage, cons21);

      //  Define 2. right panel
      rightPanel2 = new JPanel();
      rightPanel2.setLayout(new GridBagLayout());
      cons22 = new GridBagConstraints();
      cons22.insets = new Insets(2, 0, 2, 0);
      cons22.gridx = 0;
      cons22.gridy = 0;
      cons22.anchor = GridBagConstraints.NORTHWEST;
      cons22.fill = GridBagConstraints.BOTH;
      cons22.weightx = 1.0f;
      cons22.weighty = 0.0f;
      cons22.gridy++;
      cons22.gridy++;
      rightPanel2.add(removeImage, cons22);
      cons22.gridy++;
      rightPanel2.add(saveImage, cons22);
      cons22.gridy++;
      rightPanel2.add(insertImage, cons22);
      
      //  Define main panel
      mainImagePanel.setLayout(new GridBagLayout());
      cons1 = new GridBagConstraints();
      cons1.insets = new Insets(4, 4, 4, 4);
      cons1.gridx = 0;
      cons1.gridy = 0;
      cons1.anchor = GridBagConstraints.NORTHWEST;
      cons1.fill = GridBagConstraints.BOTH;
      cons1.weightx = 1.0f;
      cons1.weighty = 0.0f;
      mainImagePanel.add(leftPanel1, cons1);
      cons1.gridx++;
      cons1.weightx = 0.0f;
      cons1.weighty = 0.0f;
      mainImagePanel.add(rightPanel1, cons1);
      cons1.insets = new Insets(SHIFT1, 4, 4, 4);
      cons1.gridx = 0;
      cons1.gridy++;
      cons1.weightx = 1.0f;
      mainImagePanel.add(imageLabel, cons1);
      cons1.insets = new Insets(4, 4, 4, 4);
      cons1.gridy++;
      cons1.weighty = 2.0f;
      mainImagePanel.add(new JScrollPane(imageFrame), cons1);
      cons1.gridx++;
      cons1.weightx = 0.0f;
      cons1.weighty = 0.0f;
      mainImagePanel.add(rightPanel2, cons1);
      cons1.insets = new Insets(SHIFT1, 4, 4, 4);
      cons1.gridx = 0;
      cons1.gridy++;
      mainImagePanel.add(stepLabel, cons1);
      cons1.insets = new Insets(4, 4, 4, 4);
      cons1.gridy++;
      mainImagePanel.add(stepSlider, cons1);
      
      //  Define tabbed main pane
      if (config.useAiSupport()) {
        mainPanel.add(messages.getString("guiAiText"), mainTextPanel);
      }
      if (config.useAiImgSupport()) {
        mainPanel.add(messages.getString("guiAiImages"), mainImagePanel);
      }

      //  Define general button panel
      JPanel generalButtonPanel = new JPanel();
      generalButtonPanel.setLayout(new GridBagLayout());
      GridBagConstraints cons3 = new GridBagConstraints();
      cons3.insets = new Insets(4, 4, 4, 4);
      cons3.gridx = 0;
      cons3.gridy = 0;
      cons3.anchor = GridBagConstraints.WEST;
//      cons3.fill = GridBagConstraints.HORIZONTAL;
      cons3.weightx = 1.0f;
      cons3.weighty = 0.0f;
      generalButtonPanel.add(help, cons3);
      cons3.anchor = GridBagConstraints.EAST;
      cons3.gridx++;
      generalButtonPanel.add(close, cons3);
      
      //  Define check progress panel
      JPanel checkProgressPanel = new JPanel();
      checkProgressPanel.setLayout(new GridBagLayout());
      GridBagConstraints cons4 = new GridBagConstraints();
      cons4.insets = new Insets(4, 4, 4, 4);
      cons4.gridx = 0;
      cons4.gridy = 0;
      cons4.anchor = GridBagConstraints.NORTHWEST;
      cons4.fill = GridBagConstraints.HORIZONTAL;
      cons4.weightx = 4.0f;
      cons4.weighty = 0.0f;
      checkProgressPanel.add(checkProgress, cons4);

      contentPane.setLayout(new GridBagLayout());
      GridBagConstraints cons = new GridBagConstraints();
      cons.insets = new Insets(8, 8, 8, 8);
      cons.gridx = 0;
      cons.gridy = 0;
      cons.anchor = GridBagConstraints.NORTHWEST;
      cons.fill = GridBagConstraints.BOTH;
      cons.weightx = 1.0f;
      cons.weighty = 1.0f;
      contentPane.add(mainPanel, cons);
      cons.gridy++;
      cons.weighty = 0.0f;
      contentPane.add(generalButtonPanel, cons);
      cons.gridy++;
      contentPane.add(checkProgressPanel, cons);

      if (debugModeTm) {
        long runTime = System.currentTimeMillis() - startTime;
//        if (runTime > OfficeTools.TIME_TOLERANCE) {
          WtMessageHandler.printToLogFile("CheckDialog: Time to initialise panels: " + runTime);
//        }
          startTime = System.currentTimeMillis();
      }
      if (inf.canceled()) {
        return;
      }

      dialog.pack();
      // center on screen:
      Dimension screenSize = Toolkit.getDefaultToolkit().getScreenSize();
      Dimension frameSize = new Dimension(dialogWidth, dialogHeight);
      dialog.setSize(frameSize);
      dialog.setLocation(screenSize.width / 2 - frameSize.width / 2,
          screenSize.height / 2 - frameSize.height / 2);
      dialog.setLocationByPlatform(true);
      
      ToolTipManager.sharedInstance().setDismissDelay(30000);
      if (debugModeTm) {
        long runTime = System.currentTimeMillis() - startTime;
//        if (runTime > OfficeTools.TIME_TOLERANCE) {
          WtMessageHandler.printToLogFile("CheckDialog: Time to initialise dialog size: " + runTime);
//        }
      }
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
      closeDialog();
    }
  }
  
  @Override
  public void run() {
    try {
      show();
    } catch (Throwable e) {
      WtMessageHandler.showError(e);
    }
  }

  /**
   * show the dialog
   * @throws Throwable 
   */
  public void show() throws Throwable {
    if (currentDocument == null || (currentDocument.getDocumentType() != DocumentType.WRITER 
          && currentDocument.getDocumentType() != DocumentType.IMPRESS)) {
      return;
    }
    if (debugMode) {
      WtMessageHandler.printToLogFile("CheckDialog: show: Goto next Error");
    }
    if (dialogX < 0 || dialogY < 0) {
      Dimension screenSize = Toolkit.getDefaultToolkit().getScreenSize();
      Dimension frameSize = dialog.getSize();
      dialogX = screenSize.width / 2 - frameSize.width / 2;
      dialogY = screenSize.height / 2 - frameSize.height / 2;
    }
    dialog.setLocation(dialogX, dialogY);
    dialog.setAutoRequestFocus(true);
    dialog.setVisible(true);
    setText();
    setButtonState(true);
  }
  
  public void toFront() {
    dialog.setVisible(true);
    dialog.toFront();
  }

  /**
   * Get the current document
   * Wait until it is initialized (by LO/OO)
   */
  private WtSingleDocument getCurrentDocument() {
    WtSingleDocument currentDocument = documents.getCurrentDocument();
    if (currentDocument == null || (currentDocument.getDocumentType() != DocumentType.WRITER 
        && currentDocument.getDocumentType() != DocumentType.IMPRESS)) {
      return null;
    }
    documentType = currentDocument.getDocumentType();
    return currentDocument;
  }

  /**
   * Initialize the cursor / define the range for check
   * @throws Throwable 
   */
  private void setText() throws Throwable {
    if (currentDocument.getDocumentType() == DocumentType.WRITER) {
      XComponent xComponent = currentDocument.getXComponent();
      WtViewCursorTools viewCursor = new WtViewCursorTools(xComponent);
      TextParagraph tPara = viewCursor.getViewCursorParagraph();
      WtDocumentCache docCache = currentDocument.getDocumentCache();
      paraText = docCache.getTextParagraph(tPara);
      locale = docCache.getTextParagraphLocale(tPara);
      paragraph.setText(paraText);
    } else {
      XComponent xComponent = currentDocument.getXComponent();
      paraText = "";
      locale = WtOfficeDrawTools.getDocumentLocale(xComponent);
      paragraph.setText(paraText);
    }
  }

  /**
   * Initial button state
   */
  private void setAtWorkState(boolean work) {
    checkProgress.setIndeterminate(work);
    atWork = work;
  }
  
  /**
   * Initial button state
   */
  private void setButtonState(boolean enabled) {
    boolean isImpress = documentType == DocumentType.IMPRESS;
    instruction.setEnabled(enabled);
    paragraph.setEnabled(enabled);
    result.setEnabled(enabled);
//    String instructionText = (String) instruction.getSelectedItem();
    execute.setEnabled(instText == null || instText.isEmpty() ? false : enabled);
//    execute.setEnabled(enabled);
    copyResult.setEnabled(resultText == null  || resultText.isEmpty() ? false : enabled);
    reset.setEnabled(isImpress ? false : enabled);
    clear.setEnabled(paraText == null || paraText.isEmpty() ? false : enabled);
    undo.setEnabled(saveText == null ? false : enabled);
    createImage.setEnabled(resultText == null  || resultText.isEmpty() || !config.useAiImgSupport() ? false : enabled);
    overrideParagraph.setEnabled(resultText == null || resultText.isEmpty() || isImpress ? false : enabled);
    addToParagraph.setEnabled(resultText == null  || resultText.isEmpty() ? false : enabled);
    help.setEnabled(enabled);
    close.setEnabled(true);

    String instructionText = imgInstruction.getText();
    imgInstruction.setEnabled(enabled);
    exclude.setEnabled(enabled);
    imageFrame.setEnabled(enabled);
    changeImage.setEnabled(instructionText == null || instructionText.isEmpty() ? false : enabled);
    newImage.setEnabled(instructionText == null || instructionText.isEmpty() ? false : enabled);
    removeImage.setEnabled(image == null ? false : enabled);
    saveImage.setEnabled(image == null ? false : enabled);
    insertImage.setEnabled(image == null ? false : enabled);
    
    contentPane.revalidate();
    contentPane.repaint();
    dialog.setEnabled(enabled);
  }
  
  /**
   * execute AI request
   */
  private void createText() {
    try {
      setAtWorkState(true);
      setButtonState(false);
      if (!documents.isEnoughHeapSpace()) {
        closeDialog();
        return;
      }
      if (debugMode) {
        WtMessageHandler.printToLogFile("AiDialog: execute: start AI request");
      }
//      instText = (String) instruction.getSelectedItem();
      if (instText == null || instText.isBlank()) {
        return;
      }
      instText = instText.trim();
      if (instructionList.contains(instText)) {
        instructionList.remove(instText);
        instruction.removeItem(instText);
      }
      instruction.insertItemAt(instText, 0);
      instructionList.add(0, instText);
      if (instructionList.size() > MAX_INSTRUCTIONS) {
        instructionList.remove(instructionList.size() - 1);
        instruction.removeItemAt(instruction.getItemCount() - 1);
      }
      writeInstructions(instructionList);
      String text = paragraph.getText();
      WtAiRemote aiRemote = new WtAiRemote(documents, config);
      if (debugMode) {
        WtMessageHandler.printToLogFile("AiParagraphChanging: runInstruction: instruction: " + instText + ", text: " + text);
      }
      String output = aiRemote.runInstruction(instText, text, temperature, 0, locale, false);
      if (debugMode) {
        WtMessageHandler.printToLogFile("AiParagraphChanging: runAiChangeOnParagraph: output: " + output);
      }
      instruction.setSelectedIndex(0);
      result.setEnabled(true);
      result.setText(output);
      resultText = output;
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
      closeDialog();
    } finally {
      setAtWorkState(false);
      setButtonState(true);
    }
  }

  /**
   * execute AI request
   */
  private void createImage() {
    try {
      if (!documents.isEnoughHeapSpace()) {
        closeDialog();
        return;
      }
      if (debugMode) {
        WtMessageHandler.printToLogFile("AiDialog: execute image: start AI request");
      }
      imgInstText = imgInstruction.getText();
      String excludeText = exclude.getText();
      if (excludeText == null) {
        excludeText = "";
      }
      setAtWorkState(true);
      setButtonState(false);
      WtAiRemote aiRemote = new WtAiRemote(documents, config);
      if (debugMode) {
        WtMessageHandler.printToLogFile("AiParagraphChanging: runInstruction: instruction: " 
              + imgInstText + ", exclude: " + excludeText);
      }
      urlString = aiRemote.runImgInstruction(imgInstText, excludeText, step, seed, imageWidth);
      if (debugMode) {
        WtMessageHandler.printToLogFile("AiParagraphChanging: runAiChangeOnParagraph: url: " + urlString);
      }
      image = getImageFromUrl(urlString);
      imageFrame.setIcon(new ImageIcon(image));
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
      closeDialog();
    }
    setAtWorkState(false);
    setButtonState(true);
  }

  /**
   * Actions of buttons
   */
  @Override
  public void actionPerformed(ActionEvent action) {
    if (!atWork) {
      try {
        if (debugMode) {
          WtMessageHandler.printToLogFile("CheckDialog: actionPerformed: Action: " + action.getActionCommand());
        }
        if (action.getActionCommand().equals("close")) {
          closeDialog();
        } else if (action.getActionCommand().equals("help")) {
          WtMessageHandler.showMessage("Not implemented yet");
        } else if (action.getActionCommand().equals("execute")) {
          createText();
        } else if (action.getActionCommand().equals("createImage")) {
          createImageFromText();
        } else if (action.getActionCommand().equals("changeImage")) {
          createImage();
        } else if (action.getActionCommand().equals("newImage")) {
          seed = randomInteger();
          createImage();
        } else {
          if (action.getActionCommand().equals("copyResult")) {
            copyResult();
          } else if (action.getActionCommand().equals("reset")) {
            resetText();
          } else if (action.getActionCommand().equals("clear")) {
            clearText();
          } else if (action.getActionCommand().equals("undo")) {
            undo();
          } else if (action.getActionCommand().equals("overrideParagraph")) {
            writeToParagraph(true);
          } else if (action.getActionCommand().equals("addToParagraph")) {
            writeToParagraph(false);
          } else if (action.getActionCommand().equals("saveImage")) {
            saveImage();
          } else if (action.getActionCommand().equals("insertImage")) {
            insertImage();
          } else if (action.getActionCommand().equals("removeImage")) {
            removeImage();
          } else {
            WtMessageHandler.showMessage("Action '" + action.getActionCommand() + "' not supported");
          }
          setButtonState(true);
        }
      } catch (Throwable e) {
        WtMessageHandler.showError(e);
        closeDialog();
      }
    }
  }
  
  private void copyResult() {
    saveText = paraText;
    saveResult = resultText;
    paraText = resultText;
    paragraph.setText(paraText);
  }

  private void resetText() throws Throwable {
    saveText = paraText;
    saveResult = resultText;
    setText();
    setButtonState(true);
  }

  private void clearText() throws Throwable {
    saveText = paraText;
    saveResult = resultText;
    paraText = "";
    paragraph.setText(paraText);
    setButtonState(true);
  }

  private void undo() {
    if (saveText != null) {
      paraText = saveText;
      resultText = saveResult;
      saveText = null;
      paragraph.setText(paraText);
      result.setText(resultText);
      setButtonState(true);
    }
  }

  private void createImageFromText() {
    resultText = result.getText();
    String parsedText = resultText.replace("\n", "\r").replace("\r", " ").replace("\"", "\\\"").replace("\t", " ");
    imgInstruction.setText(parsedText);
    exclude.setText("");
    createImage();
    mainPanel.setSelectedComponent(mainImagePanel);
  }

  private void writeToParagraph(boolean override) throws Throwable {
    if (documentType == DocumentType.WRITER) {
      WtAiParagraphChanging.insertText(resultText, currentDocument.getXComponent(), override);
    } else {
      WtOfficeGraphicTools.insertDrawText(resultText, 256, 128, currentDocument.getXComponent());
    }
  }
  
  
  
  private List<String> readInstructions() {
    String dir = WtOfficeTools.getLOConfigDir().getAbsolutePath();
    File file = new File(dir, AI_INSTRUCTION_FILE_NAME);
    if (!file.canRead() || !file.isFile()) {
      return new ArrayList<>();
    }
    List<String> instructions = new ArrayList<>();
    BufferedReader in = null;
    try {
      in = new BufferedReader(new FileReader(file.getAbsoluteFile()));
      String row = null;
      int n = 0;
      while ((row = in.readLine()) != null) {
        if (!row.isBlank() && !instructions.contains(row)) {
          instructions.add(row);
          n++;
          if (n > MAX_INSTRUCTIONS) {
            break;
          }
        }
      }
    } catch (Throwable e) {
      WtMessageHandler.showError(e);
    } finally {
      if (in != null)
        try {
            in.close();
        } catch (IOException e) {
        }
    }
    return instructions;
  }

  private void writeInstructions(List<String> instructions) {
    String dir = WtOfficeTools.getLOConfigDir().getAbsolutePath();
    File file = new File(dir, AI_INSTRUCTION_FILE_NAME);
    PrintWriter pWriter = null;
    try {
      pWriter = new PrintWriter(new FileWriter(file.getAbsoluteFile()));
      for (String inst : instructions) {
        pWriter.println(inst);
      }
    } catch (Throwable e) {
      WtMessageHandler.showError(e);
    } finally {
      if (pWriter != null) {
        pWriter.flush();
        pWriter.close();
      }
    }
  }

  /**
   * closes the dialog
   */
  public void closeDialog() {
    dialog.setVisible(false);
    aiParent.setCloseAiDialog();
    if (debugMode) {
      WtMessageHandler.printToLogFile("AiDialog: closeDialog: Close AI Dialog");
    }
    atWork = false;
  }
  
  private BufferedImage getImageFromUrl(String urlString) {
    try {
      URL url = new URL(urlString);
      return ImageIO.read(url);
    } catch (Throwable e) {
      WtMessageHandler.showError(e);
      return null;
    }
    
  }
  
  private static int randomInteger() {
    return (int) (Math.random() * Integer.MAX_VALUE);
  }
  
  private void saveImage() {
    JFileChooser fileChooser = new JFileChooser();
    fileChooser.setDialogTitle(messages.getString("loAiDialogImgSaveTitle"));   
     
    int userSelection = fileChooser.showSaveDialog(dialog);
     
    if (userSelection == JFileChooser.APPROVE_OPTION) {
      File fileToSave = fileChooser.getSelectedFile();
      saveImage(fileToSave);
    }
  }
  
  private void saveImage(File file) {
    try {
      String name = file.getName();
      String extension = name.substring(name.lastIndexOf('.') + 1);
      if (!extension.equals("png") && !extension.equals("jpg") && !extension.equals("gif")) {
        WtMessageHandler.showMessage(messages.getString("loAiDialogImgSaveErrorMsg"));
        return;
      }
      ImageIO.write(image, extension, file);
    } catch (IOException e) {
      WtMessageHandler.showError(e);
    }
  }
  
  private void insertImage() {
    if (urlString == null) {
      return;
    }
    String extension = urlString.substring(urlString.lastIndexOf('.') + 1);
    if (!extension.equals("png") && !extension.equals("jpg") && !extension.equals("gif")) {
      return;
    }
    File dir = WtOfficeTools.getCacheDir();
    File tmpFile = new File(dir, TEMP_IMAGE_FILE_NAME);
    saveImage(tmpFile);
    if (documentType == DocumentType.IMPRESS) {
      WtOfficeGraphicTools.insertGraphicInImpress(tmpFile.getAbsolutePath(), imageWidth,
        currentDocument.getXComponent(), documents.getContext());
    } else {
      WtOfficeGraphicTools.insertGraphic(tmpFile.getAbsolutePath(), imageWidth, 
        currentDocument.getXComponent(), documents.getContext());
    }
  }
  
  private void removeImage() {
    image = null;
    imageFrame.setIcon(null);
    setButtonState(true);
  }
  
  

}
