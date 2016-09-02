/*
 * SimpleP2View.java
 */

package simplep2;

import java.awt.Color;
import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JTextArea;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.jdesktop.application.Action;
import org.jdesktop.application.ResourceMap;
import org.jdesktop.application.SingleFrameApplication;
import org.jdesktop.application.FrameView;
import org.jdesktop.application.TaskMonitor;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FilenameFilter;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.util.Calendar;
import java.util.Properties;
import javax.swing.Timer;
import javax.swing.Icon;
import javax.swing.JDialog;
import javax.swing.JFrame;
import javax.swing.UIManager;
import javax.swing.UnsupportedLookAndFeelException;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFHyperlink;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;


/**
 * The application's main frame.
 */
public class SimpleP2View extends FrameView {
    private String FileName;
    private HSSFCell cell;
    private HSSFHyperlink link;
    private int issuenumber;
    private String customerfile;
    private int ii;
    private String custname;
    private String[] mySplitResult;
    private int selNum;
    private String defdir;
    private File defdirf;
    private String Dir1;
    private String Dir2;
    private boolean success;
    private String Dir3;
    private String Dir4;
    private HSSFRow row;
    private byte[] data;
    private String tekst;
    private String tekst2;
    private String filename;
    private JTextArea jTextArea1;
    private String helptekst;

    public SimpleP2View(SingleFrameApplication app) {
        super(app);
 
        initComponents();
        jTextField4.setText("Initiating...");
        load_ini_setting();
        load_customers();
        jComboBox1.setSelectedIndex(0);
        jTextField4.setText("Ready");
        
        // status bar initialization - message timeout, idle icon and busy animation, etc
        ResourceMap resourceMap = getResourceMap();
        int messageTimeout = resourceMap.getInteger("StatusBar.messageTimeout");
        messageTimer = new Timer(messageTimeout, new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                statusMessageLabel.setText("");
            }
        });
        messageTimer.setRepeats(false);
        int busyAnimationRate = resourceMap.getInteger("StatusBar.busyAnimationRate");
        for (int i = 0; i < busyIcons.length; i++) {
            busyIcons[i] = resourceMap.getIcon("StatusBar.busyIcons[" + i + "]");
        }
        busyIconTimer = new Timer(busyAnimationRate, new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                busyIconIndex = (busyIconIndex + 1) % busyIcons.length;
                statusAnimationLabel.setIcon(busyIcons[busyIconIndex]);
            }
        });
        idleIcon = resourceMap.getIcon("StatusBar.idleIcon");
        statusAnimationLabel.setIcon(idleIcon);


        // connecting action tasks to status bar via TaskMonitor
        TaskMonitor taskMonitor = new TaskMonitor(getApplication().getContext());
        taskMonitor.addPropertyChangeListener(new java.beans.PropertyChangeListener() {
            public void propertyChange(java.beans.PropertyChangeEvent evt) {
                String propertyName = evt.getPropertyName();
                if ("started".equals(propertyName)) {
                    if (!busyIconTimer.isRunning()) {
                        statusAnimationLabel.setIcon(busyIcons[0]);
                        busyIconIndex = 0;
                        busyIconTimer.start();
                    }


                } else if ("done".equals(propertyName)) {
                    busyIconTimer.stop();
                    statusAnimationLabel.setIcon(idleIcon);
                } else if ("message".equals(propertyName)) {
                    String text = (String)(evt.getNewValue());
                    statusMessageLabel.setText((text == null) ? "" : text);
                    messageTimer.restart();
                } else if ("progress".equals(propertyName)) {
                    int value = (Integer)(evt.getNewValue());
                }
            }
        });
    }

    @Action
    public void showAboutBox() {
        if (aboutBox == null) {
            JFrame mainFrame = SimpleP2App.getApplication().getMainFrame();
            aboutBox = new SimpleP2AboutBox(mainFrame);
            aboutBox.setLocationRelativeTo(mainFrame);
        }
        SimpleP2App.getApplication().show(aboutBox);
    }

    private void AddBigProject() {
                FileOutputStream fileOut = null;
        
        Dir1 = jTextField1.getText() + "\\" + (String) jComboBox1.getSelectedItem() +  "\\" + jTextField2.getText();
        try {
            success = (new File(Dir1)).mkdirs();

        } catch (Exception exception) {
        }

        Dir2 = Dir1 + "\\Project File";
        try {
            success = (new File(Dir2)).mkdirs();

        } catch (Exception exception) {
        }

        Dir3 = Dir1 + "\\Stage File";
        try {
            success = (new File(Dir3)).mkdirs();

        } catch (Exception exception) {
        }

        Dir4 = Dir1 + "\\Quality File";
        try {
            success = (new File(Dir4)).mkdirs();

        } catch (Exception exception) {
        }
        
       if (jCheckBox5.isSelected()) {   
        FileName = Dir2 + "\\" + jTextField2.getText() + "-Project Plan" + ".xls";
        try {

            
            // create documents
            HSSFWorkbook wb = new HSSFWorkbook();

            // font for hyperlinks
            HSSFCellStyle hlink_style = wb.createCellStyle();
            HSSFFont hlink_font = wb.createFont();
            hlink_font.setUnderline(HSSFFont.U_SINGLE);
            hlink_font.setColor(HSSFColor.BLUE.index);
            hlink_style.setFont(hlink_font);
                    
            // cell styles: blue background
            HSSFCellStyle blue = wb.createCellStyle();
            blue.setFillBackgroundColor(HSSFColor.AQUA.index);
            blue.setFillForegroundColor(HSSFColor.AQUA.index);
            blue.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            // Date format
            HSSFCellStyle dateStyle = wb.createCellStyle();
            dateStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat("m/d/yy h:mm"));
            
                
            HSSFSheet sheet1 = wb.createSheet("Project Plan");
            row = sheet1.createRow((short)0);
            
            fileOut = new FileOutputStream(FileName);
            wb.write(fileOut);
            fileOut.close();
            
        } catch (IOException ex) {
            Logger.getLogger(SimpleP2View.class.getName()).log(Level.SEVERE, null, ex);
        } finally {
            try {
                fileOut.close();
            } catch (IOException ex) {
                Logger.getLogger(SimpleP2View.class.getName()).log(Level.SEVERE, null, ex);
            }
        }
    } 
       if (jCheckBox2.isSelected()) {
        FileName = Dir2 + "\\" + jTextField2.getText() + "-Risk Log" + ".xls";
        try {

            
            // create documents
            HSSFWorkbook wb = new HSSFWorkbook();

            // font for hyperlinks
            HSSFCellStyle hlink_style = wb.createCellStyle();
            HSSFFont hlink_font = wb.createFont();
            hlink_font.setUnderline(HSSFFont.U_SINGLE);
            hlink_font.setColor(HSSFColor.BLUE.index);
            hlink_style.setFont(hlink_font);
            
            // cell styles: blue background
            HSSFCellStyle blue = wb.createCellStyle();
            blue.setFillBackgroundColor(HSSFColor.AQUA.index);
            blue.setFillForegroundColor(HSSFColor.AQUA.index);
            blue.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            // Date format
            HSSFCellStyle dateStyle = wb.createCellStyle();
            dateStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat("m/d/yy h:mm"));
            
                
            // sheet creation
            HSSFSheet sheet1 = wb.createSheet("Risk Log");
            row = sheet1.createRow((short)0);
            cell = row.createCell((short)0);
            row = sheet1.createRow((short)1);
            cell = row.createCell((short)0);
            cell.setCellValue("Risk #");
            cell.setCellStyle(blue);
            cell = row.createCell((short)1);
            cell.setCellValue("Created");
            cell.setCellStyle(blue);
            cell = row.createCell((short)2);
            cell.setCellValue("Description");
            cell.setCellStyle(blue);
            sheet1.setColumnWidth((short)2, (short)10000);
            cell = row.createCell((short)3);
            cell.setCellValue("Probability");
            cell.setCellStyle(blue);
            cell = row.createCell((short)4);
            cell.setCellValue("Impact");
            cell.setCellStyle(blue);
            cell = row.createCell((short)5);
            cell.setCellValue("Probablilty x Impact");
            cell.setCellStyle(blue);
            cell = row.createCell((short)6);
            cell.setCellValue("Risk Owner");
            cell.setCellStyle(blue);
            cell = row.createCell((short)7);
            cell.setCellValue("Action");
            cell.setCellStyle(blue);
            
            
            sheet1.setColumnWidth((short)7, (short)10000);


            for (int i = 2; i < 21; i++) {
              row = sheet1.createRow((short)i);
              issuenumber = i - 1;
              cell = row.createCell((short)0);
              cell.setCellValue("" + issuenumber);
              cell.setCellStyle(blue);
            }

            fileOut = new FileOutputStream(FileName);
            wb.write(fileOut);
            fileOut.close();
            
        } catch (IOException ex) {
            Logger.getLogger(SimpleP2View.class.getName()).log(Level.SEVERE, null, ex);
        } finally {
            try {
                fileOut.close();
            } catch (IOException ex) {
                Logger.getLogger(SimpleP2View.class.getName()).log(Level.SEVERE, null, ex);
            }
        }
    }
        if (jCheckBox7.isSelected()) {
        FileName = Dir3 + "\\" + jTextField2.getText() + "-Stage Plans" + ".xls";
        try {

            
            // create documents
            HSSFWorkbook wb = new HSSFWorkbook();

            // font for hyperlinks
            HSSFCellStyle hlink_style = wb.createCellStyle();
            HSSFFont hlink_font = wb.createFont();
            hlink_font.setUnderline(HSSFFont.U_SINGLE);
            hlink_font.setColor(HSSFColor.BLUE.index);
            hlink_style.setFont(hlink_font);
            
            // cell styles: blue background
            HSSFCellStyle blue = wb.createCellStyle();
            blue.setFillBackgroundColor(HSSFColor.AQUA.index);
            blue.setFillForegroundColor(HSSFColor.AQUA.index);
            blue.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            // Date format
            HSSFCellStyle dateStyle = wb.createCellStyle();
            dateStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat("m/d/yy h:mm"));
            
                
            HSSFSheet sheet1 = wb.createSheet("Project Plan");
            row = sheet1.createRow((short)0);
            
            fileOut = new FileOutputStream(FileName);
            wb.write(fileOut);
            fileOut.close();
            
        } catch (IOException ex) {
            Logger.getLogger(SimpleP2View.class.getName()).log(Level.SEVERE, null, ex);
        } finally {
            try {
                fileOut.close();
            } catch (IOException ex) {
                Logger.getLogger(SimpleP2View.class.getName()).log(Level.SEVERE, null, ex);
            }
        }
    }
        
        if (jCheckBox3.isSelected()) {
        FileName = Dir4 + "\\" + jTextField2.getText() + "-Issue Log" + ".xls";
        try {

            
            // create documents
            HSSFWorkbook wb = new HSSFWorkbook();

            // font for hyperlinks
            HSSFCellStyle hlink_style = wb.createCellStyle();
            HSSFFont hlink_font = wb.createFont();
            hlink_font.setUnderline(HSSFFont.U_SINGLE);
            hlink_font.setColor(HSSFColor.BLUE.index);
            hlink_style.setFont(hlink_font);
            
            // cell styles: blue background
            HSSFCellStyle blue = wb.createCellStyle();
            blue.setFillBackgroundColor(HSSFColor.AQUA.index);
            blue.setFillForegroundColor(HSSFColor.AQUA.index);
            blue.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            // Date format
            HSSFCellStyle dateStyle = wb.createCellStyle();
            dateStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat("m/d/yy h:mm"));
            
                
            HSSFSheet sheet1 = wb.createSheet("Issue Log");
            row = sheet1.createRow((short)0);
            cell = row.createCell((short)0);
            
            row = sheet1.createRow((short)1);
            cell = row.createCell((short)0);
            cell.setCellValue("Issue #");
            cell.setCellStyle(blue);
            cell = row.createCell((short)1);
            cell.setCellValue("Call #");
            cell.setCellStyle(blue);
            cell = row.createCell((short)2);
            cell.setCellValue("Reported");
            cell.setCellStyle(blue);
            cell = row.createCell((short)3);
            cell.setCellValue("Description");
            cell.setCellStyle(blue);
            sheet1.setColumnWidth((short)3, (short)10000);
            cell = row.createCell((short)4);
            cell.setCellValue("Owner");
            cell.setCellStyle(blue);
            cell = row.createCell((short)5);
            cell.setCellValue("Action By");
            cell.setCellStyle(blue);
            cell = row.createCell((short)6);
            cell.setCellValue("Status");
            cell.setCellStyle(blue);
            cell = row.createCell((short)7);
            cell.setCellValue("Detail");
            cell.setCellStyle(blue);
            sheet1.setColumnWidth((short)7, (short)10000);

            for (int i = 2; i < 101; i++) {
              row = sheet1.createRow((short)i);
              issuenumber = i - 1;
              cell = row.createCell((short)0);
              cell.setCellValue("" + issuenumber);
              cell.setCellStyle(blue);
            }
            
            fileOut = new FileOutputStream(FileName);
            wb.write(fileOut);
            fileOut.close();
            
        } catch (IOException ex) {
            Logger.getLogger(SimpleP2View.class.getName()).log(Level.SEVERE, null, ex);
        } finally {
            try {
                fileOut.close();
            } catch (IOException ex) {
                Logger.getLogger(SimpleP2View.class.getName()).log(Level.SEVERE, null, ex);
            }
        }
    }
    
       if (jCheckBox6.isSelected()) {
        FileName = Dir3 + "\\" + jTextField2.getText() + "-Daily Log" + ".xls";
        try {

            
            // create documents
            HSSFWorkbook wb = new HSSFWorkbook();

            // font for hyperlinks
            HSSFCellStyle hlink_style = wb.createCellStyle();
            HSSFFont hlink_font = wb.createFont();
            hlink_font.setUnderline(HSSFFont.U_SINGLE);
            hlink_font.setColor(HSSFColor.BLUE.index);
            hlink_style.setFont(hlink_font);
            
            // cell styles: blue background
            HSSFCellStyle blue = wb.createCellStyle();
            blue.setFillBackgroundColor(HSSFColor.AQUA.index);
            blue.setFillForegroundColor(HSSFColor.AQUA.index);
            blue.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            // Date format
            HSSFCellStyle dateStyle = wb.createCellStyle();
            dateStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat("m/d/yy h:mm"));
            
                
            HSSFSheet sheet1 = wb.createSheet("Daily Log");
            row = sheet1.createRow((short)0);
            row = sheet1.createRow((short)1);
            cell = row.createCell((short)0);
            cell.setCellValue("Date");
            cell.setCellStyle(blue);
            cell = row.createCell((short)1);
            cell.setCellValue("Event");
            cell.setCellStyle(blue);
            cell = row.createCell((short)2);
            cell.setCellValue("Action / Comments");
            cell.setCellStyle(blue);

            for (int i = 2; i < 101; i++) {
              row = sheet1.createRow((short)i);
              issuenumber = i - 1;
              cell = row.createCell((short)0);
              cell.setCellValue("" + issuenumber);
              cell.setCellStyle(blue);
            }
            
            sheet1.autoSizeColumn((short)0);
            sheet1.autoSizeColumn((short)1);
            sheet1.autoSizeColumn((short)2);
            sheet1.autoSizeColumn((short)3);
            
            fileOut = new FileOutputStream(FileName);
            wb.write(fileOut);
            fileOut.close();
            
        } catch (IOException ex) {
            Logger.getLogger(SimpleP2View.class.getName()).log(Level.SEVERE, null, ex);
        } finally {
            try {
                fileOut.close();
            } catch (IOException ex) {
                Logger.getLogger(SimpleP2View.class.getName()).log(Level.SEVERE, null, ex);
            }
        }
    } 
       if (jCheckBox4.isSelected()) {        
        FileName = Dir2 + "\\" + jTextField2.getText() + "-Lessons Learned Log" + ".xls";
        try {

            
            // create documents
            HSSFWorkbook wb = new HSSFWorkbook();

            // font for hyperlinks
            HSSFCellStyle hlink_style = wb.createCellStyle();
            HSSFFont hlink_font = wb.createFont();
            hlink_font.setUnderline(HSSFFont.U_SINGLE);
            hlink_font.setColor(HSSFColor.BLUE.index);
            hlink_style.setFont(hlink_font);
            
            // cell styles: blue background
            HSSFCellStyle blue = wb.createCellStyle();
            blue.setFillBackgroundColor(HSSFColor.AQUA.index);
            blue.setFillForegroundColor(HSSFColor.AQUA.index);
            blue.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            // Date format
            HSSFCellStyle dateStyle = wb.createCellStyle();
            dateStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat("m/d/yy h:mm"));
            
                
            HSSFSheet sheet1 = wb.createSheet("Lessons Learned");
            row = sheet1.createRow((short)0);
            row = sheet1.createRow((short)1);
            cell = row.createCell((short)0);
            cell.setCellValue("Number");
            cell.setCellStyle(blue);
            cell = row.createCell((short)1);
            cell.setCellValue("Description");
            cell.setCellStyle(blue);

            for (int i = 2; i < 101; i++) {
              row = sheet1.createRow((short)i);
              issuenumber = i - 1;
              cell = row.createCell((short)0);
              cell.setCellValue("" + issuenumber);
              cell.setCellStyle(blue);
            }
            
            sheet1.autoSizeColumn((short)0);
            sheet1.autoSizeColumn((short)1);
            
            fileOut = new FileOutputStream(FileName);
            wb.write(fileOut);
            fileOut.close();
            
        } catch (IOException ex) {
            Logger.getLogger(SimpleP2View.class.getName()).log(Level.SEVERE, null, ex);
        } finally {
            try {
                fileOut.close();
            } catch (IOException ex) {
                Logger.getLogger(SimpleP2View.class.getName()).log(Level.SEVERE, null, ex);
            }
        }
       }   //if jCheckBox4 is selected
        
        if (jCheckBox8.isSelected()) {
           add_index_page();
        }
        
        if (jCheckBox1.isSelected()) {
           add_pm_team_doc();
        }
        if (jCheckBox9.isSelected()) {
           add_quality_plan_doc();
        }
        if (jCheckBox10.isSelected()) {
           add_communication_plan_doc();
        }
        jTextField4.setText("Project Files have been prepared");
    }

    private void add_communication_plan_doc() {
        
       File f = new File (Dir2 + "\\" + jTextField2.getText() + "-Communication Plan.doc");
       File f_def = new File ("complan.xml");

       try {
		FileInputStream fin = new FileInputStream (f_def);
		int filesize = (int)f_def.length();
		data = new byte[filesize];
		fin.read (data, 0, filesize);
	} catch (FileNotFoundException exc) {
		String errorString = "File not found: complan.xml";
		data = new byte[errorString.length()];
		errorString.getBytes();
     	        jTextField4.setText(errorString);
	} catch (IOException exc) {
		String errorString = "IOException: complan.xml";
		data = new byte[errorString.length()];
		errorString.getBytes();
     	        jTextField4.setText(errorString);
	}

               //myTextArea.setText (new String (data, 0));
               tekst = (new String (data));
					// tekst = myTextArea.getText();
					tekst2 = tekst;

               try {
                       FileOutputStream fon = new FileOutputStream (f);
                       OutputStream bfon = new BufferedOutputStream(fon);
                       OutputStreamWriter bfono = new OutputStreamWriter(bfon, "8859_1");
					  	     bfono.write (tekst2);
						     bfono.close ();
					} catch (IOException exc) {
						     String errorString = "IOException: " + filename;
					}
        

    }

    private void add_index_page() {
                  File f = new File (Dir1 + "\\" + "index.html");
                  					File f_def = new File ("default.html");

                   try {
						FileInputStream fin = new FileInputStream (f_def);
						int filesize = (int)f_def.length();
						data = new byte[filesize];
						fin.read (data, 0, filesize);
					} catch (FileNotFoundException exc) {
						String errorString = "File not found: default.html";
						data = new byte[errorString.length()];
						errorString.getBytes();
     				                jTextField4.setText(errorString);
					} catch (IOException exc) {
						String errorString = "IOException: default.html";
						data = new byte[errorString.length()];
						errorString.getBytes();
     				                jTextField4.setText(errorString);
					}

                                        //myTextArea.setText (new String (data, 0));
                                        tekst = (new String (data));
					// tekst = myTextArea.getText();
					tekst2 = tekst.replaceAll("---PROJECT---",(String) jComboBox1.getSelectedItem() + " - " + jTextField2.getText());

                                        tekst2 = tekst2.replaceAll("---RISK LOG LINK---","./Project File/" + jTextField2.getText() + "-Risk Log.xls");
					tekst2 = tekst2.replaceAll("---RISK LOG NAME---",jTextField2.getText() + "-Risk Log.xls");
                                        
					tekst2 = tekst2.replaceAll("---PROJECT PLAN LINK---","./Project File/" + jTextField2.getText() + "-Project Plan.xls");
					tekst2 = tekst2.replaceAll("---PROJECT PLAN NAME---",jTextField2.getText() + "-Project Plan.xls");

					tekst2 = tekst2.replaceAll("---LESSONS LEARNED LOG LINK---","./Project File/" + jTextField2.getText() + "-Lessons Learned Log.xls");
					tekst2 = tekst2.replaceAll("---LESSONS LEARNED LOG NAME---",jTextField2.getText() + "-Lessons Learned Log.xls");

                  			tekst2 = tekst2.replaceAll("---ISSUE LOG LINK---","./Quality File/" + jTextField2.getText() + "-Issue Log.xls");
					tekst2 = tekst2.replaceAll("---ISSUE LOG NAME---",jTextField2.getText() + "-Issue Log.xls");

                  			tekst2 = tekst2.replaceAll("---DAILY LOG LINK---","./Stage File/" + jTextField2.getText() + "-Daily Log.xls");
					tekst2 = tekst2.replaceAll("---DAILY LOG NAME---",jTextField2.getText() + "-Daily Log.xls");

                  			tekst2 = tekst2.replaceAll("---STAGE PLANS LINK---","./Stage File/" + jTextField2.getText() + "-Stage Plans.xls");
					tekst2 = tekst2.replaceAll("---STAGE PLANS NAME---",jTextField2.getText() + "-Stage Plans.xls");

                  			tekst2 = tekst2.replaceAll("---PROJECT COMM PLAN LINK---","./Project File/" + jTextField2.getText() + "-Communication Plan.doc");
					tekst2 = tekst2.replaceAll("---PROJECT COMM PLAN NAME---",jTextField2.getText() + "-Communication Plan.doc");

                                        tekst2 = tekst2.replaceAll("---PROJECT QUALITY PLAN LINK---","./Project File/" + jTextField2.getText() + "-Project Quality Plan.doc");
					tekst2 = tekst2.replaceAll("---PROJECT QUALITY PLAN NAME---",jTextField2.getText() + "-Project Quality Plan.doc");

                                        tekst2 = tekst2.replaceAll("---PROJECT TEAM LINK---","./Project File/" + jTextField2.getText() + "-Project Management Team.doc");
					tekst2 = tekst2.replaceAll("---PROJECT TEAM NAME---",jTextField2.getText() + "-Project Management Team.doc");

                                        try {
                                                FileOutputStream fon = new FileOutputStream (f);
                                                OutputStream bfon = new BufferedOutputStream(fon);
                                                OutputStreamWriter bfono = new OutputStreamWriter(bfon, "8859_1");
						bfono.write (tekst2);
						bfono.close ();
					} catch (IOException exc) {
						String errorString = "IOException: " + filename;
					}

    }

    private void add_pm_team_doc() {

       File f = new File (Dir2 + "\\" + jTextField2.getText() + "-Project Management Team.doc");
       File f_def = new File ("pmteam.xml");

       try {
		FileInputStream fin = new FileInputStream (f_def);
		int filesize = (int)f_def.length();
		data = new byte[filesize];
		fin.read (data, 0, filesize);
	} catch (FileNotFoundException exc) {
		String errorString = "File not found: default.html";
		data = new byte[errorString.length()];
		errorString.getBytes();
     	        jTextField4.setText(errorString);
	} catch (IOException exc) {
		String errorString = "IOException: default.html";
		data = new byte[errorString.length()];
		errorString.getBytes();
     	        jTextField4.setText(errorString);
	}

               //myTextArea.setText (new String (data, 0));
               tekst = (new String (data));
					// tekst = myTextArea.getText();
					tekst2 = tekst.replaceAll("---PM-NAME---",jTextField5.getText());
					tekst2 = tekst2.replaceAll("---PM-PHONE---","" + jTextField6.getText());
					tekst2 = tekst2.replaceAll("---PM-MAIL---","" + jTextField7.getText());

               try {
                       FileOutputStream fon = new FileOutputStream (f);
                       OutputStream bfon = new BufferedOutputStream(fon);
                       OutputStreamWriter bfono = new OutputStreamWriter(bfon, "8859_1");
					    	  bfono.write (tekst2);
						     bfono.close ();
					} catch (IOException exc) {
						     String errorString = "IOException: " + filename;
					}
        
    }

    private void add_quality_plan_doc() {
        
       File f = new File (Dir2 + "\\" + jTextField2.getText() + "-Project Quality Plan.doc");
       File f_def = new File ("qualplan.xml");

       try {
		FileInputStream fin = new FileInputStream (f_def);
		int filesize = (int)f_def.length();
		data = new byte[filesize];
		fin.read (data, 0, filesize);
	} catch (FileNotFoundException exc) {
		String errorString = "File not found: qualplan.xml";
		data = new byte[errorString.length()];
		errorString.getBytes();
     	        jTextField4.setText(errorString);
	} catch (IOException exc) {
		String errorString = "IOException: qualplan.xml";
		data = new byte[errorString.length()];
		errorString.getBytes();
     	        jTextField4.setText(errorString);
	}

               //myTextArea.setText (new String (data, 0));
               tekst = (new String (data));
					// tekst = myTextArea.getText();
					tekst2 = tekst;

            try {
               FileOutputStream fon = new FileOutputStream (f);
               OutputStream bfon = new BufferedOutputStream(fon);
               OutputStreamWriter bfono = new OutputStreamWriter(bfon, "8859_1");
					bfono.write (tekst2);
					bfono.close ();
					} catch (IOException exc) {
						String errorString = "IOException: " + filename;
					}
        

    }

    
    /** This method is called from within the constructor to
     * initialize the form.
     * WARNING: Do NOT modify this code. The content of this method is
     * always regenerated by the Form Editor.
     */
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        mainPanel = new javax.swing.JPanel();
        jLabel2 = new javax.swing.JLabel();
        jTextField1 = new javax.swing.JTextField();
        jButton1 = new javax.swing.JButton();
        jLabel3 = new javax.swing.JLabel();
        jComboBox1 = new javax.swing.JComboBox();
        jLabel4 = new javax.swing.JLabel();
        jTextField2 = new javax.swing.JTextField();
        jTextField3 = new javax.swing.JTextField();
        jButton2 = new javax.swing.JButton();
        jButton3 = new javax.swing.JButton();
        jRadioButton1 = new javax.swing.JRadioButton();
        jRadioButton2 = new javax.swing.JRadioButton();
        jPanel1 = new javax.swing.JPanel();
        jLabel6 = new javax.swing.JLabel();
        jTextField5 = new javax.swing.JTextField();
        jLabel7 = new javax.swing.JLabel();
        jTextField6 = new javax.swing.JTextField();
        jLabel8 = new javax.swing.JLabel();
        jTextField7 = new javax.swing.JTextField();
        jLabel9 = new javax.swing.JLabel();
        jCheckBox1 = new javax.swing.JCheckBox();
        jCheckBox2 = new javax.swing.JCheckBox();
        jCheckBox3 = new javax.swing.JCheckBox();
        jCheckBox4 = new javax.swing.JCheckBox();
        jCheckBox5 = new javax.swing.JCheckBox();
        jCheckBox6 = new javax.swing.JCheckBox();
        jCheckBox7 = new javax.swing.JCheckBox();
        jCheckBox8 = new javax.swing.JCheckBox();
        jCheckBox9 = new javax.swing.JCheckBox();
        jCheckBox10 = new javax.swing.JCheckBox();
        menuBar = new javax.swing.JMenuBar();
        javax.swing.JMenu fileMenu = new javax.swing.JMenu();
        javax.swing.JMenuItem exitMenuItem = new javax.swing.JMenuItem();
        javax.swing.JMenu helpMenu = new javax.swing.JMenu();
        javax.swing.JMenuItem aboutMenuItem = new javax.swing.JMenuItem();
        jMenuItem5 = new javax.swing.JMenuItem();
        statusPanel = new javax.swing.JPanel();
        javax.swing.JSeparator statusPanelSeparator = new javax.swing.JSeparator();
        statusMessageLabel = new javax.swing.JLabel();
        statusAnimationLabel = new javax.swing.JLabel();
        jLabel1 = new javax.swing.JLabel();
        jTextField4 = new javax.swing.JTextField();
        jLabel5 = new javax.swing.JLabel();
        buttonGroup1 = new javax.swing.ButtonGroup();

        mainPanel.setName("mainPanel"); // NOI18N

        org.jdesktop.application.ResourceMap resourceMap = org.jdesktop.application.Application.getInstance(simplep2.SimpleP2App.class).getContext().getResourceMap(SimpleP2View.class);
        jLabel2.setText(resourceMap.getString("jLabel2.text")); // NOI18N
        jLabel2.setName("jLabel2"); // NOI18N

        jTextField1.setText(resourceMap.getString("jTextField1.text")); // NOI18N
        jTextField1.setName("jTextField1"); // NOI18N

        jButton1.setText(resourceMap.getString("jButton1.text")); // NOI18N
        jButton1.setName("jButton1"); // NOI18N
        jButton1.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                select_dir(evt);
            }
        });

        jLabel3.setText(resourceMap.getString("jLabel3.text")); // NOI18N
        jLabel3.setName("jLabel3"); // NOI18N

        jComboBox1.setName("jComboBox1"); // NOI18N
        jComboBox1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                select_cust(evt);
            }
        });

        jLabel4.setText(resourceMap.getString("jLabel4.text")); // NOI18N
        jLabel4.setName("jLabel4"); // NOI18N

        jTextField2.setText(resourceMap.getString("jTextField2.text")); // NOI18N
        jTextField2.setName("jTextField2"); // NOI18N
        jTextField2.addFocusListener(new java.awt.event.FocusAdapter() {
            public void focusGained(java.awt.event.FocusEvent evt) {
                reset_background_color(evt);
            }
        });

        jTextField3.setEditable(false);
        jTextField3.setText(resourceMap.getString("jTextField3.text")); // NOI18N
        jTextField3.setFocusable(false);
        jTextField3.setName("jTextField3"); // NOI18N

        jButton2.setText(resourceMap.getString("jButton2.text")); // NOI18N
        buttonGroup1.add(jButton2);
        jButton2.setName("jButton2"); // NOI18N
        jButton2.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                save_settings(evt);
            }
        });

        jButton3.setText(resourceMap.getString("jButton3.text")); // NOI18N
        buttonGroup1.add(jButton3);
        jButton3.setName("jButton3"); // NOI18N
        jButton3.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                create_project_files(evt);
            }
        });

        buttonGroup1.add(jRadioButton1);
        jRadioButton1.setText(resourceMap.getString("jRadioButton1.text")); // NOI18N
        jRadioButton1.setName("jRadioButton1"); // NOI18N

        buttonGroup1.add(jRadioButton2);
        jRadioButton2.setSelected(true);
        jRadioButton2.setText(resourceMap.getString("jRadioButton2.text")); // NOI18N
        jRadioButton2.setName("jRadioButton2"); // NOI18N

        jPanel1.setBorder(javax.swing.BorderFactory.createTitledBorder(null, resourceMap.getString("jPanel1.border.title"), javax.swing.border.TitledBorder.LEFT, javax.swing.border.TitledBorder.DEFAULT_POSITION, new java.awt.Font("Tahoma", 0, 11), resourceMap.getColor("jPanel1.border.titleColor"))); // NOI18N
        jPanel1.setName("jPanel1"); // NOI18N

        jLabel6.setText(resourceMap.getString("jLabel6.text")); // NOI18N
        jLabel6.setName("jLabel6"); // NOI18N

        jTextField5.setText(resourceMap.getString("jTextField5.text")); // NOI18N
        jTextField5.setName("jTextField5"); // NOI18N

        jLabel7.setText(resourceMap.getString("jLabel7.text")); // NOI18N
        jLabel7.setName("jLabel7"); // NOI18N

        jTextField6.setText(resourceMap.getString("jTextField6.text")); // NOI18N
        jTextField6.setName("jTextField6"); // NOI18N

        jLabel8.setText(resourceMap.getString("jLabel8.text")); // NOI18N
        jLabel8.setName("jLabel8"); // NOI18N

        jTextField7.setText(resourceMap.getString("jTextField7.text")); // NOI18N
        jTextField7.setName("jTextField7"); // NOI18N

        jLabel9.setText(resourceMap.getString("jLabel9.text")); // NOI18N
        jLabel9.setName("jLabel9"); // NOI18N

        jCheckBox1.setSelected(true);
        jCheckBox1.setText(resourceMap.getString("jCheckBox1.text")); // NOI18N
        jCheckBox1.setName("jCheckBox1"); // NOI18N

        jCheckBox2.setSelected(true);
        jCheckBox2.setText(resourceMap.getString("jCheckBox2.text")); // NOI18N
        jCheckBox2.setName("jCheckBox2"); // NOI18N

        jCheckBox3.setSelected(true);
        jCheckBox3.setText(resourceMap.getString("jCheckBox3.text")); // NOI18N
        jCheckBox3.setName("jCheckBox3"); // NOI18N

        jCheckBox4.setSelected(true);
        jCheckBox4.setText(resourceMap.getString("jCheckBox4.text")); // NOI18N
        jCheckBox4.setName("jCheckBox4"); // NOI18N

        jCheckBox5.setSelected(true);
        jCheckBox5.setText(resourceMap.getString("jCheckBox5.text")); // NOI18N
        jCheckBox5.setName("jCheckBox5"); // NOI18N

        jCheckBox6.setSelected(true);
        jCheckBox6.setText(resourceMap.getString("jCheckBox6.text")); // NOI18N
        jCheckBox6.setName("jCheckBox6"); // NOI18N

        jCheckBox7.setText(resourceMap.getString("jCheckBox7.text")); // NOI18N
        jCheckBox7.setName("jCheckBox7"); // NOI18N

        jCheckBox8.setSelected(true);
        jCheckBox8.setText(resourceMap.getString("jCheckBox8.text")); // NOI18N
        jCheckBox8.setName("jCheckBox8"); // NOI18N

        jCheckBox9.setSelected(true);
        jCheckBox9.setText(resourceMap.getString("jCheckBox9.text")); // NOI18N
        jCheckBox9.setName("jCheckBox9"); // NOI18N

        jCheckBox10.setSelected(true);
        jCheckBox10.setText(resourceMap.getString("jCheckBox10.text")); // NOI18N
        jCheckBox10.setName("jCheckBox10"); // NOI18N

        javax.swing.GroupLayout jPanel1Layout = new javax.swing.GroupLayout(jPanel1);
        jPanel1.setLayout(jPanel1Layout);
        jPanel1Layout.setHorizontalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                    .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                        .addComponent(jLabel9)
                        .addGroup(jPanel1Layout.createSequentialGroup()
                            .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                .addComponent(jCheckBox2)
                                .addComponent(jCheckBox3)
                                .addComponent(jCheckBox4)
                                .addComponent(jCheckBox6))
                            .addGap(2, 2, 2)
                            .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                .addComponent(jCheckBox10)
                                .addGroup(jPanel1Layout.createSequentialGroup()
                                    .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                        .addComponent(jCheckBox1)
                                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                                            .addGroup(jPanel1Layout.createSequentialGroup()
                                                .addComponent(jCheckBox5)
                                                .addGap(36, 36, 36))
                                            .addGroup(jPanel1Layout.createSequentialGroup()
                                                .addComponent(jCheckBox9)
                                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED))))
                                    .addGap(8, 8, 8)
                                    .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                        .addComponent(jCheckBox8)
                                        .addComponent(jCheckBox7))))))
                    .addGroup(jPanel1Layout.createSequentialGroup()
                        .addComponent(jLabel6)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                        .addComponent(jTextField5, javax.swing.GroupLayout.PREFERRED_SIZE, 156, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(jLabel7)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(jTextField6, javax.swing.GroupLayout.PREFERRED_SIZE, 125, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addGap(4, 4, 4)))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jLabel8)
                .addGap(5, 5, 5)
                .addComponent(jTextField7, javax.swing.GroupLayout.DEFAULT_SIZE, 218, Short.MAX_VALUE)
                .addContainerGap())
        );
        jPanel1Layout.setVerticalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel6)
                    .addComponent(jTextField5, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jLabel7)
                    .addComponent(jTextField7, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jLabel8)
                    .addComponent(jTextField6, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jLabel9)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jCheckBox6)
                    .addComponent(jCheckBox1)
                    .addComponent(jCheckBox7))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jCheckBox2)
                    .addComponent(jCheckBox5)
                    .addComponent(jCheckBox8))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jCheckBox3)
                    .addComponent(jCheckBox9))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jCheckBox4)
                    .addComponent(jCheckBox10))
                .addContainerGap())
        );

        javax.swing.GroupLayout mainPanelLayout = new javax.swing.GroupLayout(mainPanel);
        mainPanel.setLayout(mainPanelLayout);
        mainPanelLayout.setHorizontalGroup(
            mainPanelLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(mainPanelLayout.createSequentialGroup()
                .addContainerGap()
                .addGroup(mainPanelLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jLabel2)
                    .addComponent(jLabel3)
                    .addComponent(jLabel4))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(mainPanelLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jTextField1, javax.swing.GroupLayout.DEFAULT_SIZE, 436, Short.MAX_VALUE)
                    .addGroup(mainPanelLayout.createSequentialGroup()
                        .addGroup(mainPanelLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(jComboBox1, javax.swing.GroupLayout.PREFERRED_SIZE, 106, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jTextField2, javax.swing.GroupLayout.PREFERRED_SIZE, 87, javax.swing.GroupLayout.PREFERRED_SIZE))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addGroup(mainPanelLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addGroup(mainPanelLayout.createSequentialGroup()
                                .addComponent(jRadioButton1)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                .addComponent(jRadioButton2))
                            .addComponent(jTextField3, javax.swing.GroupLayout.DEFAULT_SIZE, 324, Short.MAX_VALUE))))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(mainPanelLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                    .addComponent(jButton3, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addComponent(jButton1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addComponent(jButton2, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                .addContainerGap())
            .addComponent(jPanel1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
        );
        mainPanelLayout.setVerticalGroup(
            mainPanelLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(mainPanelLayout.createSequentialGroup()
                .addContainerGap()
                .addGroup(mainPanelLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel2)
                    .addComponent(jTextField1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jButton1))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(mainPanelLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel3)
                    .addComponent(jComboBox1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jTextField3, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jButton2))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(mainPanelLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel4)
                    .addComponent(jTextField2, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jButton3)
                    .addComponent(jRadioButton1)
                    .addComponent(jRadioButton2))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                .addComponent(jPanel1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                .addContainerGap())
        );

        menuBar.setName("menuBar"); // NOI18N

        fileMenu.setText(resourceMap.getString("fileMenu.text")); // NOI18N
        fileMenu.setName("fileMenu"); // NOI18N

        javax.swing.ActionMap actionMap = org.jdesktop.application.Application.getInstance(simplep2.SimpleP2App.class).getContext().getActionMap(SimpleP2View.class, this);
        exitMenuItem.setAction(actionMap.get("quit")); // NOI18N
        exitMenuItem.setName("exitMenuItem"); // NOI18N
        fileMenu.add(exitMenuItem);

        menuBar.add(fileMenu);

        helpMenu.setText(resourceMap.getString("helpMenu.text")); // NOI18N
        helpMenu.setName("helpMenu"); // NOI18N

        aboutMenuItem.setAction(actionMap.get("showAboutBox")); // NOI18N
        aboutMenuItem.setName("aboutMenuItem"); // NOI18N
        helpMenu.add(aboutMenuItem);

        jMenuItem5.setText(resourceMap.getString("jMenuItem5.text")); // NOI18N
        jMenuItem5.setName("jMenuItem5"); // NOI18N
        jMenuItem5.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                show_help1(evt);
            }
        });
        helpMenu.add(jMenuItem5);

        menuBar.add(helpMenu);

        statusPanel.setName("statusPanel"); // NOI18N

        statusPanelSeparator.setName("statusPanelSeparator"); // NOI18N

        statusMessageLabel.setName("statusMessageLabel"); // NOI18N

        statusAnimationLabel.setHorizontalAlignment(javax.swing.SwingConstants.LEFT);
        statusAnimationLabel.setName("statusAnimationLabel"); // NOI18N

        jLabel1.setFont(resourceMap.getFont("jLabel1.font")); // NOI18N
        jLabel1.setText(resourceMap.getString("jLabel1.text")); // NOI18N
        jLabel1.setName("jLabel1"); // NOI18N

        jTextField4.setEditable(false);
        jTextField4.setText(resourceMap.getString("jTextField4.text")); // NOI18N
        jTextField4.setName("jTextField4"); // NOI18N

        jLabel5.setText(resourceMap.getString("jLabel5.text")); // NOI18N
        jLabel5.setName("jLabel5"); // NOI18N

        javax.swing.GroupLayout statusPanelLayout = new javax.swing.GroupLayout(statusPanel);
        statusPanel.setLayout(statusPanelLayout);
        statusPanelLayout.setHorizontalGroup(
            statusPanelLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(statusPanelSeparator, javax.swing.GroupLayout.DEFAULT_SIZE, 679, Short.MAX_VALUE)
            .addGroup(statusPanelLayout.createSequentialGroup()
                .addGroup(statusPanelLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(statusPanelLayout.createSequentialGroup()
                        .addContainerGap()
                        .addComponent(statusMessageLabel)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(jLabel5))
                    .addComponent(jLabel1))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                .addComponent(jTextField4, javax.swing.GroupLayout.DEFAULT_SIZE, 608, Short.MAX_VALUE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(statusAnimationLabel)
                .addContainerGap())
        );
        statusPanelLayout.setVerticalGroup(
            statusPanelLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(statusPanelLayout.createSequentialGroup()
                .addComponent(statusPanelSeparator, javax.swing.GroupLayout.PREFERRED_SIZE, 2, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(statusPanelLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, statusPanelLayout.createSequentialGroup()
                        .addGroup(statusPanelLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(statusMessageLabel)
                            .addComponent(statusAnimationLabel))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(jLabel1))
                    .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, statusPanelLayout.createSequentialGroup()
                        .addGroup(statusPanelLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(jLabel5)
                            .addComponent(jTextField4, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                        .addContainerGap())))
        );

        setComponent(mainPanel);
        setMenuBar(menuBar);
        setStatusBar(statusPanel);
    }// </editor-fold>//GEN-END:initComponents

    private void select_cust(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_select_cust

        String selName = (String)jComboBox1.getSelectedItem();
        selNum = jComboBox1.getSelectedIndex();
        if (selNum > -1) {
            try {
                // System.out.println(selNum);
                jTextField3.setText(custnames[selNum]);
            } finally {
            }


    }//GEN-LAST:event_select_cust

    }   
    
        private void save_settings(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_save_settings
                   File f_ini = new File ("settings.ini");
        Properties p = new Properties();
        p.setProperty("DEFDIR",jTextField1.getText());
        p.setProperty("CUSTFILE",customerfile);
        p.setProperty("DEF_PM",jTextField5.getText());
        p.setProperty("DEF_PM_PHONE",jTextField6.getText());
        p.setProperty("DEF_PM_MAIL",jTextField7.getText());
        
        FileOutputStream out = null;
        try {
            out = new FileOutputStream(f_ini);
        } catch (FileNotFoundException ex) {
            ex.printStackTrace();
        }
        try {
          p.store(out, "---SimpleP2 settings---");
        } catch (IOException ex) {
            ex.printStackTrace();
        }
        try {
            out.close();
        } catch (IOException ex) {
            ex.printStackTrace();
        }
        
        jTextField4.setText("Settings saved to settings.ini");


        }//GEN-LAST:event_save_settings

        private void select_dir(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_select_dir
                
          FilenameFilter filter = new FilenameFilter() { 
          public boolean accept(File dir, String name) { 
          return !name.startsWith("Thumbs.d"); 
          } 
          }; 


          DirectoryChooser GetDir = new DirectoryChooser() ;
          jTextField1.setText(GetDir.getSelection());
          
        }//GEN-LAST:event_select_dir

        private void create_project_files(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_create_project_files
            String TheProject = jTextField2.getText();
            if (TheProject != null && TheProject.length() == 0) {
              jTextField4.setBackground(Color.RED);
              jTextField4.setText("Enter a Project ID before generating project files");
              jTextField2.setBackground(Color.RED);
            } else {
               jTextField4.setText("Writing project Files...");
               jTextField2.setBackground(Color.WHITE);
               jTextField4.setBackground(Color.WHITE);
               if (jRadioButton1.isSelected()) {
                 AddProject();
               } else {
                 AddBigProject();  
               }
            }
        }//GEN-LAST:event_create_project_files

        private void show_help1(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_show_help1
            
            JFrame helpframe = new JFrame("PRINCE2 File generator Help");
            helpframe.setSize(700, 800);
            jTextArea1 = new javax.swing.JTextArea();
            helpframe.add(jTextArea1);
            jTextArea1.setFocusable(Boolean.FALSE);
            // read help file
       File f_help = new File ("help.txt");

       try {
		FileInputStream fin = new FileInputStream (f_help);
		int filesize = (int)f_help.length();
		data = new byte[filesize];
		fin.read (data, 0, filesize);
	} catch (FileNotFoundException exc) {
		String errorString = "File not found: help.txt";
		data = new byte[errorString.length()];
		errorString.getBytes();
     	        jTextField4.setText(errorString);
	} catch (IOException exc) {
		String errorString = "IOException: help.txt";
		data = new byte[errorString.length()];
		errorString.getBytes();
     	        jTextField4.setText(errorString);
	}

            //myTextArea.setText (new String (data, 0));
            tekst = (new String (data));
 	    // tekst = myTextArea.getText();
	    helptekst = tekst;
            
            // show help file
            jTextArea1.setBackground(Color.lightGray);
            jTextArea1.setText(helptekst);
            
            helpframe.setVisible(true);

        }//GEN-LAST:event_show_help1

        private void reset_background_color(java.awt.event.FocusEvent evt) {//GEN-FIRST:event_reset_background_color
               jTextField2.setBackground(Color.WHITE);
               jTextField4.setBackground(Color.WHITE);
               jTextField4.setText("Ready");

        }//GEN-LAST:event_reset_background_color
    
    
    @Action
    public void AddProject() {
        FileOutputStream fileOut = null;
        
        Dir1 = jTextField1.getText() + "\\" + (String) jComboBox1.getSelectedItem() +  "\\" + jTextField2.getText();
        try {
            success = (new File(Dir1)).mkdirs();

        } catch (Exception exception) {
        }

        FileName = jTextField1.getText() + "\\" + (String) jComboBox1.getSelectedItem() + "\\" + jTextField2.getText() + "\\" + jTextField2.getText() + ".xls";
        try {

    //cell style for hyperlinks
    //by default hypelrinks are blue and underlined
            
            
            // create directory
            
            // create documents
            HSSFWorkbook wb = new HSSFWorkbook();

            // font for hyperlinks
            HSSFCellStyle hlink_style = wb.createCellStyle();
            HSSFFont hlink_font = wb.createFont();
            hlink_font.setUnderline(HSSFFont.U_SINGLE);
            hlink_font.setColor(HSSFColor.BLUE.index);
            hlink_style.setFont(hlink_font);
            
            // cell styles: blue background
            HSSFCellStyle blue = wb.createCellStyle();
            blue.setFillBackgroundColor(HSSFColor.AQUA.index);
            blue.setFillForegroundColor(HSSFColor.AQUA.index);
            blue.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            // Date format
            HSSFCellStyle dateStyle = wb.createCellStyle();
            dateStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat("m/d/yy h:mm"));
            
                
            HSSFSheet sheet1 = wb.createSheet("Index");
            row = sheet1.createRow((short)0);
            row.createCell((short)0).setCellValue("Index");
            row.createCell((short)1).setCellValue("");
            row.createCell((short)2).setCellValue("");
            row.createCell((short)3).setCellValue("");
            row.createCell((short)4).setCellValue("");
            row.createCell((short)5).setCellValue("");
            row.createCell((short)6).setCellValue("");
            row.createCell((short)7).setCellValue("");

            row.createCell((short)8).setCellValue("Created:");
            Calendar now = Calendar.getInstance();
            String creationDate = String.format("%1$td/%1$tm/%1$tY", now);
            row.createCell((short)9).setCellValue(creationDate);
            
            
            row = sheet1.createRow((short)1);
            row = sheet1.createRow((short)2);
            cell = row.createCell((short)1);
            cell.setCellValue("Project Mandate");
            link = new HSSFHyperlink(HSSFHyperlink.LINK_DOCUMENT);
            link.setAddress("'Project Mandate'!A1");
            cell.setHyperlink(link);
            cell.setCellStyle(hlink_style);

            row = sheet1.createRow((short)3);
            cell = row.createCell((short)1);
            cell.setCellValue("Business Case");
            link = new HSSFHyperlink(HSSFHyperlink.LINK_DOCUMENT);
            link.setAddress("'Business Case'!A1");
            cell.setHyperlink(link);
            cell.setCellStyle(hlink_style);
            
            row = sheet1.createRow((short)4);
            cell = row.createCell((short)1);
            cell.setCellValue("Project Brief");
            link = new HSSFHyperlink(HSSFHyperlink.LINK_DOCUMENT);
            link.setAddress("'Project Brief'!A1");
            cell.setHyperlink(link);
            cell.setCellStyle(hlink_style);

            row = sheet1.createRow((short)5);
            row = sheet1.createRow((short)6);
            cell = row.createCell((short)0);
            cell.setCellValue("Logs:");
            cell = row.createCell((short)1);
            cell.setCellValue("Issue Log");
            link = new HSSFHyperlink(HSSFHyperlink.LINK_DOCUMENT);
            link.setAddress("'Issue Log'!A1");
            cell.setHyperlink(link);
            cell.setCellStyle(hlink_style);
            
            row = sheet1.createRow((short)7);
            cell = row.createCell((short)1);
            cell.setCellValue("Daily Log");
            link = new HSSFHyperlink(HSSFHyperlink.LINK_DOCUMENT);
            link.setAddress("'Daily Log'!A1");
            cell.setHyperlink(link);
            cell.setCellStyle(hlink_style);

            row = sheet1.createRow((short)8);
            cell = row.createCell((short)1);
            cell.setCellValue("Lessons Learned Log");
            link = new HSSFHyperlink(HSSFHyperlink.LINK_DOCUMENT);
            link.setAddress("'Lessons Learned Log'!A1");
            cell.setHyperlink(link);
            cell.setCellStyle(hlink_style);
            
            row = sheet1.createRow((short)9);
            cell = row.createCell((short)1);
            cell.setCellValue("Quality Log");
            link = new HSSFHyperlink(HSSFHyperlink.LINK_DOCUMENT);
            link.setAddress("'Quality Log'!A1");
            cell.setHyperlink(link);
            cell.setCellStyle(hlink_style);

            row = sheet1.createRow((short)10);
            cell = row.createCell((short)1);
            cell.setCellValue("Risk Log");
            link = new HSSFHyperlink(HSSFHyperlink.LINK_DOCUMENT);
            link.setAddress("'Risk Log'!A1");
            cell.setHyperlink(link);
            cell.setCellStyle(hlink_style);
            
            row = sheet1.createRow((short)11);
            row = sheet1.createRow((short)12);
            cell = row.createCell((short)0);
            cell.setCellValue("Plans:");
            cell = row.createCell((short)1);
            cell.setCellValue("Project Plan");
            link = new HSSFHyperlink(HSSFHyperlink.LINK_DOCUMENT);
            link.setAddress("'Project Plan'!A1");
            cell.setHyperlink(link);
            cell.setCellStyle(hlink_style);

            row = sheet1.createRow((short)13);
            cell = row.createCell((short)0);
            cell = row.createCell((short)1);
            cell.setCellValue("Project Quality Plan");
            link = new HSSFHyperlink(HSSFHyperlink.LINK_DOCUMENT);
            link.setAddress("'PQP'!A1");
            cell.setHyperlink(link);
            cell.setCellStyle(hlink_style);

            row = sheet1.createRow((short)14);
            cell = row.createCell((short)0);
            cell = row.createCell((short)1);
            cell.setCellValue("Communication Plan");
            link = new HSSFHyperlink(HSSFHyperlink.LINK_DOCUMENT);
            link.setAddress("'CP'!A1");
            cell.setHyperlink(link);
            cell.setCellStyle(hlink_style);
            
            row = sheet1.createRow((short)15);
            cell = row.createCell((short)0);
            cell = row.createCell((short)1);
            cell.setCellValue("Configuration management plan");
            link = new HSSFHyperlink(HSSFHyperlink.LINK_DOCUMENT);
            link.setAddress("'CMP'!A1");
            cell.setHyperlink(link);
            cell.setCellStyle(hlink_style);
            
            row = sheet1.createRow((short)16);
            cell = row.createCell((short)0);
            cell = row.createCell((short)1);
            cell.setCellValue("Initiation Stage Plan");
            link = new HSSFHyperlink(HSSFHyperlink.LINK_DOCUMENT);
            link.setAddress("'ISP'!A1");
            cell.setHyperlink(link);
            cell.setCellStyle(hlink_style);

            row = sheet1.createRow((short)17);
            row = sheet1.createRow((short)18);
            cell = row.createCell((short)0);
            cell.setCellValue("Reports:");
            cell = row.createCell((short)1);
            cell.setCellValue("End Stage reports");
            link = new HSSFHyperlink(HSSFHyperlink.LINK_DOCUMENT);
            link.setAddress("'ESR'!A1");
            cell.setHyperlink(link);
            cell.setCellStyle(hlink_style);
            
            row = sheet1.createRow((short)19);
            cell = row.createCell((short)0);
            cell = row.createCell((short)1);
            cell.setCellValue("Exception reports");
            link = new HSSFHyperlink(HSSFHyperlink.LINK_DOCUMENT);
            link.setAddress("'EXR'!A1");
            cell.setHyperlink(link);
            cell.setCellStyle(hlink_style);

            row = sheet1.createRow((short)20);
            cell = row.createCell((short)0);
            cell = row.createCell((short)1);
            cell.setCellValue("Other reports");
            link = new HSSFHyperlink(HSSFHyperlink.LINK_DOCUMENT);
            link.setAddress("'OR'!A1");
            cell.setHyperlink(link);
            cell.setCellStyle(hlink_style);
            
            // sheet creation
            HSSFSheet sheet2 = wb.createSheet("Project Mandate");
            row = sheet2.createRow((short)0);
            cell = row.createCell((short)0);
            cell.setCellValue("To Index Page");
            link = new HSSFHyperlink(HSSFHyperlink.LINK_DOCUMENT);
            link.setAddress("'Index'!A1");
            cell.setHyperlink(link);
            cell.setCellStyle(hlink_style);

            HSSFSheet sheet3 = wb.createSheet("Business Case");
            row = sheet3.createRow((short)0);
            cell = row.createCell((short)0);
            cell.setCellValue("To Index Page");
            link = new HSSFHyperlink(HSSFHyperlink.LINK_DOCUMENT);
            link.setAddress("'Index'!A1");
            cell.setHyperlink(link);
            cell.setCellStyle(hlink_style);

            HSSFSheet sheet4 = wb.createSheet("Project Brief");
            row = sheet4.createRow((short)0);
            cell = row.createCell((short)0);
            cell.setCellValue("To Index Page");
            link = new HSSFHyperlink(HSSFHyperlink.LINK_DOCUMENT);
            link.setAddress("'Index'!A1");
            cell.setHyperlink(link);
            cell.setCellStyle(hlink_style);

            HSSFSheet sheet5 = wb.createSheet("Issue Log");
            row = sheet5.createRow((short)0);
            cell = row.createCell((short)0);
            cell.setCellValue("To Index Page");
            link = new HSSFHyperlink(HSSFHyperlink.LINK_DOCUMENT);
            link.setAddress("'Index'!A1");
            cell.setHyperlink(link);
            cell.setCellStyle(hlink_style);
            
            row = sheet5.createRow((short)1);
            cell = row.createCell((short)0);
            cell.setCellValue("Issue #");
            cell.setCellStyle(blue);
            cell = row.createCell((short)1);
            cell.setCellValue("Call #");
            cell.setCellStyle(blue);
            cell = row.createCell((short)2);
            cell.setCellValue("Reported");
            cell.setCellStyle(blue);
            cell = row.createCell((short)3);
            cell.setCellValue("Description");
            cell.setCellStyle(blue);
            sheet5.setColumnWidth((short)3, (short)10000);
            cell = row.createCell((short)4);
            cell.setCellValue("Owner");
            cell.setCellStyle(blue);
            cell = row.createCell((short)5);
            cell.setCellValue("Action By");
            cell.setCellStyle(blue);
            cell = row.createCell((short)6);
            cell.setCellValue("Status");
            cell.setCellStyle(blue);
            cell = row.createCell((short)7);
            cell.setCellValue("Detail");
            cell.setCellStyle(blue);
            sheet5.setColumnWidth((short)7, (short)10000);

            for (int i = 2; i < 101; i++) {
              row = sheet5.createRow((short)i);
              issuenumber = i - 1;
              cell = row.createCell((short)0);
              cell.setCellValue("" + issuenumber);
              cell.setCellStyle(blue);
            }
            
            HSSFSheet sheet6 = wb.createSheet("Daily Log");
            row = sheet6.createRow((short)0);
            cell = row.createCell((short)0);
            cell.setCellValue("To Index Page");
            link = new HSSFHyperlink(HSSFHyperlink.LINK_DOCUMENT);
            link.setAddress("'Index'!A1");
            cell.setHyperlink(link);
            cell.setCellStyle(hlink_style);

            HSSFSheet sheet7 = wb.createSheet("Lessons Learned Log");
            row = sheet7.createRow((short)0);
            cell = row.createCell((short)0);
            cell.setCellValue("To Index Page");
            link = new HSSFHyperlink(HSSFHyperlink.LINK_DOCUMENT);
            link.setAddress("'Index'!A1");
            cell.setHyperlink(link);
            cell.setCellStyle(hlink_style);
            
            HSSFSheet sheet8 = wb.createSheet("Quality Log");
            row = sheet8.createRow((short)0);
            cell = row.createCell((short)0);
            cell.setCellValue("To Index Page");
            link = new HSSFHyperlink(HSSFHyperlink.LINK_DOCUMENT);
            link.setAddress("'Index'!A1");
            cell.setHyperlink(link);
            cell.setCellStyle(hlink_style);

           
            HSSFSheet sheet9 = wb.createSheet("Risk Log");
            row = sheet9.createRow((short)0);
            cell = row.createCell((short)0);
            cell.setCellValue("To Index Page");
            link = new HSSFHyperlink(HSSFHyperlink.LINK_DOCUMENT);
            link.setAddress("'Index'!A1");
            cell.setHyperlink(link);
            cell.setCellStyle(hlink_style);

            row = sheet9.createRow((short)1);
            cell = row.createCell((short)0);
            cell.setCellValue("Risk #");
            cell.setCellStyle(blue);
            cell = row.createCell((short)1);
            cell.setCellValue("Created");
            cell.setCellStyle(blue);
            cell = row.createCell((short)2);
            cell.setCellValue("Description");
            cell.setCellStyle(blue);
            sheet9.setColumnWidth((short)2, (short)10000);
            cell = row.createCell((short)3);
            cell.setCellValue("Probability");
            cell.setCellStyle(blue);
            cell = row.createCell((short)4);
            cell.setCellValue("Impact");
            cell.setCellStyle(blue);
            cell = row.createCell((short)5);
            cell.setCellValue("Probablilty x Impact");
            cell.setCellStyle(blue);
            cell = row.createCell((short)6);
            cell.setCellValue("Action");
            cell.setCellStyle(blue);
            sheet9.setColumnWidth((short)6, (short)10000);


            for (int i = 2; i < 21; i++) {
              row = sheet9.createRow((short)i);
              issuenumber = i - 1;
              cell = row.createCell((short)0);
              cell.setCellValue("" + issuenumber);
              cell.setCellStyle(blue);
            }

            
            HSSFSheet sheet10 = wb.createSheet("Project Plan");
            row = sheet10.createRow((short)0);
            cell = row.createCell((short)0);
            cell.setCellValue("To Index Page");
            link = new HSSFHyperlink(HSSFHyperlink.LINK_DOCUMENT);
            link.setAddress("'Index'!A1");
            cell.setHyperlink(link);
            cell.setCellStyle(hlink_style);
            
            HSSFSheet sheet11 = wb.createSheet("PQP");
            row = sheet11.createRow((short)0);
            cell = row.createCell((short)0);
            cell.setCellValue("To Index Page");
            link = new HSSFHyperlink(HSSFHyperlink.LINK_DOCUMENT);
            link.setAddress("'Index'!A1");
            cell.setHyperlink(link);
            cell.setCellStyle(hlink_style);
            
            HSSFSheet sheet12 = wb.createSheet("CP");
            row = sheet12.createRow((short)0);
            cell = row.createCell((short)0);
            cell.setCellValue("To Index Page");
            link = new HSSFHyperlink(HSSFHyperlink.LINK_DOCUMENT);
            link.setAddress("'Index'!A1");
            cell.setHyperlink(link);
            cell.setCellStyle(hlink_style);
            
            HSSFSheet sheet13 = wb.createSheet("ISP");
            row = sheet13.createRow((short)0);
            cell = row.createCell((short)0);
            cell.setCellValue("To Index Page");
            link = new HSSFHyperlink(HSSFHyperlink.LINK_DOCUMENT);
            link.setAddress("'Index'!A1");
            cell.setHyperlink(link);
            cell.setCellStyle(hlink_style);
            
            HSSFSheet sheet14 = wb.createSheet("CMP");
            row = sheet14.createRow((short)0);
            cell = row.createCell((short)0);
            cell.setCellValue("To Index Page");
            link = new HSSFHyperlink(HSSFHyperlink.LINK_DOCUMENT);
            link.setAddress("'Index'!A1");
            cell.setHyperlink(link);
            cell.setCellStyle(hlink_style);
           
            HSSFSheet sheet15 = wb.createSheet("ESR");
            row = sheet15.createRow((short)0);
            cell = row.createCell((short)0);
            cell.setCellValue("To Index Page");
            link = new HSSFHyperlink(HSSFHyperlink.LINK_DOCUMENT);
            link.setAddress("'Index'!A1");
            cell.setHyperlink(link);
            cell.setCellStyle(hlink_style);
            
            HSSFSheet sheet16 = wb.createSheet("EXR");
            row = sheet16.createRow((short)0);
            cell = row.createCell((short)0);
            cell.setCellValue("To Index Page");
            link = new HSSFHyperlink(HSSFHyperlink.LINK_DOCUMENT);
            link.setAddress("'Index'!A1");
            cell.setHyperlink(link);
            cell.setCellStyle(hlink_style);
            
            HSSFSheet sheet17 = wb.createSheet("OR");
            row = sheet17.createRow((short)0);
            cell = row.createCell((short)0);
            cell.setCellValue("To Index Page");
            link = new HSSFHyperlink(HSSFHyperlink.LINK_DOCUMENT);
            link.setAddress("'Index'!A1");
            cell.setHyperlink(link);
            cell.setCellStyle(hlink_style);

            sheet1.autoSizeColumn((short)1);
            sheet1.autoSizeColumn((short)9);
            
            fileOut = new FileOutputStream(FileName);
            wb.write(fileOut);
            fileOut.close();
            
        } catch (IOException ex) {
            Logger.getLogger(SimpleP2View.class.getName()).log(Level.SEVERE, null, ex);
        } finally {
            try {
                fileOut.close();
            } catch (IOException ex) {
                Logger.getLogger(SimpleP2View.class.getName()).log(Level.SEVERE, null, ex);
            }
        }

    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.ButtonGroup buttonGroup1;
    private javax.swing.JButton jButton1;
    private javax.swing.JButton jButton2;
    private javax.swing.JButton jButton3;
    private javax.swing.JCheckBox jCheckBox1;
    private javax.swing.JCheckBox jCheckBox10;
    private javax.swing.JCheckBox jCheckBox2;
    private javax.swing.JCheckBox jCheckBox3;
    private javax.swing.JCheckBox jCheckBox4;
    private javax.swing.JCheckBox jCheckBox5;
    private javax.swing.JCheckBox jCheckBox6;
    private javax.swing.JCheckBox jCheckBox7;
    private javax.swing.JCheckBox jCheckBox8;
    private javax.swing.JCheckBox jCheckBox9;
    private javax.swing.JComboBox jComboBox1;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JLabel jLabel3;
    private javax.swing.JLabel jLabel4;
    private javax.swing.JLabel jLabel5;
    private javax.swing.JLabel jLabel6;
    private javax.swing.JLabel jLabel7;
    private javax.swing.JLabel jLabel8;
    private javax.swing.JLabel jLabel9;
    private javax.swing.JMenuItem jMenuItem5;
    private javax.swing.JPanel jPanel1;
    private javax.swing.JRadioButton jRadioButton1;
    private javax.swing.JRadioButton jRadioButton2;
    private javax.swing.JTextField jTextField1;
    private javax.swing.JTextField jTextField2;
    private javax.swing.JTextField jTextField3;
    private javax.swing.JTextField jTextField4;
    private javax.swing.JTextField jTextField5;
    private javax.swing.JTextField jTextField6;
    private javax.swing.JTextField jTextField7;
    private javax.swing.JPanel mainPanel;
    private javax.swing.JMenuBar menuBar;
    private javax.swing.JLabel statusAnimationLabel;
    private javax.swing.JLabel statusMessageLabel;
    private javax.swing.JPanel statusPanel;
    // End of variables declaration//GEN-END:variables
    private String[] custnames = new String[500];

    private final Timer messageTimer;
    private final Timer busyIconTimer;
    private final Icon idleIcon;
    private final Icon[] busyIcons = new Icon[15];
    private int busyIconIndex = 0;

    private JDialog aboutBox;

    private void load_customers() {
        readafile(customerfile);      
    }

    private void load_ini_setting() {
        File f_ini = new File ("settings.ini");
        FileInputStream fini = null;
        try {
            fini = new FileInputStream(f_ini);
        } catch (FileNotFoundException ex) {
            ex.printStackTrace();
        }
              Properties p = new Properties();
        try {
            p.load(fini);
        } catch (IOException ex) {
            ex.printStackTrace();
        }
        try {
            fini.close();
        } catch (IOException ex) {
            ex.printStackTrace();
        }
        jTextField1.setText(p.getProperty("DEFDIR"));
        defdir = p.getProperty("DEFDIR");
        customerfile = p.getProperty("CUSTFILE");
        jTextField5.setText(p.getProperty("DEF_PM"));
        jTextField6.setText(p.getProperty("DEF_PM_PHONE"));
        jTextField7.setText(p.getProperty("DEF_PM_MAIL"));
        // System.out.println(customerfile);

    }

    private void readafile(String bestand) {
        BufferedReader dis = null;
        // DataInputStream dis = null; 
        String record = null; 
        int recCount = 0; 

        try { 

           File f = new File(bestand); 
           FileInputStream fis = new FileInputStream(f); 
           BufferedInputStream bis = new BufferedInputStream(fis); 
           dis = new BufferedReader(new InputStreamReader(bis));
           // dis = new DataInputStream(bis);  

           while ( (record= dis.readLine()) != null ) { 
              recCount++;
                try {
                    mySplitResult = record.split(";");
                    String custnum = mySplitResult[0];
                    custname = "";
                    for (int i = 1; i < mySplitResult.length ; i++) {
                       if (i == 1) {
                        custname = mySplitResult[i];
                       } else {
                        custname = custname + ", " + mySplitResult[i];
                       }
                    }
                    jComboBox1.addItem(custnum);
                    ii = jComboBox1.getItemCount();
                    
                    custnames[ii - 1] = custname;      
                    // System.out.println("" + recCount + " " + custnum + "  ,  " + custname); 
                } finally {
                }
              
              
           } 

        } catch (IOException e) { 
           // catch io errors from FileInputStream or readLine() 
           System.out.println("IOException error!" + e.getMessage()); 

        } finally { 
           // if the file opened okay, make sure we close it 
           if (dis != null) { 
	      try {
                 dis.close(); 
	      } catch (IOException ioe) {
	      }
           } 
        } 

    }
}
