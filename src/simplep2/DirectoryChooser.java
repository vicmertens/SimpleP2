/*
 * DirectoryChooser.java
 *
 * Created on 19 juni 2007, 9:45
 *
 * To change this template, choose Tools | Template Manager
 * and open the template in the editor.
 */

package simplep2;

import javax.swing.*;
import java.awt.event.*;
import java.awt.*;
import java.util.*;

/** A class for creating a directory chooser. Just a simple
 * "file browser" where a directory for storing logged video
 * can be selected.
 *
 * @author  Jan Lindblom <linjan-1@student.luth.se>
 * @version 1.1
 */
public class DirectoryChooser extends JPanel {
    private JFileChooser chooser;
    private String selectedFile;
    private String fsPath;
    
    /** Creates a new instance of DirectoryChooser */
    public DirectoryChooser() {
        chooser = new JFileChooser();
        chooser.setCurrentDirectory(new java.io.File("."));
        chooser.setDialogTitle("Select a directory");
        chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
        if (chooser.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
            /* a directory was selected */
            try {
                /* get the fs path of this and store as a String. */
                selectedFile = chooser.getSelectedFile().toString();
            }
            catch (Exception e) {}
        }
        else {
            /* This occurs when nothing is selected and cancel is pressed */
            selectedFile = "None";
        }
    }
    
    /** Returns the selected directory as a String.
     * @return a String with the current selection.
     */    
    public String getSelection() {
        return selectedFile;
    }
}
