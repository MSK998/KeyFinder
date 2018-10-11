package KeyFinder;

import java.io.IOException;

/**
 * KeyFinder.KeyFinder was created by Team 2 on 04/10/2018.
 *
 * Mark Scott-Kiddie
 * Lee Robbie
 * Lewis Ross
 * Reuben Hadden
 * Scott Guy
 *
 */
public class KeyFinderApp extends javax.swing.JFrame {
    public static void main(String[] args) throws IOException {

        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new KeyFinderGUI().setVisible(true);
            }
        });



        //Testing Methods
        KeyFinder sheet1 = new KeyFinder();
        sheet1.loadData();

        System.out.println("Displaying Array Now");

        sheet1.displayArray();

        System.out.println("Outputting cell:");

        sheet1.displaySpecific(2,1);
        
        
        KeyWriter sheet2 = new KeyWriter();


        //How will we get the ArrayList from KeyFinder to KeyWriter classes

        /*
        Work on

        Search
        Add Fields
        Linking GUI to current classes
        BrainStorm new features??
         */

        System.out.println("_____________________________________________________________________________________");


    }

    private javax.swing.JButton editButton;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JLabel jLabel7;
    private javax.swing.JLabel jLabel8;
    private javax.swing.JLabel jLabel9;
    private javax.swing.JMenu jMenu1;
    private javax.swing.JMenuBar jMenuBar1;
    private javax.swing.JPanel jPanel1;
    private javax.swing.JScrollPane jScrollPane1;
    private javax.swing.JTable jTable1;
    private javax.swing.JButton lostKeysBtn;
    private javax.swing.JButton searchByRoom;
    private javax.swing.JButton searchFobsBtn;
    private javax.swing.JButton searchKeysBtn;
    private java.awt.TextArea textArea1;
    // End of variables declaration//GEN-END:variables
}
