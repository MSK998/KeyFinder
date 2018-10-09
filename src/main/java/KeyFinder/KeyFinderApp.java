package KeyFinder;

import KeyFinder.KeyFinder;

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
public class KeyFinderApp {
    public static void main(String[] args) {


        //Testing Methods
        KeyFinder sheet1 = new KeyFinder();
        sheet1.loadData();

        System.out.println("Displaying Array Now");

        sheet1.displayArray();

        System.out.println("Outputting cell 2,1");

        sheet1.displaySpecific(2,1);

        //How will we get the ArrayList from KeyFinder.KeyFinder to KeyWriter classes

        /*
        Work on

        Search
        Add Fields
        Linking GUI to current classes
        BrainStorm new features??
         */
    }
}
