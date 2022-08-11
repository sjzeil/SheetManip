/**
 * 
 */
package edu.odu.cs.sheetManip;

/**
 * A spreadsheet cell name, e.g., "A21"
 * 
 * @author zeil
 *
 */
public class CellName {
    
    private String name;
    private int column;
    private int row;
    
    public CellName (String st) {
        name = "A1";
        column  = 0;
        row = 0;

        int k = st.length()-1;
        while (k >= 0 && Character.isDigit(st.charAt(k))) {
            --k;
        }
        if (k < 0) return;
        
        for (int i = 0; i <= k; ++i) {
            char c = st.charAt(i);
            if (c < 'A' || c > 'Z')
                return;
        }

        try {
            row = Integer.parseInt(st.substring(k+1)) - 1;
        } catch (NumberFormatException ex) {
            return;
        }
        column = ((int)st.charAt(k)) - (int)'A'; 
        --k;
        if (k >= 0) {
            int m = ((int)st.charAt(k)) - (int)'A';
            column += 26 * (m+1);
        }
        name=st;
    }

    /**
     * @return the name
     */
    public String getName() {
        return name;
    }

    /**
     * @return the column
     */
    public int getColumn() {
        return column;
    }

    /**
     * @return the row
     */
    public int getRow() {
        return row;
    }
}
