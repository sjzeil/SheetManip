/**
 * 
 */
package edu.odu.cs.sheetManip;

import static org.junit.jupiter.api.Assertions.*;
import org.junit.jupiter.api.*;


/**
 * @author zeil
 *
 */
public class TestCellName {
    

    @Test
    public void testCellName() {
        CellName cn = new CellName("A1");
        assertEquals("A1", cn.getName());
        assertEquals(0, cn.getColumn());
        assertEquals(0, cn.getRow());
        
        cn = new CellName("B3");
        assertEquals("B3", cn.getName());
        assertEquals(1, cn.getColumn());
        assertEquals(2, cn.getRow());
        
        cn = new CellName("AB3");
        assertEquals("AB3", cn.getName());
        assertEquals(27, cn.getColumn());
        assertEquals(2, cn.getRow());
        
        
    }

}
