
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.ArrayList;
import java.util.Dictionary;
import java.util.HashMap;
import java.util.TreeMap;

/**
 * Created by dgotbaum on 1/7/14.
 */
public class Market {

    public static final String[] HEADER = {"Submarket & Class", "Number of Buildings","Inventory","Direct Available Space"
            ,"Direct Availability","Sublet Available Space","Sublet Availability","Total Available Space","Total Availability",
            "Direct Vacant Space","Direct Vacancy","Sublet Vacant Space","Sublet Vacancy","Total Vacant Space","Total Vacancy","Occupied Space","Net Absorption",
            "Weighted Direct Average Rent","Weighted Sublease Average Rent","Weighted Overall Average Rent","Under Construction","Under Construction (SF)"};

    public int marketIndex;
    public int subMarketIndex;
    public int classIndex;
    public Workbook wb;
    public Sheet s;
    public TreeMap<String, TreeMap<String,TreeMap<String,ArrayList<Row>>>> MARKETS;
    public Market(Workbook wb) {
        this.MARKETS = new TreeMap<String, TreeMap<String, TreeMap<String, ArrayList<Row>>>>();
        this.wb = wb;
        this.s = wb.getSheetAt(0);
        Row headerRow = s.getRow(0);
        for (Cell c : headerRow) {
            if (c.getRichStringCellValue().getString().contains("Market (my data)"))
                this.marketIndex = c.getColumnIndex();
            else if (c.getRichStringCellValue().getString().contains("submarket"))
                this.subMarketIndex = c.getColumnIndex();
            else if (c.getRichStringCellValue().getString().contains("Class"))
                this.classIndex = c.getColumnIndex();

        }
        //Iterates through the rows and populates the hashmap of the buildings by district
        for (Row r: this.s) {
            if (!(r.getRowNum() == 0) && r.getCell(0) != null) {
                String marketName = r.getCell(marketIndex).getStringCellValue();
                String subMarketName = r.getCell(subMarketIndex).getStringCellValue();
                String Grade = r.getCell(classIndex).getStringCellValue();



                if (!MARKETS.containsKey(marketName)) {
                    TreeMap<String, TreeMap<String,ArrayList<Row>>> sub  = new TreeMap<String, TreeMap<String,ArrayList<Row>>>();
                    TreeMap<String, ArrayList<Row>> classes = new TreeMap<String, ArrayList<Row>>();
                    classes.put("A", new ArrayList<Row>());
                    classes.put("B", new ArrayList<Row>());
                    classes.put("C", new ArrayList<Row>());
                    classes.get(Grade).add(r);
                    sub.put(subMarketName, classes);
                    MARKETS.put(marketName,sub);
                }
                else {
                    if (MARKETS.get(marketName).containsKey(subMarketName))
                        MARKETS.get(marketName).get(subMarketName).get(Grade).add(r);
                    else {
                        TreeMap<String, ArrayList<Row>> classes = new TreeMap<String, ArrayList<Row>>();
                        classes.put("A", new ArrayList<Row>());
                        classes.put("B", new ArrayList<Row>());
                        classes.put("C", new ArrayList<Row>());
                        classes.get(Grade).add(r);
                        MARKETS.get(marketName).put(subMarketName, classes);
                    }
                }
            }

        }
    }
}
