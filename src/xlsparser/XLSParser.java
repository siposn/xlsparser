package xlsparser;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.Format;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class XLSParser {

    //Összesített adatok
    private final Workbook newWorkbook;
    //Munkafüzet
    private final Sheet newSheet;
    //Címsor
    private final Row header;
    //Összes xls kérdéseit (címsor lesz belőle) tartalmazó set
    private final Set<String> headerRow;
    //Beolvasandó mappa lista
    private final List<File> folders;
    //Betegenkénti kérdés-válasz párok
    private final List<Map<String,String>> dataSet;
    
    public XLSParser() {
        this.newWorkbook = new XSSFWorkbook();
        this.newSheet = this.newWorkbook.createSheet("Betegek");
        this.header = this.newSheet.createRow(0);
        this.headerRow = new LinkedHashSet<>();
        this.folders = new ArrayList<>();
        this.dataSet = new ArrayList<>();
    }

    public static void main(String[] args) {
        XLSParser parser = new XLSParser();
        parser.start();
    }
    
    public void start() {
        folders.add(new File("C:\\Users\\Norbi\\Desktop\\new"));

        folders.stream().filter((folder) -> (folder.exists())).forEach((folder) -> {
            openFolder(folder);
        });        
    }
     
    public void openFolder(File folder) {
        for(File entry : folder.listFiles()) {
            if(!entry.isDirectory()) {
                createDataSet(entry);
            }
        }
        createXlsX();
    }
    
    public void createDataSet(File file) {
       try {
           //Egy beteg munkafüzetében levé kérdés-válasz párok
           Map<String,String> rows = new HashMap<>();
           
           try (FileInputStream xlsx = new FileInputStream(file)) {
               XSSFWorkbook workbook = new XSSFWorkbook(xlsx);
               XSSFSheet sheet = workbook.getSheetAt(0);
               int rowNum = sheet.getLastRowNum()+1;
               for(int i = 0; i<rowNum; i++) {
                   if(sheet.getRow(i) != null) {
                       Cell c = sheet.getRow(i).getCell(0);
                       if(c != null && !c.getStringCellValue().isEmpty()) {
                           //címsor halmazhoz kérdés hozzáadása
                           headerRow.add(c.getStringCellValue());
                           //ha a sor nem üres
                           if(sheet.getRow(i).getCell(1) != null) {
                               if(sheet.getRow(i).getCell(1).getCellType() == XSSFCell.CELL_TYPE_STRING) {
                                   //eltároljuk map-ben a kérdést és hozzá a választ
                                   rows.put(c.getStringCellValue(), sheet.getRow(i).getCell(1).getStringCellValue());
                               } else if(sheet.getRow(i).getCell(1).getCellType() == XSSFCell.CELL_TYPE_NUMERIC) {
                                   if(DateUtil.isCellDateFormatted(sheet.getRow(i).getCell(1))) {
                                       //ha dátum mező
                                       Format formatter = new SimpleDateFormat("yyyy.MM.dd");                                     
                                       rows.put(c.getStringCellValue(), formatter.format(sheet.getRow(i).getCell(1).getDateCellValue()));
                                   }else {
                                       rows.put(c.getStringCellValue(), sheet.getRow(i).getCell(1).getRawValue());
                                   }
                               }
                           }
                       }
                   }
               }
               //beteg adatainak hozzáadása a listához
               dataSet.add(rows);
           }
       } catch(IOException ex) {
         Logger.getLogger(XLSParser.class.getName()).log(Level.SEVERE, null, ex); 
       }
    }
    
    public void createXlsX() {
        try {
            //XLSX összeállítás
            try (FileOutputStream fos = new FileOutputStream("betegek.xlsx")) {
                //Header összeállítás
                int i = 0;
                for(String headerR : headerRow) {
                    header.createCell(i).setCellValue(headerR);
                    i++;
                }
                
                //Betegszámmal megegyező sorok létrehozása
                for(int j = 1; j<dataSet.size()+1; j++) {
                    this.newSheet.createRow(j);
                }
                
                //Sorok feltöltése adatokkal
                //Betegenként lépés
                for(int k = 0; k<dataSet.size(); k++) {
                    //Aktuális beteg adatai (kérdés-válasz)
                    Map<String,String> data = dataSet.get(k);
                    //Címsor olvasása és a címsor aktuális kérdésének megfelelő érték kivétele
                    int hr = 0;
                    for(String h : headerRow) {
                        this.newSheet.getRow(k+1).createCell(hr).setCellValue(data.get(h));
                        hr++;
                    }
                }
                this.newWorkbook.write(fos);
                fos.close();
                System.out.println("Az XLS elkészült!");
            }
            
        } catch (IOException ex) {
            Logger.getLogger(XLSParser.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
}
