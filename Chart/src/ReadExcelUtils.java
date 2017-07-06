

import java.io.FileInputStream;  
import java.io.FileNotFoundException;  
import java.io.IOException;  
import java.io.InputStream;  
import java.math.BigDecimal;
import java.util.Date;  
import java.util.HashMap;  
import java.util.Map;  

import org.apache.poi.hssf.usermodel.HSSFWorkbook;  
import org.apache.poi.ss.usermodel.Cell;  
import org.apache.poi.ss.usermodel.DateUtil;  
import org.apache.poi.ss.usermodel.Row;  
import org.apache.poi.ss.usermodel.Sheet;  
import org.apache.poi.ss.usermodel.Workbook;  
import org.apache.poi.xssf.usermodel.XSSFWorkbook;    
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class ReadExcelUtils {  
    private Logger logger = LoggerFactory.getLogger(ReadExcelUtils.class);  
    private Workbook wb;  
    private Sheet sheet;  
    private Row row;  
  
    public ReadExcelUtils(String filepath) {  
        if(filepath==null){  
            return;  
        }  
        String ext = filepath.substring(filepath.lastIndexOf("."));  
        try {  
            InputStream is = new FileInputStream(filepath);  
            if(".dat".equals(ext)||".xlsx".equals(ext)){  
                wb = new XSSFWorkbook(is);  
            }else if (".xls".equals(ext)) {
				wb = new HSSFWorkbook(is);
			}else{  
                wb=null;  
            }  
        } catch (FileNotFoundException e) {  
            logger.error("FileNotFoundException", e);  
        } catch (IOException e) {  
            logger.error("IOException", e);  
        }  
    } 
    
    public String[] readExcelTitle() throws Exception{  
        if(wb==null){  
            throw new Exception("Workbook对象为空！");  
        }  
        sheet = wb.getSheetAt(0);  
        row = sheet.getRow(0);  
        // 标题总列数  
        int colNum = row.getPhysicalNumberOfCells();  
        System.out.println("colNum:" + colNum);  
        String[] title = new String[colNum];  
        for (int i = 0; i < colNum; i++) {  
            // title[i] = getStringCellValue(row.getCell((short) i));  
            title[i] = row.getCell(i).getCellFormula();  
        }  
        return title;  
    }  
    public Map<Integer, Map<Integer,Object>> readExcelContent() throws Exception{  
        if(wb==null){  
            throw new Exception("Workbook对象为空！");  
        }  
        Map<Integer, Map<Integer,Object>> content = new HashMap<Integer, Map<Integer,Object>>();  
        sheet = wb.getSheetAt(0);  
        // 得到总行数  
        int rowNum = sheet.getLastRowNum(); 
        row = sheet.getRow(0);
        // 正文内容应该从第二行开始,第一行为表头的标题 
        int z=1;
        for (int i = 1; i<=rowNum; i++) {  
            row = sheet.getRow(i);
            if(row!=null){
            	int colNum=row.getLastCellNum();
            	Map<Integer,Object> cellValue = new HashMap<Integer, Object>();
            	int m=0;
                for(int k=0;k<colNum;k++){
	                if(row.getCell(k)==null){
	                	
	                }else if(getCellFormatValue(row.getCell(k)).equals("")){
	                	 
	                }else{
	                	Object obj = getCellFormatValue(row.getCell(k));  
	               	    cellValue.put(m, obj);
	               	    m++;
	                }
               	
                }
                content.put(z, cellValue); 
                z++;
            }
             
        }  
        return content;  
    }  
    private Object getCellFormatValue(Cell cell) {  
        Object cellvalue = "";  
        if (cell != null) {  
            // 判断当前Cell的Type  
            switch (cell.getCellType()) {  
            case Cell.CELL_TYPE_NUMERIC: 
            case Cell.CELL_TYPE_FORMULA: {  
                // 判断当前的cell是否为Date  
                if (DateUtil.isCellDateFormatted(cell)) {  
                    Date date = cell.getDateCellValue();  
                    cellvalue = date;  
                } else { 
                   
                    if(String.valueOf(cell.getNumericCellValue()).indexOf("E")>0){
                    	BigDecimal bd= new BigDecimal(String.valueOf(cell.getNumericCellValue()));
                    	cellvalue= bd.toPlainString();
                    }else{
                    	cellvalue = String.valueOf(cell.getNumericCellValue());
                    }
                    
                }  
                break;  
            }  
            case Cell.CELL_TYPE_STRING: 
                cellvalue = cell.getRichStringCellValue().getString();  
                break;  
            default: 
                cellvalue = "";  
            }  
        } else {  
            cellvalue = "";  
        }  
        return cellvalue;  
    }  
}  