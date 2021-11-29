package Reading;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadCellExample {
    private static final String fileName = "D:\\NguyenDongHung.xlsx";
    private static int row = 0;
    private static int col = 0;
    public static void main(String[] args)
    {
        ReadCellExample rc=new ReadCellExample();
        String vOutput=rc.ReadCellData(row, col);
        System.out.println(vOutput);
    }
    //method defined for reading a cell
    public String ReadCellData(int vRow, int vColumn)
    {
        String value=null;
        Workbook wb=null;
        try
        {
            FileInputStream fis=new FileInputStream(fileName);
            wb=new XSSFWorkbook(fis);
        }
        catch(FileNotFoundException e)
        {
            e.printStackTrace();
        }
        catch(IOException e1)
        {
            e1.printStackTrace();
        }
        Sheet sheet=wb.getSheetAt(0);
        Row row=sheet.getRow(vRow);
        Cell cell=row.getCell(vColumn);
        value=cell.getStringCellValue();
        return value;        
    }
}
