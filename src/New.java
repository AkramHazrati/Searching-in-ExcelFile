import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
public class New {


        public static void main(String args[]) throws IOException
        {
//obtaining input bytes from a file
            FileInputStream fis=new FileInputStream(new File("C:\\Users\\Iran\\Desktop\\numbers.xlsx"));
//creating workbook instance that refers to .xls file
            XSSFWorkbook wb=new XSSFWorkbook(fis);
//creating a Sheet object to retrieve the object
            XSSFSheet sheet=wb.getSheetAt(0);

//evaluating cell type
            FormulaEvaluator formulaEvaluator=wb.getCreationHelper().createFormulaEvaluator();
            for(Row row: sheet)     //iteration over row using for each loop
            {
                for(Cell cell: row)    //iteration over cell using for each loop
                {
                   // if(row.equals(0) && cell.equals()){
                     //   System.out.println("2");
                    //}
                   // System.out.println("s");
                    switch(formulaEvaluator.evaluateInCell(cell).getCellType())
                    {
                        case Cell.CELL_TYPE_NUMERIC:   //field that represents numeric cell type
//getting the value of the cell as a number


                           int cellcol = cell.getColumnIndex();
                           int cellrow = cell.getRowIndex();

                           if(cellcol==1 && cellrow == 1){
                               System.out.println();
                               System.out.println();
                               System.out.println();
                               System.out.print(cell.getNumericCellValue()+ "\t\t");
                               System.out.println();
                               System.out.println();
                               System.out.println();
                           }
                            // System.out.print(cell.getNumericCellValue()+ "\t\t");
                            break;
                        case Cell.CELL_TYPE_STRING:    //field that represents string cell type
//getting the value of the cell as a string
                            //System.out.print(cell.getStringCellValue()+ "\t\t");
                            break;
                    }
                }
                System.out.println();
            }
        }
    }

