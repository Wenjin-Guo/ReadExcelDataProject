import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;

public class readExcelData {

    static String filePath = "C:\\DataSampleFiles\\sampledatahockey.xlsx";
    static int numOfRows;
    static int numOfColumns;
    static String[][] excelDataArray;

    public static void main (String[] args) throws Exception {

        readDataFromExcel(filePath);

        for (int i =1; i<numOfRows; i++){
            String ID = excelDataArray[i][0];
            String Team = excelDataArray[i][1];
            String Country = excelDataArray[i][2];
            String firstName = excelDataArray[i][3];
            String lastName = excelDataArray[i][4];
            String Weight = excelDataArray[i][5];
            String Height = excelDataArray[i][6];
            String DOB= excelDataArray[i][7];
            String homeTown = excelDataArray[i][8];
            String Province = excelDataArray[i][9];
            String Position = excelDataArray[i][10];
            String Age= excelDataArray[i][11];
            String HeightFT = excelDataArray[i][12];
            String  Htln= excelDataArray[i][13];
            String BMI = excelDataArray[i][14];

            System.out.println("|ID: "+ ID + "|Team: "+Team+"|Country: "+Country+"|firstName: "+firstName+"|lastName: "+lastName+"|Weight: "+Weight+"|Height: "+Height+"|DOB: "+DOB+"|homeTown: "+homeTown+"|Province: "+Province+"|Position: "+Position+"|Age: "+Age+"|HeightFT: "+HeightFT+"|Htln: "+Htln+"|BMI: "+BMI);

            /*
            add web automation process here, parameter is already setup
             */


        }




    }

    public static void readDataFromExcel(String filePath) throws Exception {
        File xlFile = new File(filePath);
        FileInputStream testDataStream = new FileInputStream(filePath);

        XSSFWorkbook wBook = new XSSFWorkbook(testDataStream);
        XSSFSheet wSheet = wBook.getSheetAt(1);
        numOfRows = wSheet.getLastRowNum() + 1;
        numOfColumns = wSheet.getRow(0).getLastCellNum();

        System.out.println("Excel File name " + xlFile.getName());
        System.out.println("Total number of rows are "+ numOfRows);
        System.out.println();
        System.out.println("Total number of columns are "+ numOfColumns);
        System.out.println();

        excelDataArray = new String[numOfRows][numOfColumns];

        for (int i =0; i < numOfRows; i++){
            XSSFRow wRow = wSheet.getRow(i);
            for(int j = 0; j < numOfColumns; j++){
                XSSFCell cellData = wRow.getCell(j);
                cellData.setCellType(CellType.STRING);
                String cellValue = cellData.getStringCellValue();
                excelDataArray[i][j] = cellValue;
//                System.out.println(cellValue);
            }
        }


    }
}
