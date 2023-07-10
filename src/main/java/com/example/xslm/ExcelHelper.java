package ma.ensaj.medisafe.helper;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import ma.ensaj.medisafe.beans.Medicaments;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.multipart.MultipartFile;

public class ExcelHelper {
    public static String TYPE = "application/vnd.ms-excel.sheet.macroenabled.12";
    static String SHEET = "Fiche_de_Travail";

    public static boolean hasExcelFormat(MultipartFile file) {
        System.out.println(file.getContentType());

        if (!TYPE.equals(file.getContentType())) {
            return false;
        }

        return true;
    }



    public static List<Medicaments> excelToTutorials(InputStream is) {
        boolean input = false;
        try {
            Workbook workbook = new XSSFWorkbook(is);

            Sheet sheet = workbook.getSheet(SHEET);

            Iterator<Row> rows = sheet.iterator();


            List<Medicaments> medicaments = new ArrayList<Medicaments>();

            int rowNumber = 20;
            hamza:
            while (rows.hasNext()) {
                Row currentRow = rows.next();


                // skip header
                if (rowNumber == 20) {
                    rowNumber++;
                    continue;
                }
                if (rowNumber == 25) {
                    break;
                }

                Iterator<Cell> cellsInRow = currentRow.iterator();



                int cellIdx = 0;
                while (cellsInRow.hasNext()) {
                    Cell currentCell = cellsInRow.next();

                   try {
                       if(currentCell.getStringCellValue().equals("INPUT")){
                           System.out.println("INPUT");
                           input = true;
                           rows.next();
                           break;
                       } else if (currentCell.getStringCellValue().equals("Intermediate Output")) {
                           input = false;
                           break hamza;

                       }
                   }catch (Exception e){

                   }
                    if(input){
                        switch (cellIdx) {
                            case 2:
                                System.out.println(currentCell.getStringCellValue());
                                break;
                            default:
                                break;
                        }
                    }

                    cellIdx++;
                }

                //medicaments.add(medicament);
            }

            workbook.close();

            return medicaments;
        } catch (IOException e) {
            throw new RuntimeException("fail to parse Excel file: " + e.getMessage());
        }
    }
}