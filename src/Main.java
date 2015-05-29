import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.List;

/**
 * Created by prabhavadhikary on 5/22/15.
 */
public class Main {

    public static void main(String[] args) throws FileNotFoundException, IOException  {

        try {
            //load's and reads 2 excel files,
            FileInputStream file1 = new FileInputStream("test.xlsx");
            FileInputStream file2 = new FileInputStream("test2.xlsx");

            //load workbooks from FileInputStream variables declared above:
            XSSFWorkbook workbook1 = new XSSFWorkbook(file1);
            XSSFWorkbook workbook2 = new XSSFWorkbook(file2);

            //load sheets from XSSFWorkbook variables declared above:
            XSSFSheet sheet1 = workbook1.getSheetAt(0);
            XSSFSheet sheet2 = workbook2.getSheetAt(0);

            //declaration of CompareServices Class object:
            CompareServices service = new CompareServices();

            //call getCells method from CompareService Class to extract cell of sheet1
            List<Sheet> sht1 = service.getCells(sheet1);

            //call getCells method from CompareService Class to extract cell of sheet2
            List<Sheet> sht2 = service.getCells(sheet2);

            //call compare method from CompareService Class.
            service.compare(sht1, sht2);


        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

    }

}
