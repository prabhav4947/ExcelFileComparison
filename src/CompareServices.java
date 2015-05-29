import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by prabhavadhikary on 5/22/15.
 */
public class CompareServices{

    //method extracts data from each row and cell of sheet (defined in method argument),
    //extracted data is than added to List of Object (Sheet)
    //extracted data is also printed.. in case of large files please comment out print data section from line 62.

    public List<Sheet> getCells(XSSFSheet sheet) {

        XSSFRow row1 = sheet.getRow(0);

        //List of Object
        List<Sheet> sht = new ArrayList<Sheet>();

        int lastRow = sheet.getLastRowNum();

        for (int i = 9; i <= lastRow; i++) {

            row1 = sheet.getRow(i);

            Sheet sht1obj = new Sheet();

            int y = 0;

            // extraction of data
            XSSFCell cell1 = row1.getCell(y);
            XSSFCell cell2 = row1.getCell(y + 1);
            XSSFCell cell3 = row1.getCell(y + 2);
            XSSFCell cell4 = row1.getCell(y + 3);
            XSSFCell cell5 = row1.getCell(y + 4);

            //adding data to list object
            sht1obj.setDonor(String.valueOf(cell1));
            sht1obj.setOrganization(String.valueOf(cell2));
            sht1obj.setDescription(String.valueOf(cell3));
            sht1obj.setPledge(String.valueOf(cell4));
            sht1obj.setFund(String.valueOf(cell5));
            sht1obj.setRowNumber(i + 1);

            //indicates end of entries which require comparison
            if ((cell1.getCellType() == XSSFCell.CELL_TYPE_BLANK) && (cell2.getCellType() == XSSFCell.CELL_TYPE_BLANK) && (cell3.getCellType() == XSSFCell.CELL_TYPE_BLANK) &&
                    (cell4.getCellType() == XSSFCell.CELL_TYPE_BLANK) && (cell5.getCellType()) == (XSSFCell.CELL_TYPE_BLANK))
            {
                break;
            }

            sht.add(sht1obj);

        }

        //print data * ------------------ *
        System.out.println("*--- Sheet ---*");

        for (Sheet val : sht) {

            System.out.print(val.getRowNumber());
            System.out.print("\t" + val.getDonor());
            System.out.print("\t" + val.getOrganization());
            System.out.print("\t" + val.getDescription());
            System.out.print("\t" + val.getPledge());
            System.out.print("\t" + val.getFund());

            System.out.println("");

        }

        return sht;

    }

    //method compares two List Objects pass in arguments. (Lists include extraction of cell data from sheet1 and sheet2
    //the method prints the changes made in sheet1 to sheet2.. through values set in mark variables.

    public void compare(List<Sheet> st1, List<Sheet> st2){

        System.out.println("------");

        //initialization of mark's.
        String newDonor = null;
        String organization = null;
        String fund = null;
        String pledge = null;


        for (Sheet sheet1: st1) {

            newDonor = sheet1.getDonor(); //mark for new Donor
            organization = sheet1.getOrganization(); // mark for new Organization

            for (Sheet sheet2: st2) {

                if (sheet1.getDonor().equals(sheet2.getDonor())) {

                    newDonor = null;//new Donor mark reset to null

                    if (sheet1.getOrganization().equals(sheet2.getOrganization())) {

                        organization = null;//new organization mark reset to null

                        if (sheet1.getDescription().equals(sheet2.getDescription())) {

                            if(!sheet1.getPledge().equals(sheet2.getPledge())){
                                fund = sheet1.getFund(); //mark for change in funding
                            }
                            if(!sheet1.getFund().equals(sheet2.getFund())){
                                pledge = sheet1.getPledge(); // mark for change in pledge
                            }

                            if((sheet1.getDescription().equals(sheet2.getDescription())) && (sheet1.getPledge().equals(sheet2.getPledge())) && (sheet1.getFund().equals(sheet2.getFund()))){
                                System.out.println("Row No." + (sheet1.getRowNumber()) + " of Sheet 1 is equal to Row No." + (sheet2.getRowNumber()) + " of Sheet 2");
                                fund = null;
                                pledge = null;

                                //if row x is of sheet1 is equal to row x of sheet2 breaks loop to start checking for values in next row.
                                break;
                            }
                        }
                    }
                }
            }


            // if a new Donor mark has been initialized and not reset, prints new Donor Name.
            if (newDonor != null) {
                System.out.println("*****");
                System.out.println("New Donor (" + sheet1.getDonor() + ") listed at Row No." + (sheet1.getRowNumber()));
                System.out.println("*****");
                organization = null;
                fund = null;
                pledge= null;

            }

            // if a new organization mark has been initialized and not reset, prints new organization Name.
            if (organization != null) {
                System.out.println("######");
                System.out.println("New organization (" + sheet1.getOrganization() + ") listed under Donor (" + sheet1.getDonor() + ") at Row No." + (sheet1.getRowNumber()));
                System.out.println("######");
            }

            //prints change in funding between sheet1 and sheet2
            if (fund != null) {
                System.out.println("---------");
                System.out.println("Funding has changed under Donor (" + sheet1.getDonor() + ") and organization (" + sheet1.getOrganization() + ") at Row No." + (sheet1.getRowNumber()));
                System.out.println("Updated fund is: "+sheet1.getFund());
                System.out.println("---------");

            }

            //prints change in pledge between sheet1 and sheet2
            if (pledge != null) {
                System.out.println("---------");
                System.out.println("Pledge has changed under Donor (" + sheet1.getDonor() + ") and organization (" + sheet1.getOrganization() + ") at Row No." + (sheet1.getRowNumber()));
                System.out.println("Updated pledge is: "+sheet1.getPledge());
                System.out.println("---------");

            }
        }
    }

}
