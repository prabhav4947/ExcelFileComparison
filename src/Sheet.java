/**
 * Created by prabhavadhikary on 5/22/15.
 */
public class Sheet {

    //following are variables defining each column of excel workbook.

    private String donor;
    private String organization;
    private String description;
    private String pledge;
    private String fund;

    //variable declaring row numbers of excel sheets.
    private int rowNumber;

    //getter and setter for all variables declared above

    public String getDonor() {
        return donor;
    }

    public void setDonor(String donor) {
        this.donor = donor;
    }

    public String getOrganization() {
        return organization;
    }

    public void setOrganization(String organization) {
        this.organization = organization;
    }


    public String getDescription() {
        return description;
    }

    public void setDescription(String description) {
        this.description = description;
    }

    public int getRowNumber() {
        return rowNumber;
    }

    public void setRowNumber(int rowNumber) {
        this.rowNumber = rowNumber;
    }

    public String getPledge() {
        return pledge;
    }

    public void setPledge(String pledge) {
        this.pledge = pledge;
    }

    public String getFund() {
        return fund;
    }

    public void setFund(String fund) {
        this.fund = fund;
    }

}
