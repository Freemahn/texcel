package freemahn;

/**
 * Created by pgordon on 07.03.2017.
 */
public class Employee {
    private String name;
    private String costCenter;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getCostCenter() {
        return costCenter;
    }

    public void setCostCenter(String costCenter) {
        this.costCenter = costCenter;
    }

    public Employee(String name, String costCenter) {

        this.name = name;
        this.costCenter = costCenter;
    }
}
