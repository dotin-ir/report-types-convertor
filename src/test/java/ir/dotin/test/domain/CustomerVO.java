package ir.dotin.test.domain;

import java.io.Serializable;
import java.util.List;

/**
 * Created by r.rastakfard on 7/21/2016.
 */
public class CustomerVO implements Serializable {

    private Long customerNumber;
    private String nationalCode;
    private List<AddressVO> addresses;
    private String birthLocation;

    public CustomerVO(Long customerNumber, String nationalCode) {
        this.customerNumber = customerNumber;
        this.nationalCode = nationalCode;
    }

    public Long getCustomerNumber() {
        return customerNumber;
    }

    public void setCustomerNumber(Long customerNumber) {
        this.customerNumber = customerNumber;
    }

    public String getNationalCode() {
        return nationalCode;
    }

    public void setNationalCode(String nationalCode) {
        this.nationalCode = nationalCode;
    }

    public List<AddressVO> getAddresses() {
        return addresses;
    }

    public void setAddresses(List<AddressVO> addresses) {
        this.addresses = addresses;
    }

    public String getBirthLocation() {
        return birthLocation;
    }

    public void setBirthLocation(String birthLocation) {
        this.birthLocation = birthLocation;
    }
}
