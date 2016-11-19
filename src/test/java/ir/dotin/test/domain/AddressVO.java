package ir.dotin.test.domain;

import java.io.Serializable;

/**
 * Created by r.rastakfard on 8/10/2016.
 */
public class AddressVO implements Serializable {

    private String type;
    private String address;

    public AddressVO(String type, String address) {
        this.type = type;
        this.address = address;
    }

    public String getType() {
        return type;
    }

    public void setType(String type) {
        this.type = type;
    }

    public String getAddress() {
        return address;
    }

    public void setAddress(String address) {
        this.address = address;
    }
}
