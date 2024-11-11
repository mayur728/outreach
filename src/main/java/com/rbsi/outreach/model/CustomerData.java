package com.rbsi.outreach.model;


import lombok.*;

@Getter
@Setter
@EqualsAndHashCode
@NoArgsConstructor
@AllArgsConstructor
public class CustomerData {
    private String customerId;
    private String fullName;
    private String firstName;
    private String email;
    private String analyst;
    private String customerType;
    private String brand;
    private String jurisdiction;
    private int daysNotice;
    private String letterDate;
    private String responseDate;
    private boolean webformResponse;
    private boolean hooyouResponse;
    private String salutation;
    private String addressLine1;
    private String addressLine2;
    private String city;
    private String country;
    private String postalCode;
    private String documentPath;
    private String status;

    // Getters and setters...
}

