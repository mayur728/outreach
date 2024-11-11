package com.rbsi.outreach.service;

import com.rbsi.outreach.model.CustomerData;
import com.rbsi.outreach.utility.LetterServiceUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.util.ArrayList;
import java.util.List;

@Service
public class GenerateLetterService {

    @Autowired
    private LetterServiceUtils letterServiceUtils;

    private static final String OUTPUT_DIRECTORY = "path/to/output/directory/"; // Set your output directory

    public void generateLetters() {
        // Load customer data
        List<CustomerData> customerDataList = letterServiceUtils.loadCustomerData("C:\\Users\\FD RBSI notifications\\customer_data_sample.xlsx");
        List<CustomerData> eligibleCustomers = new ArrayList<>();

        // Generate letters for each eligible customer
        for (CustomerData customer : customerDataList) {
            if ("Next".equalsIgnoreCase(customer.getStatus())) {
                String filename = letterServiceUtils.generateLetter(customer);
                customer.setDocumentPath(OUTPUT_DIRECTORY + filename);
                eligibleCustomers.add(customer);
                System.out.println("Generated letter: " + filename);
            }
        }

        // Generate doc2Email.xlsx for eligible customers
        if (!eligibleCustomers.isEmpty()) {
            String doc2EmailPath = "C:\\Users\\FD RBSI notifications\\doc2Email.xlsx";
            letterServiceUtils.generateDoc2Email(eligibleCustomers, doc2EmailPath);
            System.out.println("Generated doc2Email.xlsx at: " + doc2EmailPath);
        }
    }
}

