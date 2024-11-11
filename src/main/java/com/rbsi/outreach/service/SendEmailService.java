package com.rbsi.outreach.service;

import com.rbsi.outreach.model.CustomerData;
import com.rbsi.outreach.utility.LetterServiceUtils;
import jakarta.mail.MessagingException;
import jakarta.mail.internet.MimeMessage;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.mail.javamail.JavaMailSender;
import org.springframework.mail.javamail.MimeMessageHelper;
import org.springframework.stereotype.Service;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

@Service
public class SendEmailService {

    @Autowired
    private JavaMailSender mailSender;

    @Autowired
    private LetterServiceUtils letterServiceUtils;

    public void sendEmails() {
        String doc2EmailPath = "path/to/output/directory/doc2Email.xlsx";
        // Load data from doc2Email.xlsx
        List<CustomerData> emailDataList = letterServiceUtils.loadCustomerData(doc2EmailPath);
        List<CustomerData> updatedCustomers = new ArrayList<>();

        // Iterate through the list and send an email for each entry
        for (CustomerData customer : emailDataList) {
            try {
                sendEmail(customer);
                updatedCustomers.add(customer);
                System.out.println("Email sent to: " + customer.getEmail());
            } catch (MessagingException e) {
                System.err.println("Failed to send email to: " + customer.getEmail());
                e.printStackTrace();
            }
        }

        // Update the status in doc2Email.xlsx
        letterServiceUtils.updateDoc2EmailStatus(doc2EmailPath, updatedCustomers);
    }

    private void sendEmail(CustomerData customer) throws MessagingException {
        MimeMessage message = mailSender.createMimeMessage();
        MimeMessageHelper helper = new MimeMessageHelper(message, true);

        // Set email details
        helper.setTo(customer.getEmail());
        helper.setSubject(getEmailSubject(customer.getAnalyst()));
        helper.setText("Dear " + customer.getSalutation() + " " + customer.getFullName() + ",\n\nPlease find the attached document.\n\nBest regards,\nCustomer Support Team");

        // Attach the document if available
        File document = new File(customer.getDocumentPath());
        if (document.exists()) {
            helper.addAttachment(document.getName(), document);
        } else {
            System.err.println("Document not found for customer: " + customer.getFullName());
        }

        // Send email
        mailSender.send(message);
    }

    private String getEmailSubject(String branch) {
        // Determine subject based on branch (similar to VB logic)
        if ("NWI".equalsIgnoreCase(branch)) {
            return "NatWest International - Request for Important Information";
        } else if ("IOMB".equalsIgnoreCase(branch) || "IOM".equalsIgnoreCase(branch)) {
            return "Isle of Man Bank - Request for Important Information";
        } else {
            return "Request for Important Information";
        }
    }
}


