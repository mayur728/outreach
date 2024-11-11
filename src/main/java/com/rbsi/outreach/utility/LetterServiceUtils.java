package com.rbsi.outreach.utility;

import com.rbsi.outreach.model.CustomerData;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.springframework.stereotype.Service;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

@Service
public class LetterServiceUtils {

    private static final String TEMPLATE_PATH = "C:\\Users\\FD RBSI notifications\\letter_template.docx"; // Set your actual template path
    private static final String OUTPUT_DIRECTORY = "C:\\Users\\FD RBSI notifications\\"; // Set your output directory

    public List<CustomerData> loadCustomerData(String filePath) {
        List<CustomerData> customerDataList = new ArrayList<>();

        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheet("Sheet1"); // Always read from "Sheet1"

            // Start reading from the third row (index 2)
            Iterator<Row> rowIterator = sheet.rowIterator();
            for (int i = 0; i < 2; i++) {
                if (rowIterator.hasNext()) {
                    rowIterator.next(); // Skip first two header rows
                }
            }

            // Iterate through rows and populate CustomerData
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                CustomerData customer = new CustomerData();

                customer.setCustomerId(getCellStringValue(row.getCell(0)));
                customer.setFullName(getCellStringValue(row.getCell(1)));
                customer.setFirstName(getCellStringValue(row.getCell(2)));
                customer.setAnalyst(getCellStringValue(row.getCell(3)));
                customer.setCustomerType(getCellStringValue(row.getCell(4)));
                customer.setBrand(getCellStringValue(row.getCell(5)));
                customer.setJurisdiction(getCellStringValue(row.getCell(6)));
                customer.setDaysNotice((int) row.getCell(8).getNumericCellValue()); // Assuming numeric value
                customer.setLetterDate(getCellDateValue(row.getCell(9)));
                customer.setResponseDate(getCellDateValue(row.getCell(10)));
                customer.setWebformResponse(getCellBooleanValue(row.getCell(11)));
                customer.setHooyouResponse(getCellBooleanValue(row.getCell(12)));
                customer.setSalutation(getCellStringValue(row.getCell(13)));
                customer.setAddressLine1(getCellStringValue(row.getCell(14)));
                customer.setAddressLine2(getCellStringValue(row.getCell(15)));
                customer.setCity(getCellStringValue(row.getCell(16)));
                customer.setCountry(getCellStringValue(row.getCell(17)));
                customer.setPostalCode(getCellStringValue(row.getCell(18)));
                customer.setDocumentPath(getCellStringValue(row.getCell(19)));
                customer.setStatus(getCellStringValue(row.getCell(20)));

                customerDataList.add(customer);
            }

        } catch (IOException e) {
            e.printStackTrace();
            // Handle exception, e.g., log error or throw custom exception
        }

        return customerDataList;
    }
    private String getCellStringValue(Cell cell) {
        if (cell == null) {
            return "";
        }
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                return String.valueOf((int) cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            default:
                return "";
        }
    }

    private String getCellDateValue(Cell cell) {
        if (cell == null || cell.getCellType() != CellType.NUMERIC || !DateUtil.isCellDateFormatted(cell)) {
            return "";
        }
        SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");
        return dateFormat.format(cell.getDateCellValue());
    }

    private boolean getCellBooleanValue(Cell cell) {
        if (cell == null) {
            return false;
        }
        return cell.getCellType() == CellType.BOOLEAN ? cell.getBooleanCellValue() : Boolean.parseBoolean(getCellStringValue(cell));
    }
    public String generateLetter(CustomerData customer) {
        String fileName = "Letter_" + customer.getFullName() + "_" + customer.getAnalyst() + ".docx";
        String outputFilePath = OUTPUT_DIRECTORY + fileName;

        try (FileInputStream fis = new FileInputStream(TEMPLATE_PATH);
             XWPFDocument document = new XWPFDocument(fis);
             FileOutputStream fos = new FileOutputStream(outputFilePath)) {

            // Iterate through paragraphs to replace placeholders
            for (XWPFParagraph paragraph : document.getParagraphs()) {
                for (XWPFRun run : paragraph.getRuns()) {
                    String text = run.getText(0);
                    if (text != null) {
                        text = text.replace("<FullName>", customer.getFullName())
                                .replace("<FirstName>", customer.getFirstName())
                                .replace("<Branch>", customer.getAnalyst())
                                .replace("<CustomerType>", customer.getCustomerType())
                                .replace("<Salutation>", customer.getSalutation())
                                .replace("<AddressLine1>", customer.getAddressLine1())
                                .replace("<AddressLine2>", customer.getAddressLine2())
                                .replace("<City>", customer.getCity())
                                .replace("<Country>", customer.getCountry())
                                .replace("<PostalCode>", customer.getPostalCode());
                        run.setText(text, 0);
                    }
                }
            }

            // Save the modified document to the output directory
            document.write(fos);

        } catch (IOException e) {
            e.printStackTrace();
        }

        return fileName;
    }

    public void generateDoc2Email(List<CustomerData> customerDataList, String outputFilePath) {
        try (Workbook workbook = new XSSFWorkbook();
             FileOutputStream fos = new FileOutputStream(outputFilePath)) {

            Sheet sheet = workbook.createSheet("Sheet1");
            Row headerRow = sheet.createRow(0);

            // Create header row
            headerRow.createCell(0).setCellValue("Customer ID");
            headerRow.createCell(1).setCellValue("Full Name");
            headerRow.createCell(2).setCellValue("Email");
            headerRow.createCell(3).setCellValue("Branch");
            headerRow.createCell(4).setCellValue("Document Path");
            headerRow.createCell(5).setCellValue("Output");

            // Populate rows with customer data
            int rowIndex = 1;
            for (CustomerData customer : customerDataList) {
                Row row = sheet.createRow(rowIndex++);
                row.createCell(0).setCellValue(customer.getCustomerId());
                row.createCell(1).setCellValue(customer.getFullName());
                row.createCell(2).setCellValue(customer.getEmail());
                row.createCell(3).setCellValue(customer.getAnalyst());
                row.createCell(4).setCellValue(customer.getDocumentPath());
                row.createCell(5).setCellValue("Pending"); // Set initial status as Pending
            }

            // Write the workbook to the output file
            workbook.write(fos);

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public void updateDoc2EmailStatus(String filePath, List<CustomerData> updatedCustomers) {
        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis);
             FileOutputStream fos = new FileOutputStream(filePath)) {

            Sheet sheet = workbook.getSheet("Sheet1");
            SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
            String sentDate = dateFormat.format(new Date());

            // Iterate through rows and update the Output column for each customer
            for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                String customerId = row.getCell(0).getStringCellValue();

                for (CustomerData customer : updatedCustomers) {
                    if (customer.getCustomerId().equals(customerId)) {
                        row.createCell(5).setCellValue("Sent: " + sentDate);
                    }
                }
            }

            // Write the updated workbook to the output file
            workbook.write(fos);

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
