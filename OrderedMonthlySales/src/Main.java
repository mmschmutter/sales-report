import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.*;

import java.io.*;
import java.util.*;

/**
 * Parses data from Word document containing sales data,
 * groups sales by month, orders by quantity sold, and
 * prints ordered monthly data to new Word document.
 *
 * @author Mordechai Schmutter
 * @version 1.0
 */

public class Main {
    public static void main(String[] args) throws Exception {
        try {
            System.out.println("Enter directory of Sales file:");
            BufferedReader reader = new BufferedReader(new InputStreamReader(System.in)); // open command line input reader
            String directory = reader.readLine(); // ask user for command line input
            reader.close(); // close command line input reader
            XWPFDocument document = new XWPFDocument(OPCPackage.open(directory)); // open Word document
            TreeSet<Product> products = new TreeSet<>(parse(document)); // parse data and store in TreeSet
            print(products, document); // print ordered monthly data to new Word document
            System.out.println("Results printed in 'Result.docx'");
        } catch (Exception e) {
            System.out.println("Sales file not found in that directory");
        }
    }

    /**
     * Parses data from Word document into a Collection of Products that are grouped by month
     *
     * @param document Word document to be parsed
     * @return Collection<Product> Products from Word document, grouped by month
     */
    static Collection<Product> parse(XWPFDocument document) {
        HashMap<String, Product> products = new HashMap<>(); // products sold by month
        XWPFTable table = document.getTables().get(0); // get first table in document
        for (int i = 1; i < table.getNumberOfRows(); i++) { // iterate over rows, starting after headers
            XWPFTableRow row = table.getRow(i);
            String name = row.getCell(0).getText(); // get product name
            String[] date = row.getCell(1).getText().split("/"); // split date data
            int month = Integer.parseInt(date[1]); // get month
            int year = Integer.parseInt(date[2]); // get year
            int quantity = Integer.parseInt(row.getCell(2).getText()); // get quantity
            Product product = new Product(name, month, year, quantity); // create product object from data
            if (products.containsKey(product.getID())) { // if product-month combination already exists
                products.get(product.getID()).increaseQuantity(product.getQuantity()); // increase quantity by this date's amount
            } else { // if product-month combination does not yet exist
                products.put(product.getID(), product); // add product-month combination
            }
        }
        return products.values();
    }

    /**
     * Prints Products grouped by month and sorted by sale quantity to new Word document
     *
     * @param products Products grouped by month and sorted by sale quantity
     * @param document Word document that data was parsed from
     */
    static void print(TreeSet<Product> products, XWPFDocument document) throws Exception {
        XWPFParagraph title = document.getParagraphs().get(0); // get title line of Word document
        XWPFRun setTitle = title.createRun(); // edit title line
        setTitle.setText(" (ordered by quantity per month)");
        document.removeBodyElement(document.getPosOfTable(document.getTables().get(0))); // delete original table
        XWPFTable table = document.createTable(products.size() + 1, 3); // create new table
        String[] headers = {"Product Name", "Sale Date", "Sale Quantity"};
        for (int i = 0; i < 3; i++) { // iterate over each cell in header row
            XWPFParagraph cell = table.getRow(0).getCell(i).getParagraphs().get(0);
            XWPFRun setHeaders = cell.createRun(); // set header
            setHeaders.setBold(true); // set header formatting to bold
            setHeaders.setText(headers[i]); // set header text
        }
        int row = 1;
        for (Product product : products) { // add ordered product data to table cells
            table.getRow(row).getCell(0).setText(product.getName()); // add product name
            table.getRow(row).getCell(1).setText(product.getMonth() + "/" + product.getYear()); // product date
            table.getRow(row).getCell(2).setText(product.getQuantity() + ""); // add product quantity
            row++;
        }
        document.write(new FileOutputStream("Result.docx")); // write data to new Word document
    }
}