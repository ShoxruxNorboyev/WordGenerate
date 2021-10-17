package uz.pdp.excelExample;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.*;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

public class ExcelGenerate {
    public static void main(String[] args) throws IOException {
//        studentList();
        productCount();
    }

    public static void studentList() throws IOException {
        List<Student> studentList = new ArrayList<>();
        studentList.add(new Student(1, "Abdulloh", "+998991234567", 20));
        studentList.add(new Student(2, "Abdurrohman", "+998991234568", 21));
        studentList.add(new Student(3, "Abdurroshid", "+998991234569", 22));
        studentList.add(new Student(4, "Abdulbosit", "+998991234560", 23));
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFFont font = workbook.createFont();
        font.setBold(true);
        XSSFCellStyle style = workbook.createCellStyle();
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setFont(font);
        XSSFSheet sheet = workbook.createSheet();
        XSSFRow row1 = sheet.createRow(0);
        XSSFCell row1Cell1 = row1.createCell(0);
        row1Cell1.setCellValue("ID");
        row1Cell1.setCellStyle(style);
        sheet.autoSizeColumn(row1Cell1.getColumnIndex());
        XSSFCell row1Cell2 = row1.createCell(1);
        row1Cell2.setCellValue("FIO");
        row1Cell2.setCellStyle(style);
        sheet.autoSizeColumn(row1Cell2.getColumnIndex());
        XSSFCell row1Cell3 = row1.createCell(2);
        row1Cell3.setCellValue("PhoneNumber");
        row1Cell3.setCellStyle(style);
        sheet.autoSizeColumn(row1Cell3.getColumnIndex());
        XSSFCell row1Cell4 = row1.createCell(3);
        row1Cell4.setCellValue("Age");
        row1Cell4.setCellStyle(style);
        sheet.autoSizeColumn(row1Cell4.getColumnIndex());
        font.setBold(false);
        style.setFont(font);
        for (int i = 0; i < studentList.size(); i++) {
            XSSFRow row = sheet.createRow(i + 1);
            XSSFCell cell = row.createCell(0);
            cell.setCellValue(studentList.get(i).getId());
            cell.setCellStyle(style);
            sheet.autoSizeColumn(cell.getColumnIndex());
            cell = row.createCell(1);
            cell.setCellStyle(style);
            cell.setCellValue(studentList.get(i).getName());
            sheet.autoSizeColumn(cell.getColumnIndex());
            cell = row.createCell(2);
            cell.setCellStyle(style);
            cell.setCellValue(studentList.get(i).getPhoneNumber());
            sheet.autoSizeColumn(cell.getColumnIndex());
            cell = row.createCell(3);
            cell.setCellStyle(style);
            cell.setCellValue(studentList.get(i).getAge());
            sheet.autoSizeColumn(cell.getColumnIndex());
        }

        File file = new File("D:\\Teaching\\G50\\JavaAdvanced\\studentList.xlsx");
        FileOutputStream fileOutputStream = new FileOutputStream(file);
        workbook.write(fileOutputStream);
        workbook.close();
        fileOutputStream.close();
    }

    public static void productCount() throws IOException {
        List<ProductCount> productCountList = new ArrayList<>();
        ProductCount productCount = new ProductCount();
        productCount.setName("Novey1");
        productCount.setImgUrl("D:\\All\\Other\\novey\\products\\katalog 4.jpg");
        List<ProductShopCount> productShopCountList=new ArrayList<>();
        ProductShopCount productShopCount = new ProductShopCount();
        productShopCount.setShopName("Shop1");
        productShopCount.setAmount(20);
        productShopCountList.add(productShopCount);
        productShopCount = new ProductShopCount();
        productShopCount.setShopName("Shop2");
        productShopCount.setAmount(30);
        productShopCountList.add(productShopCount);
        productShopCount = new ProductShopCount();
        productShopCount.setShopName("Shop3");
        productShopCount.setAmount(40);
        productShopCountList.add(productShopCount);
        productCount.setProductShopCountList(productShopCountList);
        productCountList.add(productCount);
        productCount = new ProductCount();
        productCount.setName("Novey2");
        productCount.setImgUrl("D:\\All\\Other\\novey\\products\\katalog 6.jpg");
        productShopCountList=new ArrayList<>();
        productShopCount = new ProductShopCount();
        productShopCount.setShopName("Shop1");
        productShopCount.setAmount(10);
        productShopCountList.add(productShopCount);
        productShopCount = new ProductShopCount();
        productShopCount.setShopName("Shop2");
        productShopCount.setAmount(12);
        productShopCountList.add(productShopCount);
        productShopCount = new ProductShopCount();
        productShopCount.setShopName("Shop3");
        productShopCount.setAmount(15);
        productShopCountList.add(productShopCount);
        productCount.setProductShopCountList(productShopCountList);
        productCountList.add(productCount);
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFCellStyle cellStyle = workbook.createCellStyle();
        XSSFFont font = workbook.createFont();
        font.setBold(true);
        cellStyle.setFont(font);
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        XSSFSheet sheet = workbook.createSheet();
        XSSFRow row = sheet.createRow(0);

        XSSFCell cell = row.createCell(0);
        cell.setCellValue("T/R");
        cell.setCellStyle(cellStyle);
        sheet.autoSizeColumn(cell.getColumnIndex());

        cell = row.createCell(1);
        cell.setCellValue("Maxsulot Nomi");
        cell.setCellStyle(cellStyle);
        sheet.autoSizeColumn(cell.getColumnIndex());

        cell = row.createCell(2);
        cell.setCellValue("Rasmi");
        cell.setCellStyle(cellStyle);
        sheet.autoSizeColumn(cell.getColumnIndex());

        cell = row.createCell(3);
        cell.setCellValue("Soni");
        cell.setCellStyle(cellStyle);
        sheet.autoSizeColumn(cell.getColumnIndex());
        sheet.addMergedRegion(new CellRangeAddress(0,0,3,4));

        int rowCount=1;
        for (int i = 0; i < productCountList.size(); i++) {
            ProductCount pc = productCountList.get(i);
            row = sheet.createRow(rowCount);

            cell = row.createCell(0);
            cell.setCellValue(i+1);
            cell.setCellStyle(cellStyle);
            sheet.autoSizeColumn(cell.getColumnIndex());
            sheet.addMergedRegion(new CellRangeAddress(rowCount,(rowCount+pc.getProductShopCountList().size())-1,0,0));

            cell = row.createCell(1);
            cell.setCellValue(pc.getName());
            cell.setCellStyle(cellStyle);
            sheet.autoSizeColumn(cell.getColumnIndex());
            sheet.addMergedRegion(new CellRangeAddress(rowCount,(rowCount+pc.getProductShopCountList().size())-1,1,1));

            cell = row.createCell(2);

            InputStream inputStream = new FileInputStream(new File(pc.getImgUrl()));
            //Get the contents of an InputStream as a byte[].
            byte[] bytes = IOUtils.toByteArray(inputStream);
            //Adds a picture to the workbook
            int pictureIdx = workbook.addPicture(bytes, Workbook.PICTURE_TYPE_JPEG);
            //close the input stream
            inputStream.close();
            //Returns an object that handles instantiating concrete classes
            CreationHelper helper = workbook.getCreationHelper();

            //Creates the top-level drawing patriarch.
            Drawing drawing = sheet.createDrawingPatriarch();

            //Create an anchor that is attached to the worksheet
            ClientAnchor anchor = helper.createClientAnchor();
            //set top-left corner for the image
            anchor.setCol1(2);
            anchor.setRow1(rowCount);
            //Creates a picture
            Picture pict = drawing.createPicture(anchor, pictureIdx);
            //Reset the image to the original size
            pict.resize(1,pc.getProductShopCountList().size());
            sheet.addMergedRegion(new CellRangeAddress(rowCount,(rowCount+pc.getProductShopCountList().size())-1,2,2));
            cell.setCellStyle(cellStyle);
            sheet.autoSizeColumn(cell.getColumnIndex());

            for (int i1 = 0; i1 < pc.getProductShopCountList().size(); i1++) {
                ProductShopCount shopCount = pc.getProductShopCountList().get(i1);

                cell = row.createCell(3);
                cell.setCellValue(shopCount.getShopName());
                cell.setCellStyle(cellStyle);
                sheet.autoSizeColumn(cell.getColumnIndex());

                cell = row.createCell(4);
                cell.setCellValue(shopCount.getAmount());
                cell.setCellStyle(cellStyle);
                sheet.autoSizeColumn(cell.getColumnIndex());

                if (i1< pc.getProductShopCountList().size()-1){
                    row=sheet.createRow(++rowCount);
                }else {
                    rowCount++;
                }
            }

        }

        FileOutputStream fileOutputStream=new FileOutputStream(new File("D:\\Teaching\\G50\\JavaAdvanced\\productCount.xlsx"));
        workbook.write(fileOutputStream);
        workbook.close();
        fileOutputStream.close();

    }
}
