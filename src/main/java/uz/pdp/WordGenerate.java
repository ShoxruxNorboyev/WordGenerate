package uz.pdp;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;

import java.io.*;
import java.math.BigInteger;

public class WordGenerate {
    public static void main(String[] args) throws IOException, InvalidFormatException {
//        simpleText();
        generateTable();
    }

//    public static void simpleText() throws IOException {
//        XWPFDocument document = new XWPFDocument();
//        XWPFParagraph paragraph = document.createParagraph();
//        XWPFRun run = paragraph.createRun();
//        run.setText("Hello From Java ");
//        FileOutputStream out = new FileOutputStream(new File("D:\\Teaching\\G50\\JavaAdvanced\\simpleWordText.docx"));
//        document.write(out);
//        document.close();
//        out.close();
//    }

    public static void generateTable() throws IOException, InvalidFormatException {
        XWPFDocument xwpfDocument = new XWPFDocument();
        XWPFTable table = xwpfDocument.createTable();
        XWPFTableRow row = table.getRow(0);
         row.getCell(0).setText("T/R");

       row.addNewTableCell().setText("Mahsulot Nomi");

       row.addNewTableCell().setText("Rasmi");

       row.addNewTableCell().setText("Soni");
       row.addNewTableCell();
        mergeCellsHorizontal(table, 0, 3, 4);
        int rowCount = 1;
        table.createRow();
        table.createRow();
        table.createRow();
        table.createRow();
        table.createRow();
        table.createRow();
        table.createRow();
        for (int i = 0; i < 2; i++) {

           table.getRow(rowCount).getCell(0).setText(""+(i+1));
            mergeCellVertically(table, 0,rowCount,rowCount+2);

            row.addNewTableCell().setText("Novei1");
            mergeCellVertically(table, 1,rowCount,rowCount+2);
            mergeCellVertically(table, 2,rowCount,rowCount+2);

            XWPFParagraph paragraph1 = table.getRow(1).getCell(2).addParagraph();
            XWPFRun run1 = paragraph1.createRun();
            FileInputStream fis = new FileInputStream("D:\\All\\Other\\novey\\products\\katalog 4.jpg");
            run1.addPicture(fis, XWPFDocument.PICTURE_TYPE_JPEG, "Name", Units.pixelToEMU(70), Units.pixelToEMU(70));
            for (int j = 0; j < 3; j++) {
                row.addNewTableCell().setText("Shop "+(j+1));
                 row.addNewTableCell().setText(""+(j+1));
                if (j < 2){
                    row = table.createRow();
                    rowCount++;
                }else {
                    rowCount++;
                }
            }
        }
        FileOutputStream fileOutputStream = new FileOutputStream(new File("D:\\Teaching\\G50\\JavaAdvanced\\wordProductCount.docx"));
        xwpfDocument.write(fileOutputStream);
        xwpfDocument.close();
        fileOutputStream.close();

    }
//    public static void mergeCellVertically(XWPFTable table, int col, int fromRow, int toRow) {
//        for (int rowIndex = fromRow; rowIndex <= toRow; rowIndex++) {
//            XWPFTableCell cell = table.getRow(rowIndex).getCell(col);
//            if (rowIndex == fromRow) {
//                // The first merged cell is set with RESTART merge value
//                cell.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.RESTART);
//            } else {
//                // Cells which join (merge) the first one, are set with CONTINUE
//                cell.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.CONTINUE);
//            }
//        }
//    }
public static void mergeCellVertically(XWPFTable table, int col, int startRow, int endRow) {
    for (int i = startRow; i <= endRow; i++) {
        XWPFTableCell cell = table.getRow(i).getCell(col);
        if (i == startRow) {
            // The first merged cell is set with RESTART merge value
            cell.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.RESTART);
        } else {
            // Cells which join (merge) the first one, are set with CONTINUE
            cell.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.CONTINUE);
        }
    }
}

    static void mergeCellHorizontally(XWPFTable table, int row, int fromCol, int toCol) {
        XWPFTableCell cell = table.getRow(row).getCell(fromCol);
        // Try getting the TcPr. Not simply setting an new one every time.
        CTTcPr tcPr = cell.getCTTc().getTcPr();
        if (tcPr == null) tcPr = cell.getCTTc().addNewTcPr();
        // The first merged cell has grid span property set
        if (tcPr.isSetGridSpan()) {
            tcPr.getGridSpan().setVal(BigInteger.valueOf(toCol - fromCol + 1));
        } else {
            tcPr.addNewGridSpan().setVal(BigInteger.valueOf(toCol - fromCol + 1));
        }
        // Cells which join (merge) the first one, must be removed
        for (int colIndex = toCol; colIndex > fromCol; colIndex--) {
            table.getRow(row).getCtRow().removeTc(colIndex);
            table.getRow(row).removeCell(colIndex);
        }
    }

    public static void mergeCellsHorizontal(XWPFTable table, int row, int startCell, int endCell) {
        for (int i = startCell; i <= endCell; i++) {
            XWPFTableCell cell = table.getRow(row).getCell(i);
            if (i == startCell) {
                // The first merged cell is set with RESTART merge value
                cell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.RESTART);
            } else {
                // Cells which join (merge) the first one, are set with CONTINUE
                cell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.CONTINUE);
            }
        }
    }

}
