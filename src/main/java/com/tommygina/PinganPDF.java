package com.tommygina;

import com.itextpdf.kernel.colors.DeviceRgb;
import com.itextpdf.kernel.font.PdfFont;
import com.itextpdf.kernel.font.PdfFontFactory;
import com.itextpdf.kernel.geom.PageSize;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfPage;
import com.itextpdf.kernel.pdf.PdfReader;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.kernel.pdf.canvas.parser.PdfTextExtractor;
import com.itextpdf.kernel.pdf.canvas.parser.listener.SimpleTextExtractionStrategy;
import com.itextpdf.layout.Document;
import com.itextpdf.layout.element.Paragraph;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.BufferedReader;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.StringReader;

class PinganPDF {

    private static final float nDiv = (float) 3.9;
    private static final String excelFilePath = "pingan.xlsx";

    // 创建单个pdf页面
    private void createPage(Document document, float pageHeight, PdfFont font) {


        createFixedContent(document, pageHeight, font);
        // createVariableContent();
    }

    private void readExcel(String excelFilePath) {


    }

    // 固定格式文件
    private void createFixedContent(Document document, float pageHeight, PdfFont font) {
        System.out.println("Creating Fixed Content");
        // row1 column1
        Paragraph p1 = new Paragraph("客户存款月结单")
                .setFont(font)
                .setFontSize(11)
                .setFontColor(new DeviceRgb(0, 0, 0))
                // .setFixedPosition((float)222.5, (float)(pageHeight - nDiv -  45), (float)77);
                .setFixedPosition((float) 222.5, (float) (pageHeight - nDiv - 41.0), (float) 77);
        document.add(p1);

        // row1 column2
        Paragraph p2 = new Paragraph("结单号：2412310251999000006599")
                .setFont(font)
                .setFontSize(11)
                .setFontColor(new DeviceRgb(0, 0, 0))
                .setFixedPosition((float) 342.8, (float) (pageHeight - nDiv - 41.0), (float) 155.80402);
        document.add(p2);

        // row1 column3
        Paragraph p3 = new Paragraph("2024年12月")
                .setFont(font)
                .setFontSize(11)
                .setFontColor(new DeviceRgb(0, 0, 0))
                .setFixedPosition((float) 583.4, (float) (pageHeight - nDiv - 41.0), (float) 52.492);

        document.add(p3);

        // row1 column4
        Paragraph p4 = new Paragraph("第1页  共94页")
                .setFont(font)
                .setFontSize(11)
                .setFontColor(new DeviceRgb(0, 0, 0))
                .setFixedPosition((float) 663.6, (float) (pageHeight - nDiv - 41.0), (float) 63.800003);
        document.add(p4);

        // row2 column1
        Paragraph p5 = new Paragraph("客户行:平安银行杭州分行营业部")
                .setFont(font)
                .setFontSize(11)
                .setFontColor(new DeviceRgb(0, 0, 0))
                .setFixedPosition((float) 22.0, (float) (pageHeight - nDiv - 61.0), (float) 156.618);
        document.add(p5);

        // row2 column2
        Paragraph p6 = new Paragraph("户名:上海华通铂银交易市场有限公司")
                .setFont(font)
                .setFontSize(11)
                .setFontColor(new DeviceRgb(0, 0, 0))
                .setFixedPosition((float) 289.33, (float) (pageHeight - nDiv - 61.0), (float) 178.618);
        document.add(p6);

        // row2 column3
        Paragraph p8 = new Paragraph("验 证 码:")
                .setFont(font)
                .setFontSize(11)
                .setFontColor(new DeviceRgb(0, 0, 0))
                .setFixedPosition((float) 556.67, (float) (pageHeight - nDiv - 61.0), (float) 40.172);
        document.add(p8);

        // row3 column1
        Paragraph p9 = new Paragraph("账  号:15877266778899")
                .setFont(font)
                .setFontSize(11)
                .setFontColor(new DeviceRgb(0, 0, 0))
                .setFixedPosition((float) 22.0, (float) (pageHeight - nDiv - 76.0), (float) 100.32001);
        document.add(p9);

        // row3 column2
        Paragraph p10 = new Paragraph("币种:RMB")
                .setFont(font)
                .setFontSize(11)
                .setFontColor(new DeviceRgb(0, 0, 0))
                .setFixedPosition((float) 289.33, (float) (pageHeight - nDiv - 76.0), (float) 47.542);
        document.add(p10);

        // row3 column3
        Paragraph p11 = new Paragraph("承前余额:67,673,054.52")
                .setFont(font)
                .setFontSize(11)
                .setFontColor(new DeviceRgb(0, 0, 0))
                .setFixedPosition((float) 556.67, (float) (pageHeight - nDiv - 76.0), (float) 105.29201);
        document.add(p11);

        // row4
        document.add(new Paragraph("序号").setFixedPosition((float) 22.0, (float) (pageHeight - nDiv - 100.0), (float) 22.0).setFontSize(11).setFont(font).setFontColor(new DeviceRgb(0, 0, 0)));
        document.add(new Paragraph("日期").setFixedPosition((float) 70.12, (float) (pageHeight - nDiv - 100.0), (float) 22.0).setFontSize(11).setFont(font).setFontColor(new DeviceRgb(0, 0, 0)));
        document.add(new Paragraph("借/贷方发生额").setFixedPosition((float) 134.28, (float) (pageHeight - nDiv - 100.0), (float) 69.673996).setFontSize(11).setFont(font).setFontColor(new DeviceRgb(0, 0, 0)));
        document.add(new Paragraph("对方户名").setFixedPosition((float) 326.76, (float) (pageHeight - nDiv - 100.0), (float) 44.0).setFontSize(11).setFont(font).setFontColor(new DeviceRgb(0, 0, 0)));
        document.add(new Paragraph("对方账户").setFixedPosition((float) 607.46, (float) (pageHeight - nDiv - 100.0), (float) 44.0).setFontSize(11).setFont(font).setFontColor(new DeviceRgb(0, 0, 0)));
        document.add(new Paragraph("摘要").setFixedPosition((float) 743.8, (float) (pageHeight - nDiv - 100.0), (float) 22.0).setFontSize(11).setFont(font).setFontColor(new DeviceRgb(0, 0, 0)));

        System.out.println("Fixed Content is done");
    }

    // 写入可变内容
    private void createVariableContent(Document document, float pageHeight, PdfFont font) {
        System.out.println("Creating Variable Content");
        System.out.println("Variable Content is done");
    }

    // 创建新的平安PDF
    public void createNewPinganPDF(String pinganModifiedPdfFilePath) {

        readExcel(excelFilePath);



        try {
            // 初始化内部变量
            // PdfFont font = PdfFontFactory.createFont("fonts/STSongStd-Light.ttf", "Identity-H");
            PdfFont font = PdfFontFactory.createFont("STSongStd-Light", "UniGB-UCS2-H");
            PdfWriter pdfWriter = new PdfWriter(pinganModifiedPdfFilePath);
            PdfDocument pdfDoc = new PdfDocument(pdfWriter);
            Document document = new Document(pdfDoc, PageSize.A4.rotate(), true);
            float pageHeight = PageSize.A4.rotate().getHeight();

            createPage(document, pageHeight, font);

            document.close();

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public void pdfToExcel(String pdfFilePath, String excelFilePath) {
        try {
            System.out.println("Reading Pingan Pdf is processing");

            PdfReader reader = new PdfReader(pdfFilePath);
            PdfDocument pdfDoc = new PdfDocument(reader);
            int nPages = pdfDoc.getNumberOfPages();

            System.out.println("Pages is " + nPages);

            SimpleTextExtractionStrategy strategy = new SimpleTextExtractionStrategy();

            PdfPage page = null;
            String text = null;

            for (int i = 1; i <= nPages; i++) {
                page = pdfDoc.getPage(i);
                text = PdfTextExtractor.getTextFromPage(page, strategy);
            }

            BufferedReader bufferReader = null;

            if (text != null) {
                bufferReader = new BufferedReader(new StringReader(text));
            }

            // 写入excel中
            // 创建 Excel 工作簿
            Workbook workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet("Sheet1");

            String line = null;
            if (bufferReader != null) {
                int rowNum = 0;
                while ((line = bufferReader.readLine()) != null) {
                    if (!line.isEmpty() && Character.isDigit(line.charAt(0))) {
                        // System.out.println(line); // 对于包含“交易资金支付结算服务”的行，不做输出，且上一行会少“交易资金支付结算服务”
                        String[] fields = line.split(" ");
                        Row row = sheet.createRow(rowNum);
                        for (int colIndex = 0; colIndex < fields.length; colIndex++) {
                            Cell cell = row.createCell(colIndex);
                            cell.setCellValue(fields[colIndex]);
                            if (fields.length == 6) {
                                cell = row.createCell(6);
                                cell.setCellValue("交易资金支付结算服务");
                            }
                        }
                        rowNum++;
                    }
                }
            }

            FileOutputStream fileOut = new FileOutputStream(excelFilePath);
            workbook.write(fileOut);

            System.out.println("Excel is done");

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
