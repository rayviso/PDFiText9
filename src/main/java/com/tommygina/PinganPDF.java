// update36
package com.tommygina;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.itextpdf.io.image.ImageData;
import com.itextpdf.io.image.ImageDataFactory;
import com.itextpdf.kernel.colors.DeviceRgb;
import com.itextpdf.kernel.font.PdfFont;
import com.itextpdf.kernel.font.PdfFontFactory;
import com.itextpdf.kernel.font.PdfFontFactory.EmbeddingStrategy;
import com.itextpdf.kernel.geom.PageSize;
import com.itextpdf.kernel.pdf.*;
import com.itextpdf.kernel.pdf.canvas.PdfCanvas;
import com.itextpdf.kernel.pdf.canvas.parser.PdfTextExtractor;
import com.itextpdf.kernel.pdf.canvas.parser.listener.SimpleTextExtractionStrategy;
import com.itextpdf.layout.Document;
import com.itextpdf.layout.element.AreaBreak;
import com.itextpdf.layout.element.Image;
import com.itextpdf.layout.element.Paragraph;
import com.itextpdf.layout.properties.AreaBreakType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.DecimalFormat;
import java.text.DecimalFormatSymbols;
import java.util.Locale;
import java.util.Objects;

public class PinganPDF {

    private static final String pinganConfigJsonFilePath = "pingan.json";
    private static PdfWriter pdfWriter;
    private static WriterProperties writerProperties;
    private static String date;
    private static String balanceBroughtDown;
    // "打印时间:2025-01-01"
    private static String printedTime;
    private String configDate;
    private String configBalanceBroughtDown;
    private String configPrintDate;

    public String getConfigDate() {
        return configDate;
    }
    private static final PdfVersion pdfVersion15 = PdfVersion.PDF_1_5;
    private static final float pageA4Height = PageSize.A4.rotate().getHeight(); // 595.0 2479px
    private static final float pageA4Width = PageSize.A4.rotate().getWidth();  // 842.0 3508px
    private static PdfFont fontChinese;
    private static PdfDocument pdfDoc;

    public void setConfigDate(String configDate) {
        this.configDate = configDate;
    }

    public String getConfigBalanceBroughtDown() {
        return configBalanceBroughtDown;
    }
    private static Document document;
    private static String pageInfo;

    static {
        try {
            fontChinese = PdfFontFactory.createFont("STSongStd-Light", "UniGB-UCS2-H", EmbeddingStrategy.PREFER_NOT_EMBEDDED, false);

            // fontChinese = PdfFontFactory.createTtcFont("");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static final float nDiv = (float) 3.9;
    private static final String excelFilePath = "pingan.xlsx";

    // row1 3/4：少一个页面信息，动态输入即可
    private static final String monthlyStatement = "客户存款月结单";
    private static final String statementNumber = "结单号：2412310251999000006599";

    public void setConfigBalanceBroughtDown(String configBalanceBroughtDown) {
        this.configBalanceBroughtDown = configBalanceBroughtDown;
    }

    // row2 3/3
    private static final String clientBank = "客户行:平安银行杭州分行营业部";
    private static final String accountName = "户名:上海华通铂银交易市场有限公司";
    private static final String verificationCode = "验 证 码:";

    // row3 3/3
    private static final String accountNumber = "账  号:15877266778899";
    private static final String currency = "币种:RMB";

    public String getConfigPrintDate() {
        return configPrintDate;
    }

    // head 7/7
    private static final String headSerialNumber = "序号";
    private static final String headDealDate = "日期";
    private static final String headTransactionAmount = "借/贷方发生额";
    private static final String headBalance = "余额";
    private static final String headReciprocalAccountName = "对方户名";
    private static final String headReciprocalAccount = "对方账户";
    private static final String headSummary = "摘要";

    // row4 5/5
    private static final String printedTimes = "已打印次数:1";

    public void setConfigPrintDate(String configPrintDate) {
        this.configPrintDate = configPrintDate;
    }
    private static final String printedType = "打印方式:系统PDF生成";
    private static final String deviceNumber = "设备编号:0000";
    private static final String tellerNumber = "柜员号:";

    // inforow
    private static final String info1 = "温馨提示：根据国家财政部颁发的《内部会计控制规范-货币资金（试行）》第十九条的规定，单位应与开户银行定期进行对账，此月结单每月初印发，请收到后及时";
    private static final String info2 = "核对财务；核对不符的，应在印发当月15日前与开户行联系，查明原因，及时处理；逾期没有联系的，视为财务核对相符，请妥善保管月结单，并在您的地址发生变";
    private static final String info3 = "换时，请及时书面通知我行。";

    // png
    private static final String pinganLogo = "pinganLogo.png";
    private static final String pinganCachet = "pinganCachet.png";

    // 创建单个pdf页面
    private void createPage(int startRow, int endRow, Sheet sheet) {
        createFixedContent();
        // 读取当前批次的数据
        for (int i = startRow; i < endRow; i++) {
            Row row = sheet.getRow(i);
            int m = i % 25;

            if (row != null) {
                createVariableContent(row, m, i);
            }
        }

//        document.add(new AreaBreak());
        document.add(new AreaBreak(AreaBreakType.NEXT_PAGE));

    }

    // 读取
    private void readExcelAndCreatePages(String excelFilePath) {
        try {
            InputStream is = new FileInputStream(excelFilePath);
//            FileMagic fm = FileMagic.valueOf(is);
//            System.out.println(fm);
            org.apache.poi.ss.usermodel.Workbook workbook = org.apache.poi.ss.usermodel.WorkbookFactory.create(is);
            Sheet sheet = workbook.getSheetAt(0);
            int nRows = sheet.getPhysicalNumberOfRows();
            int nPDFPages = (int) Math.ceil((double) nRows / 25);
            int batchSize = 25;
            System.out.println("【共计创建PDF页面数为" + nPDFPages + "页】");

            int currentPage = 1;

            for (int startRow = 0; startRow < nRows; startRow += batchSize) {
//            for (int startRow = 0; startRow < 5; startRow += batchSize) {
                int endRow = Math.min(startRow + batchSize, nRows);
//                System.out.println("Reading rows from " + (startRow + 1) + " to " + endRow);
//                System.out.println("Current Page is " + currentPage);
                pageInfo = "第" + currentPage + "页  共" + nPDFPages + "页";
                createPage(startRow, endRow, sheet);
                currentPage++;
            }

            // 关闭资源
            workbook.close();
            is.close();

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    // 从Excel中获取对应的值，输出为String
    private String getCellString(Cell cell, boolean bSymbol) {
        // 根据单元格的类型读取数据
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                double value = cell.getNumericCellValue();
                String pattern;
                if (bSymbol) {
                    pattern = "+#,##0.00;-#,##0.00";
                } else {
                    pattern = "#,##0.00";
                }


                // 设置区域符号（确保逗号为千分位，点号作小数点）
                DecimalFormatSymbols symbols = new DecimalFormatSymbols(Locale.US);
                symbols.setGroupingSeparator(',');
                symbols.setDecimalSeparator('.');

                DecimalFormat formatter = new DecimalFormat(pattern, symbols);
                String formattedValue = formatter.format(value);
                return formattedValue;
            // return Double.toString(cell.getNumericCellValue());
            case BOOLEAN:
                return Boolean.toString(cell.getBooleanCellValue());
            default:
                return "UNKNOWN";
        }
    }

    // 从json文件中获取配置信息
    private void getConfig(String configFilePath) {

        ObjectMapper mapper = new ObjectMapper();
        try {
            File configFile = new File(configFilePath);
            PinganPDF pinganPDF = mapper.readValue(configFile, PinganPDF.class);
            date = pinganPDF.getConfigDate();
            balanceBroughtDown = pinganPDF.getConfigBalanceBroughtDown();
            printedTime = pinganPDF.getConfigPrintDate();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    // 固定格式文件
    private void createFixedContent() {
//        System.out.println("Creating Fixed Content");
        getConfig(pinganConfigJsonFilePath);
        // image
        try {
            ImageData logoImageData = ImageDataFactory.create(pinganLogo);
            ImageData cachetImageData = ImageDataFactory.create(pinganCachet);

            Image imageLogo = new Image(logoImageData);
            Image imageCachet = new Image(cachetImageData);

            document.add(imageLogo.setFixedPosition((float) 20, (float) (pageA4Height - nDiv - 36.1)).scaleAbsolute((float) 160, (float) 30));
//            document.add(imageCachet.setFixedPosition((float) 714.8, (float) (pageA4Height - nDiv - 501.1)).scaleAbsolute((float) 72.95, (float) 33.90));
            document.add(imageCachet.setFixedPosition((float) 715.0, (float) (pageA4Height - nDiv - 501.1)).scaleAbsolute((float) 72.95, (float) 33.90));

        } catch (IOException e) {
            e.printStackTrace();
        }

        // 横线
        try {
            PdfPage pdfPage = pdfDoc.getLastPage(); // 不要增加新页面，在最后一页进行操作，使用pdfCanvas前需要获取当前页面。
//            pdfPage = pdfDoc.addNewPage();
            PdfCanvas pdfCanvas = new PdfCanvas(pdfPage);


            // 设置横线的起始点和结束点
//            float x1 = (float) 19.68173199241533;
            float x1 = (float) 19.99173199241533;
            float x2 = (float) 821.8323212442695;

            float y0 = (float) -0.8;

            float y1 = (float) (511.9650529246574 + y0);
            float y2 = (float) (485.8427717639152 + y0);
//            float y3 = (float) (104.169166886686 + y0);
//            float y4 = (float) (79.68701245709622 + y0);
            float y3 = (float) (104.369166886686 + y0);
            float y4 = (float) (79.88701245709622 + y0);


            // 绘制线条
            // 第一根线
            // 设置线条的粗细和颜色
            pdfCanvas.setLineWidth((float) 2);  // 线条粗细
            pdfCanvas.setStrokeColor(new DeviceRgb(150, 150, 150));
            pdfCanvas.moveTo(x1, y1);
            pdfCanvas.lineTo(x2, y1);
            pdfCanvas.stroke();

            // 第二根线
            pdfCanvas.setLineWidth((float) 2);  // 线条粗细

            pdfCanvas.moveTo(x1, y2);
            pdfCanvas.lineTo(x2, y2);
            pdfCanvas.stroke();

            pdfCanvas.setLineWidth((float) 0.5);  // 线条粗细
            pdfCanvas.setStrokeColor(new DeviceRgb(100, 100, 100));

            pdfCanvas.moveTo(x1, y3);
            pdfCanvas.lineTo(x2, y3);
            pdfCanvas.stroke();

            pdfCanvas.moveTo(x1, y4);
            pdfCanvas.lineTo(x2, y4);
            pdfCanvas.stroke();

        } catch (Exception e) {
            e.printStackTrace();
        }

        // row1
        document.add(new Paragraph(monthlyStatement).setFixedPosition((float) 222.5, (float) (pageA4Height - nDiv - 41.0), (float) 77).setFontSize(11).setFont(fontChinese).setFontColor(new DeviceRgb(0, 0, 0)));
        document.add(new Paragraph(statementNumber).setFixedPosition((float) 342.8, (float) (pageA4Height - nDiv - 41.0), (float) 180).setFontSize(11).setFont(fontChinese).setFontColor(new DeviceRgb(0, 0, 0)));
        document.add(new Paragraph(date).setFixedPosition((float) 583.4, (float) (pageA4Height - nDiv - 41.0), (float) 100).setFontSize(11).setFont(fontChinese).setFontColor(new DeviceRgb(0, 0, 0)));
        document.add(new Paragraph(pageInfo).setFixedPosition((float) 663.6, (float) (pageA4Height - nDiv - 41.0), (float) 100).setFontSize(11).setFont(fontChinese).setFontColor(new DeviceRgb(0, 0, 0)));


        // row2
        document.add(new Paragraph(clientBank).setFixedPosition((float) 22.0, (float) (pageA4Height - nDiv - 61.0), (float) 156.618).setFontSize(11).setFont(fontChinese).setFontColor(new DeviceRgb(0, 0, 0)));
        document.add(new Paragraph(accountName).setFixedPosition((float) 289.33, (float) (pageA4Height - nDiv - 61.0), (float) 178.618).setFontSize(11).setFont(fontChinese).setFontColor(new DeviceRgb(0, 0, 0)));
        document.add(new Paragraph(verificationCode).setFixedPosition((float) 556.67, (float) (pageA4Height - nDiv - 61.0), (float) 40.172).setFontSize(11).setFont(fontChinese).setFontColor(new DeviceRgb(0, 0, 0)));

        // row3
        document.add(new Paragraph(accountNumber).setFixedPosition((float) 22.0, (float) (pageA4Height - nDiv - 76.0), (float) 100.32001).setFontSize(11).setFont(fontChinese).setFontColor(new DeviceRgb(0, 0, 0)));
        document.add(new Paragraph(currency).setFixedPosition((float) 289.33, (float) (pageA4Height - nDiv - 76.0), (float) 47.542).setFontSize(11).setFont(fontChinese).setFontColor(new DeviceRgb(0, 0, 0)));
        document.add(new Paragraph(balanceBroughtDown).setFixedPosition((float) 556.67, (float) (pageA4Height - nDiv - 76.0), (float) 105.29201).setFontSize(11).setFont(fontChinese).setFontColor(new DeviceRgb(0, 0, 0)));

        // row4
        document.add(new Paragraph(headSerialNumber).setFixedPosition((float) 22.0, (float) (pageA4Height - nDiv - 100.0), (float) 22.0).setFontSize(11).setFont(fontChinese).setFontColor(new DeviceRgb(0, 0, 0)));
        document.add(new Paragraph(headDealDate).setFixedPosition((float) 70.12, (float) (pageA4Height - nDiv - 100.0), (float) 22.0).setFontSize(11).setFont(fontChinese).setFontColor(new DeviceRgb(0, 0, 0)));
        document.add(new Paragraph(headTransactionAmount).setFixedPosition((float) 134.28, (float) (pageA4Height - nDiv - 100.0), (float) 69.673996).setFontSize(11).setFont(fontChinese).setFontColor(new DeviceRgb(0, 0, 0)));
        document.add(new Paragraph(headBalance).setFixedPosition((float) 230.52, (float) (pageA4Height - nDiv - 100.0), (float) 22).setFontSize(11).setFont(fontChinese).setFontColor(new DeviceRgb(0, 0, 0)));
        document.add(new Paragraph(headReciprocalAccountName).setFixedPosition((float) 326.76, (float) (pageA4Height - nDiv - 100.0), (float) 44.0).setFontSize(11).setFont(fontChinese).setFontColor(new DeviceRgb(0, 0, 0)));
        document.add(new Paragraph(headReciprocalAccount).setFixedPosition((float) 607.46, (float) (pageA4Height - nDiv - 100.0), (float) 44.0).setFontSize(11).setFont(fontChinese).setFontColor(new DeviceRgb(0, 0, 0)));
        document.add(new Paragraph(headSummary).setFixedPosition((float) 743.8, (float) (pageA4Height - nDiv - 100.0), (float) 22.0).setFontSize(11).setFont(fontChinese).setFontColor(new DeviceRgb(0, 0, 0)));

        // row5
        document.add(new Paragraph(printedTimes).setFixedPosition((float) 22.0, (float) (pageA4Height - nDiv - 507.0), (float) 64).setFontSize(11).setFont(fontChinese).setFontColor(new DeviceRgb(0, 0, 0)));
        document.add(new Paragraph(printedTime).setFixedPosition((float) 142.3, (float) (pageA4Height - nDiv - 507.0), (float) 300).setFontSize(11).setFont(fontChinese).setFontColor(new DeviceRgb(0, 0, 0)));
        document.add(new Paragraph(printedType).setFixedPosition((float) 382.9, (float) (pageA4Height - nDiv - 507.0), (float) 120).setFontSize(11).setFont(fontChinese).setFontColor(new DeviceRgb(0, 0, 0)));
        document.add(new Paragraph(deviceNumber).setFixedPosition((float) 543.3, (float) (pageA4Height - nDiv - 507.0), (float) 72).setFontSize(11).setFont(fontChinese).setFontColor(new DeviceRgb(0, 0, 0)));
        document.add(new Paragraph(tellerNumber).setFixedPosition((float) 703.7, (float) (pageA4Height - nDiv - 507.0), (float) 40).setFontSize(11).setFont(fontChinese).setFontColor(new DeviceRgb(0, 0, 0)));

        // info row
//        document.add(new Paragraph(info1).setFixedPosition((float) 22.0, (float) (pageA4Height - nDiv - 528.7), (float) 800).setFontSize(11).setFont(fontChinese).setFontColor(new DeviceRgb(0, 0, 0)));
        document.add(new Paragraph(info1).setFixedPosition((float) 22.0, (float) (pageA4Height - nDiv - 529.1), (float) 800).setFontSize(11).setFont(fontChinese).setFontColor(new DeviceRgb(0, 0, 0)));
        document.add(new Paragraph(info2).setFixedPosition((float) 22.0, (float) (pageA4Height - nDiv - 540.0), (float) 800).setFontSize(11).setFont(fontChinese).setFontColor(new DeviceRgb(0, 0, 0)));
        document.add(new Paragraph(info3).setFixedPosition((float) 22.0, (float) (pageA4Height - nDiv - 551.0), (float) 800).setFontSize(11).setFont(fontChinese).setFontColor(new DeviceRgb(0, 0, 0)));

//        System.out.println("Fixed Content is done");
    }

    private void createVariableContent(Row row, int m, int i) {
//        for (int i = 0; i < 7; i++) {
//            System.out.println(getCellString(row.getCell(i)));
//        }
        document.add(new Paragraph(Integer.toString(i + 1)).setFixedPosition((float) 22.0, (float) (pageA4Height - nDiv - 124.0 - 15 * m), (float) 30.0).setFontSize(11).setFont(fontChinese).setFontColor(new DeviceRgb(0, 0, 0)));
        document.add(new Paragraph(getCellString(row.getCell(1), false)).setFixedPosition((float) 70.12, (float) (pageA4Height - nDiv - 124.0 - 15 * m), (float) 42).setFontSize(11).setFont(fontChinese).setFontColor(new DeviceRgb(0, 0, 0)));
        document.add(new Paragraph(getCellString(row.getCell(2), true)).setFixedPosition((float) 134.28, (float) (pageA4Height - nDiv - 124.0 - 15 * m), (float) 70.0).setFontSize(11).setFont(fontChinese).setFontColor(new DeviceRgb(0, 0, 0)));
        document.add(new Paragraph(getCellString(row.getCell(3), false)).setFixedPosition((float) 230.52, (float) (pageA4Height - nDiv - 124.0 - 15 * m), (float) 64.0).setFontSize(11).setFont(fontChinese).setFontColor(new DeviceRgb(0, 0, 0)));
        document.add(new Paragraph(getCellString(row.getCell(4), false)).setFixedPosition((float) 326.76, (float) (pageA4Height - nDiv - 124.0 - 15 * m), (float) 260.0).setFontSize(11).setFont(fontChinese).setFontColor(new DeviceRgb(0, 0, 0)));
        document.add(new Paragraph(getCellString(row.getCell(5), false)).setFixedPosition((float) 607.46, (float) (pageA4Height - nDiv - 124.0 - 15 * m), (float) 100.0).setFontSize(11).setFont(fontChinese).setFontColor(new DeviceRgb(0, 0, 0)));
        String s = getCellString(row.getCell(6), true);
        if (Objects.equals(s, "交易资金支付结算服务")) {
            document.add(new Paragraph(s).setFixedPosition((float) 743.8, (float) (pageA4Height - nDiv - 124.0 - 15 * m + 4.2), (float) 100.0).setFontSize(8).setFont(fontChinese).setFontColor(new DeviceRgb(0, 0, 0)));
        } else {
            document.add(new Paragraph(s).setFixedPosition((float) 743.8, (float) (pageA4Height - nDiv - 124.0 - 15 * m), (float) 100.0).setFontSize(11).setFont(fontChinese).setFontColor(new DeviceRgb(0, 0, 0)));
        }
    }

    // #1 创建pdf文件
    public void createNewPinganPDF(String pinganModifiedPdfFilePath) {
        try {
//            writerProperties = new WriterProperties().setPdfVersion(pdfVersion15).useSmartMode().setCompressionLevel(CompressionConstants.BEST_COMPRESSION);
            writerProperties = new WriterProperties().setPdfVersion(pdfVersion15).useSmartMode().setCompressionLevel(CompressionConstants.BEST_COMPRESSION);

            pdfWriter = new PdfWriter(pinganModifiedPdfFilePath, writerProperties);

            pdfDoc = new PdfDocument(pdfWriter);
//            System.out.println(pdfDoc.getPdfVersion());
            document = new Document(pdfDoc, PageSize.A4.rotate());

            readExcelAndCreatePages(excelFilePath);

            // 将多出来的一页进行删除
            pdfDoc.removePage(pdfDoc.getNumberOfPages());
            pdfDoc.close();
            pdfWriter.close();
            document.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public void pdfToExcel(String pdfFilePath, String excelFilePath) {
        try {
            System.out.println("【从平安PDF中读取相关数据，进行中...】");

            PdfReader reader = new PdfReader(pdfFilePath);
            PdfDocument pdfDoc = new PdfDocument(reader);
            int nPages = pdfDoc.getNumberOfPages();

            System.out.println("【读取到" + nPages + "页】");

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

//            System.out.println("Excel is done");

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}


