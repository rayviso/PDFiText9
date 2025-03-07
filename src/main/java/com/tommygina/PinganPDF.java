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
    // static final变量
    private static final PdfVersion pdfVersion15 = PdfVersion.PDF_1_5;
    private static final float pageA4Height = PageSize.A4.rotate().getHeight(); // 595.0 2479px 用来进行竖向定位
    // private static final float pageA4Width = PageSize.A4.rotate().getWidth();  // 842.0 3508px
    private static final float nDiv = (float) 3.9; // 定位偏移值

    // row1 3/4：少一个页面信息，动态输入即可
    private static final String monthlyStatement = "客户存款月结单";
    private static final String statementNumber = "结单号：2412310251999000006599";
    // 日期动态传递
    // 页面动态传递

    // row2 3/3
    private static final String clientBank = "客户行:平安银行杭州分行营业部";
    private static final String accountName = "户名:上海华通铂银交易市场有限公司";
    private static final String verificationCode = "验 证 码:";

    // row3 3/3
    private static final String accountNumber = "账  号:15877266778899";
    private static final String currency = "币种:RMB";
    // 承前余额 动态传递

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
    // 打印时间动态传递
    private static final String printedType = "打印方式:系统PDF生成";
    private static final String deviceNumber = "设备编号:0000";
    private static final String tellerNumber = "柜员号:";

    // inforow
    private static final String info1 = "温馨提示：根据国家财政部颁发的《内部会计控制规范-货币资金（试行）》第十九条的规定，单位应与开户银行定期进行对账，此月结单每月初印发，请收到后及时";
    private static final String info2 = "核对财务；核对不符的，应在印发当月15日前与开户行联系，查明原因，及时处理；逾期没有联系的，视为财务核对相符，请妥善保管月结单，并在您的地址发生变";
    private static final String info3 = "换时，请及时书面通知我行。";

    // static变量
    private static PdfFont fontChinese;
//    static {
//        try {
//            fontChinese = PdfFontFactory.createFont("STSongStd-Light", "UniGB-UCS2-H", EmbeddingStrategy.PREFER_NOT_EMBEDDED, false);
//            // fontChinese = PdfFontFactory.createTtcFont("");
//        } catch (IOException e) {
//            e.printStackTrace();
//        }
//    }

    // TODO：需要调整的变量
    private static PdfWriter pdfWriter;
    private static WriterProperties writerProperties;
    private static PdfDocument pdfDoc;
    private static Document document;

    // configDate
    private String configDate;
    // configBalanceBroughtDown
    private String configBalanceBroughtDown;
    // configPrintDate
    private String configPrintDate;
    // pageInfo
    private String pageInfo;

    public String getConfigDate() {
        return configDate;
    }

    public void setConfigDate(String configDate) {
        this.configDate = configDate;
    }

    public String getConfigBalanceBroughtDown() {
        return configBalanceBroughtDown;
    }

    public void setConfigBalanceBroughtDown(String configBalanceBroughtDown) {
        this.configBalanceBroughtDown = configBalanceBroughtDown;
    }

    public String getConfigPrintDate() {
        return configPrintDate;
    }

    public void setConfigPrintDate(String configPrintDate) {
        this.configPrintDate = configPrintDate;
    }

    public String getPageInfo() {
        return pageInfo;
    }

    public void setPageInfo(String pageInfo) {
        this.pageInfo = pageInfo;
    }

    // <1> 分两个函数执行，从PDF中读取一页，并创建json文件，读取全部页面，创建excel文件
    public void getPDFInfo(String pinganPdfFilePath, String pinganConfigJsonFilePath, String pinganExcelFilePath) {

        if (createJson(pinganPdfFilePath, pinganConfigJsonFilePath)) {
            if (createExcel(pinganPdfFilePath, pinganExcelFilePath)) {
                System.out.println("【执行成功】生成 pingan.xlsx 和 pingan.json");
            } else {
                System.out.println("【部分执行成功】pingan.json");
            }
        }
    }

    // <1.1> 从PDF中获取对应的日期，承前余额，打印日期等数据
    private boolean createJson(String pinganPdfFilePath, String pinganConfig) {

        try {
            System.out.println("【从平安PDF中读取“日期等变化信息”并生成Json文件，进行中...】");
            PdfReader reader = new PdfReader(pinganPdfFilePath);
            PdfDocument pdfDoc = new PdfDocument(reader);
            SimpleTextExtractionStrategy strategy = new SimpleTextExtractionStrategy();
            PdfPage page = pdfDoc.getPage(1);
            String text = PdfTextExtractor.getTextFromPage(page, strategy);
            BufferedReader bufferReader = null;

            if (text != null) {
                bufferReader = new BufferedReader(new StringReader(text));
            }

            File configFile = new File(pinganConfig);
            PinganPDF pinganPDF = new PinganPDF();
            ObjectMapper mapper = new ObjectMapper();

            String line = null;
            if (bufferReader != null) {
                int rowNum = 0;
                while ((line = bufferReader.readLine()) != null) {
                    if (line.startsWith("客户存款月结单")) {
                        String[] fields = line.split(" ");
                        System.out.println("【日期】：" + fields[2]);
                        pinganPDF.setConfigDate(fields[2]);
                    }

                    if (line.startsWith("账")) {
                        String[] fields = line.split(" ");
                        System.out.println("【承前余额】：" + fields[4]);
                        pinganPDF.setConfigBalanceBroughtDown(fields[4]);
                    }

                    if (line.startsWith("已打印次数")) {
                        String[] fields = line.split(" ");
                        System.out.println("【打印时间】：" + fields[1]);
                        pinganPDF.setConfigPrintDate(fields[1]);
                    }
                }
            }

            mapper.writeValue(configFile, pinganPDF);
            System.out.println("【配置文件创建成功】配置文件为当前目录，文件名为 pingan.json");
            return true;
        } catch (IOException e) {
            e.printStackTrace();
            return false;
        }
    }

    // <1.2> 创建Excel文件
    private boolean createExcel(String pinganPdfFilePath, String pinganExcelFilePath) {
        try {
            System.out.println("【从平安PDF中读取“交易数据”并生成Excel文件，进行中...】");

            PdfReader reader = new PdfReader(pinganPdfFilePath);
            PdfDocument pdfDoc = new PdfDocument(reader);
            int nPages = pdfDoc.getNumberOfPages();

            System.out.println("【本次读取数据" + nPages + "页】");

            // 使用SimpleTextExtractionStrategy进行文字读取即可，相关坐标信息读取使用的PdfStructure类进行获取
            SimpleTextExtractionStrategy strategy = new SimpleTextExtractionStrategy();

            PdfPage page = null;
            String text = null;

            // 一次性把所有页面的数据都读取到text文件中
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

            FileOutputStream fileOut = new FileOutputStream(pinganExcelFilePath);
            workbook.write(fileOut);
            fileOut.close();
            workbook.close();
            return true;
        } catch (Exception e) {
            e.printStackTrace();
            return false;
        }
    }

    // <2> 创建pdf文件
    public boolean createNewPDF(String pinganModifiedPdfFilePath, String pinganLogo, String pinganCachet, String pinganConfigJsonFilePath, String pinganExcelFilePath) {
        try {
            if (!getConfigFromJson(pinganConfigJsonFilePath)) {
                return false;
            }

            // WriterProperties writerProperties = new WriterProperties().setPdfVersion(pdfVersion15).useSmartMode().setCompressionLevel(CompressionConstants.BEST_COMPRESSION);
            // PdfWriter pdfWriter = new PdfWriter(pinganModifiedPdfFilePath, writerProperties);
            // PdfDocument pdfDoc = new PdfDocument(pdfWriter);
            // Document document = new Document(pdfDoc, PageSize.A4.rotate());

            fontChinese = PdfFontFactory.createFont("STSongStd-Light", "UniGB-UCS2-H", EmbeddingStrategy.PREFER_NOT_EMBEDDED, true);
            writerProperties = new WriterProperties().setPdfVersion(pdfVersion15).useSmartMode().setCompressionLevel(CompressionConstants.BEST_COMPRESSION);
            pdfWriter = new PdfWriter(pinganModifiedPdfFilePath, writerProperties);

            pdfDoc = new PdfDocument(pdfWriter);
            document = new Document(pdfDoc, PageSize.A4.rotate());

            // 创建页面s
            createPages(pinganExcelFilePath, pinganLogo, pinganCachet, pdfWriter, pdfDoc, document);
            // 将多出来的一页进行删除
            pdfDoc.removePage(pdfDoc.getNumberOfPages());
            pdfDoc.close();

            document.flush();
            document.close();

            pdfWriter.flush();
            pdfWriter.close();


            return true;
        } catch (IOException e) {
            e.printStackTrace();
            return false;
        }
    }

    // <2.1> 从json文件中获取配置信息
    private boolean getConfigFromJson(String pinganConfigJsonFilePath) {
        try {
            File pinganConfigJsonFile = new File(pinganConfigJsonFilePath);

            if (!pinganConfigJsonFile.exists()) {
                System.out.println("【运行错误】缺少 pingan.json 文件，运行 <1> 选项生成 pingan.json ");
                return false;
            } else {
                ObjectMapper mapper = new ObjectMapper();
                PinganPDF pinganPDF = mapper.readValue(pinganConfigJsonFile, PinganPDF.class);
                configDate = pinganPDF.getConfigDate();
                configBalanceBroughtDown = pinganPDF.getConfigBalanceBroughtDown();
                configPrintDate = pinganPDF.getConfigPrintDate();
                return true;
            }

        } catch (IOException e) {
            e.printStackTrace();
            return false;
        }
    }

    // <2.2> 创建PDF的不同页面
    private void createPages(String pinganExcelFilePath, String pinganLogo, String pinganCachet, PdfWriter pdfWriter, PdfDocument pdfDoc, Document document) {
        try {
            InputStream is = new FileInputStream(pinganExcelFilePath);
            // FileMagic fm = FileMagic.valueOf(is);
            // System.out.println(fm);
            // workbook用显示声明
            org.apache.poi.ss.usermodel.Workbook workbook = org.apache.poi.ss.usermodel.WorkbookFactory.create(is);
            Sheet sheet = workbook.getSheetAt(0);
            int nRows = sheet.getPhysicalNumberOfRows();
            int nPDFPages = (int) Math.ceil((double) nRows / 25);
            int batchSize = 25;
            System.out.println("【共计创建PDF页面数为" + nPDFPages + "页】");

            int currentPage = 1;

            for (int startRow = 0; startRow < nRows; startRow += batchSize) {
                int endRow = Math.min(startRow + batchSize, nRows);
                pageInfo = "第" + currentPage + "页  共" + nPDFPages + "页";
                createPage(startRow, endRow, sheet, pinganLogo, pinganCachet, pdfDoc, document);
                currentPage++;
            }

            workbook.close();
            is.close();

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    // <2.3> 创建PDF的单个pdf页面
    private void createPage(int startRow, int endRow, Sheet sheet, String pinganLogo, String pinganCachet, PdfDocument pdfDoc, Document document) {
        createPageFixedContent(pinganLogo, pinganCachet, pdfDoc, document);
        // 读取当前批次的数据
        for (int i = startRow; i < endRow; i++) {
            Row row = sheet.getRow(i);
            int m = i % 25;

            if (row != null) {
                createPageVariableContent(row, m, i, document);
            }
        }

        document.add(new AreaBreak(AreaBreakType.NEXT_PAGE));
    }

    // <2.3.1> 固定格式文件
    private void createPageFixedContent(String pinganLogo, String pinganCachet, PdfDocument pdfDoc, Document document) {
        // System.out.println("Creating Fixed Content");

        // image
        try {
            ImageData logoImageData = ImageDataFactory.create(pinganLogo);
            ImageData cachetImageData = ImageDataFactory.create(pinganCachet);

            Image imageLogo = new Image(logoImageData);
            Image imageCachet = new Image(cachetImageData);

            document.add(imageLogo.setFixedPosition((float) 20, (float) (pageA4Height - nDiv - 36.1)).scaleAbsolute((float) 160, (float) 30));
            // document.add(imageCachet.setFixedPosition((float) 714.8, (float) (pageA4Height - nDiv - 501.1)).scaleAbsolute((float) 72.95, (float) 33.90));
            document.add(imageCachet.setFixedPosition((float) 715.0, (float) (pageA4Height - nDiv - 501.1)).scaleAbsolute((float) 72.95, (float) 33.90));

        } catch (IOException e) {
            e.printStackTrace();
        }

        // 横线
        try {
            PdfPage pdfPage = pdfDoc.getLastPage(); // 不要增加新页面，在最后一页进行操作，使用pdfCanvas前需要获取当前页面。
            // pdfPage = pdfDoc.addNewPage();
            PdfCanvas pdfCanvas = new PdfCanvas(pdfPage);

            // 设置横线的起始点和结束点
            // float x1 = (float) 19.68173199241533;
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
        // configDate从json文件中进行读取
        document.add(new Paragraph(configDate).setFixedPosition((float) 583.4, (float) (pageA4Height - nDiv - 41.0), (float) 100).setFontSize(11).setFont(fontChinese).setFontColor(new DeviceRgb(0, 0, 0)));
        document.add(new Paragraph(pageInfo).setFixedPosition((float) 663.6, (float) (pageA4Height - nDiv - 41.0), (float) 100).setFontSize(11).setFont(fontChinese).setFontColor(new DeviceRgb(0, 0, 0)));

        // row2
        document.add(new Paragraph(clientBank).setFixedPosition((float) 22.0, (float) (pageA4Height - nDiv - 61.0), (float) 156.618).setFontSize(11).setFont(fontChinese).setFontColor(new DeviceRgb(0, 0, 0)));
        document.add(new Paragraph(accountName).setFixedPosition((float) 289.33, (float) (pageA4Height - nDiv - 61.0), (float) 178.618).setFontSize(11).setFont(fontChinese).setFontColor(new DeviceRgb(0, 0, 0)));
        document.add(new Paragraph(verificationCode).setFixedPosition((float) 556.67, (float) (pageA4Height - nDiv - 61.0), (float) 40.172).setFontSize(11).setFont(fontChinese).setFontColor(new DeviceRgb(0, 0, 0)));

        // row3
        document.add(new Paragraph(accountNumber).setFixedPosition((float) 22.0, (float) (pageA4Height - nDiv - 76.0), (float) 100.32001).setFontSize(11).setFont(fontChinese).setFontColor(new DeviceRgb(0, 0, 0)));
        document.add(new Paragraph(currency).setFixedPosition((float) 289.33, (float) (pageA4Height - nDiv - 76.0), (float) 47.542).setFontSize(11).setFont(fontChinese).setFontColor(new DeviceRgb(0, 0, 0)));
        // configBalanceBroughtDown从json文件中进行读取
        document.add(new Paragraph(configBalanceBroughtDown).setFixedPosition((float) 556.67, (float) (pageA4Height - nDiv - 76.0), (float) 105.29201).setFontSize(11).setFont(fontChinese).setFontColor(new DeviceRgb(0, 0, 0)));

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
        // configPrintDate从json文件中进行读取
        document.add(new Paragraph(configPrintDate).setFixedPosition((float) 142.3, (float) (pageA4Height - nDiv - 507.0), (float) 300).setFontSize(11).setFont(fontChinese).setFontColor(new DeviceRgb(0, 0, 0)));
        document.add(new Paragraph(printedType).setFixedPosition((float) 382.9, (float) (pageA4Height - nDiv - 507.0), (float) 120).setFontSize(11).setFont(fontChinese).setFontColor(new DeviceRgb(0, 0, 0)));
        document.add(new Paragraph(deviceNumber).setFixedPosition((float) 543.3, (float) (pageA4Height - nDiv - 507.0), (float) 72).setFontSize(11).setFont(fontChinese).setFontColor(new DeviceRgb(0, 0, 0)));
        document.add(new Paragraph(tellerNumber).setFixedPosition((float) 703.7, (float) (pageA4Height - nDiv - 507.0), (float) 40).setFontSize(11).setFont(fontChinese).setFontColor(new DeviceRgb(0, 0, 0)));

        // info row
        document.add(new Paragraph(info1).setFixedPosition((float) 22.0, (float) (pageA4Height - nDiv - 529.1), (float) 800).setFontSize(11).setFont(fontChinese).setFontColor(new DeviceRgb(0, 0, 0)));
        document.add(new Paragraph(info2).setFixedPosition((float) 22.0, (float) (pageA4Height - nDiv - 540.0), (float) 800).setFontSize(11).setFont(fontChinese).setFontColor(new DeviceRgb(0, 0, 0)));
        document.add(new Paragraph(info3).setFixedPosition((float) 22.0, (float) (pageA4Height - nDiv - 551.0), (float) 800).setFontSize(11).setFont(fontChinese).setFontColor(new DeviceRgb(0, 0, 0)));
    }

    // <2.3.2> 非固定内容：每行交易数据
    private void createPageVariableContent(Row row, int m, int i, Document document) {
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

    // {2} 从Excel中获取对应的值，输出为String
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
                return formatter.format(value);
            // return Double.toString(cell.getNumericCellValue());
            case BOOLEAN:
                return Boolean.toString(cell.getBooleanCellValue());
            default:
                return "UNKNOWN";
        }
    }
}


