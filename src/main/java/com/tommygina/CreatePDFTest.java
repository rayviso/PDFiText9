// update36
package com.tommygina;

import com.itextpdf.kernel.colors.DeviceRgb;
import com.itextpdf.kernel.font.PdfFont;
import com.itextpdf.kernel.font.PdfFontFactory;
import com.itextpdf.kernel.geom.PageSize;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.layout.Document;
import com.itextpdf.layout.element.AreaBreak;
import com.itextpdf.layout.element.Paragraph;
import com.itextpdf.layout.properties.AreaBreakType;
import com.itextpdf.layout.properties.TextAlignment;
import com.itextpdf.layout.properties.VerticalAlignment;

import java.io.IOException;

class CreatePDFTest {
    // 创建PDF
    public static void main() {
        try {
            System.out.println("Creating PDF is processing");

            PdfFont font = PdfFontFactory.createFont("STSongStd-Light", "UniGB-UCS2-H");
            PdfWriter writer = new PdfWriter("demo.pdf");
            PdfDocument pdfDoc = new PdfDocument(writer);
            Document document = new Document(pdfDoc, PageSize.A4.rotate());

            // 页面1
            pdfDoc.addNewPage();
            document.add(new Paragraph("Page1"));

            document.add(new AreaBreak(AreaBreakType.NEXT_PAGE));


            // 页面2
            pdfDoc.addNewPage();
            document.add(new Paragraph("Page2"));
            // 设置文本的字体、大小、颜色、背景颜色、对齐方式、垂直对其方式
            Paragraph paragraph2 = new Paragraph("感谢您阅读蚂蚁小哥的博客！")
                    .setFont(font)
                    .setFontSize(18)
                    .setFontColor(new DeviceRgb(255, 0, 0))
                    .setBackgroundColor(new DeviceRgb(187, 255, 255))
                    .setTextAlignment(TextAlignment.CENTER)
                    .setVerticalAlignment(VerticalAlignment.BOTTOM);
            document.add(paragraph2);


            // 页面3
            pdfDoc.addNewPage();
            document.add(new Paragraph("Page3"));

            // 关闭
            document.close();
            System.out.println("Creating PDF is done");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
