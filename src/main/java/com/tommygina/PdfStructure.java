package com.tommygina;

import com.itextpdf.kernel.font.PdfFont;
import com.itextpdf.kernel.geom.LineSegment;
import com.itextpdf.kernel.geom.Vector;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfPage;
import com.itextpdf.kernel.pdf.PdfReader;
import com.itextpdf.kernel.pdf.canvas.parser.EventType;
import com.itextpdf.kernel.pdf.canvas.parser.PdfCanvasProcessor;
import com.itextpdf.kernel.pdf.canvas.parser.data.IEventData;
import com.itextpdf.kernel.pdf.canvas.parser.data.ImageRenderInfo;
import com.itextpdf.kernel.pdf.canvas.parser.data.TextRenderInfo;
import com.itextpdf.kernel.pdf.canvas.parser.listener.IEventListener;
import com.itextpdf.kernel.pdf.xobject.PdfImageXObject;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Collections;
import java.util.Set;

class PdfStructure {

    // 读取平安PDF的文件架构，读取第一页即可
    public void getPdfStructureByFirstPage(String pdfFilePath) {
        try {
            System.out.println("Reading PDF is processing");
            PdfReader reader = new PdfReader(pdfFilePath);
            PdfDocument pdfDoc = new PdfDocument(reader);
            PdfPage page = pdfDoc.getPage(1);
            System.out.println("[Page 1 content is]");
            System.out.println("[Page 1 Height is]" + page.getPageSize().getHeight());
            System.out.println("[Page 1 Width is]" + page.getPageSize().getWidth());
            System.out.println("[Page 1 Left is]" + page.getPageSize().getLeft());
            System.out.println("[Page 1 Right is]" + page.getPageSize().getRight());
            // 使用CustomEventListener自定义策略进行PDF第一页的内容读取
            CustomEventListener strategy = new CustomEventListener();
            PdfCanvasProcessor processor = new PdfCanvasProcessor(strategy);
            processor.processPageContent(page);
            String text = strategy.getResultantText();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private class CustomEventListener implements IEventListener {

        private final StringBuilder textBuilder = new StringBuilder();

        @Override
        public void eventOccurred(IEventData data, EventType type) {
            // 文字处理
            if (type == EventType.RENDER_TEXT) {
                // 处理文本渲染事件
                TextRenderInfo renderInfo = (TextRenderInfo) data;
                String text = renderInfo.getText();
                textBuilder.append(text);

                // 输出文本内容
                System.out.println("Text: " + text);

                // 获取并输出文本的位置信息
                LineSegment baseline = renderInfo.getBaseline();
                Vector startPoint = baseline.getStartPoint(); // 基线起点
                Vector endPoint = baseline.getEndPoint();   // 基线终点


                System.out.println("Point: " + startPoint.get(Vector.I2) + " | " + startPoint.get(Vector.I1));

                // System.out.println("Baseline Start: (" + startPoint.get(Vector.I1) + ", " + startPoint.get(Vector.I2) + ")");
                // System.out.println("Baseline End: (" + endPoint.get(Vector.I1) + ", " + endPoint.get(Vector.I2) + ")");

                // 获取字体大小
                float fontSize = renderInfo.getFontSize();
                System.out.println("Font Size: " + fontSize);

                // 获取子弟
                PdfFont fontType = renderInfo.getFont();
                System.out.println("Font: " + fontType.getFontProgram());

                // 获取文本的未缩放宽度
                float unscaledWidth = renderInfo.getUnscaledWidth();
                System.out.println("Unscaled Width: " + unscaledWidth);

                // 获取上升线和下降线
//                LineSegment ascentLine = renderInfo.getAscentLine();
//                LineSegment descentLine = renderInfo.getDescentLine();
//
//                System.out.println("Ascent Line Start: (" + ascentLine.getStartPoint().get(Vector.I1) + ", " + ascentLine.getStartPoint().get(Vector.I2) + ")");
//                System.out.println("Ascent Line End: (" + ascentLine.getEndPoint().get(Vector.I1) + ", " + ascentLine.getEndPoint().get(Vector.I2) + ")");
//                System.out.println("Descent Line Start: (" + descentLine.getStartPoint().get(Vector.I1) + ", " + descentLine.getStartPoint().get(Vector.I2) + ")");
//                System.out.println("Descent Line End: (" + descentLine.getEndPoint().get(Vector.I1) + ", " + descentLine.getEndPoint().get(Vector.I2) + ")");

                System.out.println("-----------------------------");
            }

            if (type == EventType.RENDER_IMAGE) {
                // 处理图像渲染事件
                ImageRenderInfo renderInfo = (ImageRenderInfo) data;
                try {
                    // 获取PdfImageXObject（图像对象）
                    PdfImageXObject image = renderInfo.getImage();

                    // 提取图像字节数据
                    byte[] imageData = image.getImageBytes();
                    if (imageData != null) {
                        // 生成图像文件路径
                        String imagePath = "image_" + System.currentTimeMillis() + ".png";

                        // 将图像字节数据保存到文件
                        try (FileOutputStream imageOut = new FileOutputStream(new File(imagePath))) {
                            imageOut.write(imageData);
                        }

                        // 输出图像保存路径
                        System.out.println("Image saved at: " + new File(imagePath).getAbsolutePath());
                    }
                } catch (IOException e) {
                    System.err.println("Error saving image: " + e.getMessage());
                }
                System.out.println("-----------------------------");
            }
        }

        @Override
        public Set<EventType> getSupportedEvents() {
            // 返回支持的事件类型（只处理文本渲染事件）
            return Collections.singleton(EventType.RENDER_TEXT);
        }

        // 提供一个方法用于获取提取的文本
        public String getResultantText() {
            return textBuilder.toString();
        }

        // 提供一个方法用于清空已提取的文本
        public void clear() {
            textBuilder.setLength(0);
        }
    }

}
