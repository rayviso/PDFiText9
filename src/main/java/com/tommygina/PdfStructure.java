package com.tommygina;

import com.itextpdf.kernel.font.PdfFont;
import com.itextpdf.kernel.geom.LineSegment;
import com.itextpdf.kernel.geom.Matrix;
import com.itextpdf.kernel.geom.Vector;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfName;
import com.itextpdf.kernel.pdf.PdfPage;
import com.itextpdf.kernel.pdf.PdfReader;
import com.itextpdf.kernel.pdf.canvas.parser.EventType;
import com.itextpdf.kernel.pdf.canvas.parser.data.IEventData;
import com.itextpdf.kernel.pdf.canvas.parser.data.ImageRenderInfo;
import com.itextpdf.kernel.pdf.canvas.parser.data.TextRenderInfo;
import com.itextpdf.kernel.pdf.canvas.parser.listener.IEventListener;
import com.itextpdf.kernel.pdf.xobject.PdfImageXObject;

import javax.imageio.ImageIO;
import javax.imageio.ImageReader;
import javax.imageio.metadata.IIOMetadata;
import javax.imageio.stream.ImageInputStream;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Collections;
import java.util.Iterator;
import java.util.Set;

class PdfStructure {

    // 读取平安PDF的文件架构，读取第一页即可
    public void getPdfStructureByFirstPage(String pdfFilePath, int nPage) {
        try {
            // 读取指定页面nPage
            System.out.println("Reading PDF is processing");
            PdfReader reader = new PdfReader(pdfFilePath);
            PdfDocument pdfDoc = new PdfDocument(reader);
            PdfPage page = pdfDoc.getPage(nPage);


            // 获取表格的相关属性

//            // 获取文字的相关属性
//            System.out.println("[Page 1 content is]");
//            System.out.println("[Page 1 Height is]" + page.getPageSize().getHeight());
//            System.out.println("[Page 1 Width is]" + page.getPageSize().getWidth());
//            System.out.println("[Page 1 Left is]" + page.getPageSize().getLeft());
//            System.out.println("[Page 1 Right is]" + page.getPageSize().getRight());
//            // 使用CustomEventListener自定义策略进行PDF第一页的内容读取
//            CustomEventListener cel = new CustomEventListener();
//            PdfCanvasProcessor processor = new PdfCanvasProcessor(cel);
//            processor.processPageContent(page);
//            String text = cel.getResultantText();


            // 获取图像的相关属性

//            ImageExtractorListener iel = new ImageExtractorListener("./", 1);
//            PdfCanvasProcessor processor2 = new PdfCanvasProcessor(iel);
//            processor2.processPageContent(page);




        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static class CustomEventListener implements IEventListener {

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
//            Set<EventType> events = new HashSet<>();
//            events.add(EventType.RENDER_TEXT);
//            events.add(EventType.RENDER_IMAGE);
//            return Set.of(EventType.RENDER_TEXT, EventType.RENDER_IMAGE);

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

    private static class ImageExtractorListener implements IEventListener {
        private final String outputDir;
        private final int pageNumber;

        public ImageExtractorListener(String outputDir, int pageNumber) {
            this.outputDir = outputDir;
            this.pageNumber = pageNumber;
        }

        @Override
        public void eventOccurred(IEventData data, EventType eventType) {
            if (eventType == EventType.RENDER_IMAGE && data instanceof ImageRenderInfo) {
                try {
//                  Image img = (Image)data;
                    ImageRenderInfo imageInfo = (ImageRenderInfo) data;
                    PdfImageXObject imageXObject = imageInfo.getImage();
                    System.out.println(imageXObject.getWidth());
                    System.out.println(imageXObject.getHeight());

                    if (imageXObject != null) {
                        int colorSpace = imageXObject.getPdfObject().getAsNumber(PdfName.BitsPerComponent).intValue();

                        // 生成唯一的图片文件路径

                        String imagePath = outputDir + "image_page" + pageNumber + "_" + System.currentTimeMillis() + ".png";

                        // 过滤掉 2-bit 颜色深度的图片
                        if (colorSpace == 2) {
                            System.out.println("发现 2-bit 图片，转换为 8-bit 格式...");
                            System.out.println("Image Type is " + imageXObject.identifyImageType());
                            System.out.println(imageXObject.getWidth());
                            System.out.println(imageXObject.getHeight());


                            try (FileOutputStream fos = new FileOutputStream(new File(imagePath))) {
                                // 先不做输出
                                byte[] imageBytes2 = imageXObject.getImageBytes(false);
                                fos.write(imageBytes2);
                            }

                            // BufferedImage convertedImage = convert2BitTo8Bit(imageXObject);
                            // ImageIO.write(convertedImage, "png", new File(imagePath));
                            return;
                        }


                        Matrix ctm = imageInfo.getImageCtm();
                        float x = ctm.get(Matrix.I31);  // 获取 X 坐标
                        float y = ctm.get(Matrix.I32);  // 获取 Y 坐标

                        // 保存图片

                        try (FileOutputStream fos = new FileOutputStream(new File(imagePath))) {
                            byte[] imageBytes = imageXObject.getImageBytes();
                            fos.write(imageBytes);
                        }

                        // 输出图片的位置信息
                        System.out.println("图片已保存：" + imagePath);
                        System.out.println("位置信息 -> 页码: " + pageNumber + ", X: " + x + ", Y: " + y);


                        // 调试信息
                        // 0-2
                        System.out.println(ctm.get(Matrix.I11));
                        System.out.println(ctm.get(Matrix.I12));
                        System.out.println(ctm.get(Matrix.I13));

                        // 3-5
                        System.out.println(ctm.get(Matrix.I21));
                        System.out.println(ctm.get(Matrix.I22));
                        System.out.println(ctm.get(Matrix.I23));

                        // 6-8
                        System.out.println(ctm.get(Matrix.I31));
                        System.out.println(ctm.get(Matrix.I32));
                        System.out.println(ctm.get(Matrix.I33));
                    }
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }

        @Override
        public Set<EventType> getSupportedEvents() {
            return Collections.singleton(EventType.RENDER_IMAGE);
        }


        // 2-bit 转换为 8-bit
        public BufferedImage convert2BitTo8Bit(PdfImageXObject imageXObject) throws IOException {
            byte[] imageBytes = imageXObject.getImageBytes(false);
            BufferedImage buffImage = read2BitImage(imageBytes);

            int width = (int) imageXObject.getWidth();
            int height = (int) imageXObject.getHeight();

            BufferedImage bufferedImage = new BufferedImage(width, height, BufferedImage.TYPE_BYTE_GRAY);
            int index = 0;

            for (int y = 0; y < height; y++) {
                for (int x = 0; x < width; x += 4) { // 2-bit 一次存储 4 像素
                    int pixelData = (imageBytes[index++] & 0xFF);
                    for (int i = 0; i < 4 && (x + i) < width; i++) {
                        int pixelValue = ((pixelData >> (6 - (i * 2))) & 0x03) * 85; // 2-bit -> 8-bit
                        bufferedImage.setRGB(x + i, y, new Color(pixelValue, pixelValue, pixelValue).getRGB());
                    }
                }
            }

            return bufferedImage;
        }


        public BufferedImage read2BitImage(byte[] imageBytes) {
            try (ByteArrayInputStream bis = new ByteArrayInputStream(imageBytes);
                 ImageInputStream iis = ImageIO.createImageInputStream(bis)) {


                Iterator<ImageReader> readers = ImageIO.getImageReaders(iis);
                if (!readers.hasNext()) {
                    throw new RuntimeException("无可用TIFF阅读器");
                }

                ImageReader reader = readers.next();
                reader.setInput(iis);

                // 获取图像元数据（检查位深）
                IIOMetadata metadata = reader.getImageMetadata(0);
                String bitsPerSample = metadata.getAsTree("javax_imageio_tiff_image_metadata")
                        .getAttributes()
                        .getNamedItem("BitsPerSample")
                        .getNodeValue();
                System.out.println("位深: " + bitsPerSample); // 应输出"2"

                BufferedImage image = reader.read(0);

                // 处理图像


                return image;
            } catch (IOException e) {
                e.printStackTrace();
            }
            return null;
        }
    }

}
