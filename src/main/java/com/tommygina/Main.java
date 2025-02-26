package com.tommygina;

class Main {
    // private static final Logger logger = Logger.getLogger(Main.class.getName());

    private static final String pinganPdfFilePath = "pingan.pdf";
    private static final String pinganExcelFilePath = "pingan.xlsx";
    private static final String pinganModifiedPdfFilePath = "pinan_new.pdf";

    public static void main(String[] args) {

        // 第一步：从平安PDF中获取相关页面架构信息
//        System.out.println("Step 1 is working");
//        PdfStructure pps = new PdfStructure();
//        pps.getPdfStructureByFirstPage(pinganPdfFilePath, 1);
//        System.out.println("Step 1 is done");

        // 第二步：从平安PDF中获取相关交易数据，生成文件为Excel文件
//        System.out.println("Step 2 is working");
//        PinganPDF pa = new PinganPDF();
//        pa.pdfToExcel(pinganPdfFilePath, pinganExcelFilePath);
//        System.out.println("Step 2 is done");

        // 第三步：人工修改Excel文件，并保存；保证
        // 第三列格式化为自定义格式 [>0]+#,##0.00;[<0]-#,##0.00;0.00
        // 第四列格式化为数字格式，带有千位,号；两位小数点后数字

        // 第四步：根据修改的Excel生成新的平安PDF文件
        System.out.println("Step 4 is working");
        PinganPDF pa = new PinganPDF();
        pa.createNewPinganPDF(pinganModifiedPdfFilePath);
        System.out.println("Step 4 is done");
    }
}