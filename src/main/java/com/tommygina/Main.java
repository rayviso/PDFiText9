package com.tommygina;

import java.util.Scanner;

class Main {
    // private static final Logger logger = Logger.getLogger(Main.class.getName());

    private static final String pinganPdfFilePath = "pingan.pdf";
    private static final String pinganExcelFilePath = "pingan.xlsx";
    private static final String pinganModifiedPdfFilePath = "pinan_new.pdf";


    // <1>
    private static void getDataFromPDFtoExcel() {
        System.out.println("【Step 1】进行中");
        PinganPDF pa = new PinganPDF();
        pa.pdfToExcel(pinganPdfFilePath, pinganExcelFilePath);
        System.out.println("【Step 1】完成，请在程序所在目录查找“pingan.xlsx文件，并进行修改；修改完成后运行程序<2>选项”");
    }

    // <2>
    private static void createNewPDFfromExcel() {
        System.out.println("【Step 3】进行中");
        PinganPDF pa = new PinganPDF();
        pa.createNewPinganPDF(pinganModifiedPdfFilePath);
        System.out.println("【Step 3】新PDF文件生成了，文件名为“pingan_new.pdf”");
    }

    // <0>
    private static void showProgrammerInfo() {
        System.out.println("--------------------------------------------------------------------------");
        System.out.println("该程序使用说明：");
        System.out.println("【Step 1】");
        System.out.println("【操作】输入1后回车，执行第一步");
        System.out.println("【将PDF文件进行改名，并放到程序所在目录下】首先在当前目录下放平安银行月交易明细PDF文件，并把文件改名为\"pingan.pdf\"");
        System.out.println("【读取PDF并生成Excel文件】选择程序<1>选项进行pingan.pdf文件内容读取，并在当前目录生成一个Excel文件，名为pingan.xlsx");

        System.out.println("【Step 2】");
        System.out.println("【操作】手动修改pingan.xlsx文件中的数据，其中对照PDF文件，共有7例数据");
        System.out.println("【第1列】第一列数据为“序号”，删除或增加后，不用对序号进行修复，程序会自动处理不对的序号");
        System.out.println("【第2列】第二列数据为“日期”，按需进行修改即可");
        System.out.println("【第3列】第三列数据为“交易金额”，原数据中有+号和-号，在Excel处理中可以根据需要把该列数据设置为“数值”格式（Tips：保持2位小数点），处理完成后无需恢复数据格式");
        System.out.println("【第4列】第四列数据为“账号金额”，原数据中没有+号或-号，在Excel处理中可以根据需要把该列数据设置为“数值”格式（Tips：保持2位小数点），处理完成后无需恢复数据格式");
        System.out.println("【第5列】第五列数据为“姓名”，按需进行修改即可");
        System.out.println("【第6列】第六列数据为“客户账号”，按需进行修改即可");
        System.out.println("【第7列】第七列数据为“交易类型”，按需进行修改即可");

        System.out.println("【Step 3】");
        System.out.println("【操作】输入2后回车，执行第三步");
        System.out.println("【确认执行】程序自动根据当前目录下的“pingan.xlsx”进行操作，生成“pingan_new.pdf”文件");
        System.out.println("--------------------------------------------------------------------------");
    }


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
//        System.out.println("Step 4 is working");
//        PinganPDF pa = new PinganPDF();
//        pa.createNewPinganPDF(pinganModifiedPdfFilePath);
//        System.out.println("Step 4 is done");

        Scanner scanner = new Scanner(System.in);

        while (true) {
            System.out.println("输入9回车：打开程序使用说明 | 输入1回车：读取PDF文件，生成Excel文件 | 输入2回车：根据Excel生成PDF文件 | 输入0回车：退出当前程序");
            System.out.print(">>>>>>");

            String input = scanner.nextLine().trim();

            if (input.isEmpty()) {
                System.out.println("命令不能为空！");
                continue;
            }

            switch (input.toLowerCase()) {
                case "1":
                    getDataFromPDFtoExcel();
                    break;
                case "2":
                    createNewPDFfromExcel();
                    break;
                case "9":
                    showProgrammerInfo();
                    break;
                case "0":
                    scanner.close();
                    return;
                default:
                    System.out.println("无效命令，请重新输入！");
            }
        }
    }
}

