// update36-2
package com.tommygina;

import java.io.File;
import java.util.Scanner;

class Main {
    // Main 变量
    // 相关程序使用到的同目录下的文件
    private static final String pinganPdfFilePath = "pingan.pdf";
    private static final String pinganLogo = "pinganLogo.png";
    private static final String pinganCachet = "pinganCachet.png";

    private static final String pinganExcelFilePath = "pingan.xlsx";
    private static final String pinganModifiedPdfFilePath = "pingan_new.pdf";
    private static final String pinganConfigJsonFilePath = "pingan.json";

    private static PinganPDF pp = new PinganPDF();

    // 运行前确定是否存在必要的文件
    private static boolean checkFiles(String pinganPdfFilePath, String pinganLogo, String pinganCachet) {
        File f1 = new File(pinganPdfFilePath);
        File f2 = new File(pinganLogo);
        File f3 = new File(pinganCachet);
        if (f1.exists() && f2.exists() && f3.exists()) {
            return true;
        } else {
            return false;
        }
    }

    // <1>
    private static void getDataFromPDFtoExcel() {
        System.out.println("【Step 1】进行中");
        pp.getPDFInfo(pinganPdfFilePath, pinganConfigJsonFilePath, pinganExcelFilePath);
        System.out.println("【Step 1】完成");
        System.out.println("【后续操作】pingan.json 文件用来存储读取到的日期等信心，默认不需要进行修改；请在程序所在目录查找 pingan.xlsx 文件，并进行修改；修改完成后运行程序 <2> 选项");
    }

    // <2>
    private static void createNewPDFfromExcel() {
        System.out.println("【Step 3】进行中");
        if (pp.createNewPDF(pinganModifiedPdfFilePath, pinganLogo, pinganCachet, pinganConfigJsonFilePath, pinganExcelFilePath)) {
            System.out.println("【运行成功】新PDF文件生成了，文件名为 pingan_new.pdf");
        } else {
            System.out.println("【运行失败】缺少配置信息，检查提示信息并重新运行");
        }
    }

    // <0>
    private static void showProgrammerInfo() {
        System.out.println("--------------------------------------------------------------------------");
        System.out.println("该程序使用说明：");
        System.out.println("【Step 1】");
        System.out.println("【操作】输入1后回车，执行第一步");
        System.out.println("【将PDF文件进行改名，并放到程序所在目录下】首先在当前目录下放平安银行月交易明细PDF文件，并把文件改名为 pingan.pdf ");
        System.out.println("【读取PDF生成Excel文件】选择程序 <1> 选项进行 pingan.pdf 文件内容读取在当前目录生成一个Excel文件，名为 pingan.xlsx ");
        System.out.println("【读取PDF生成Json文件】选择程序 <1> 选项进行 pingan.pdf 文件内容读取在当前目录生成一个Json文件，名为 pingan.json ，该文件为日期、承前余额、打印日期三个可变动内容");
        System.out.println();
        System.out.println("【Step 2】");
        System.out.println("【操作】手动修改 pingan.xlsx 文件中的数据，其中对照PDF文件，共有7例数据");
        System.out.println("【第1列】第一列数据为“序号”，删除或增加后，不用对序号进行修复，程序会自动处理不对的序号");
        System.out.println("【第2列】第二列数据为“日期”，按需进行修改即可");
        System.out.println("【第3列】第三列数据为“交易金额”，原数据中有+号和-号，在Excel处理中可以根据需要把该列数据设置为“数值”格式（Tips：保持2位小数点），处理完成后无需恢复数据格式");
        System.out.println("【第4列】第四列数据为“账号金额”，原数据中没有+号或-号，在Excel处理中可以根据需要把该列数据设置为“数值”格式（Tips：保持2位小数点），处理完成后无需恢复数据格式");
        System.out.println("【第5列】第五列数据为“姓名”，按需进行修改即可");
        System.out.println("【第6列】第六列数据为“客户账号”，按需进行修改即可");
        System.out.println("【第7列】第七列数据为“交易类型”，按需进行修改即可");
        System.out.println();
        System.out.println("【Step 3】");
        System.out.println("【操作】输入2后回车，执行第三步");
        System.out.println("【确认执行】程序自动根据当前目录下的 pingan.xlsx 和 pingan.json 进行操作，生成 pingan_new.pdf 文件");
//        System.out.println("--------------------------------------------------------------------------");
    }


    public static void main(String[] args) {

        Scanner scanner = new Scanner(System.in);

        while (true) {
            System.out.println("--------------------------------------------------------------------------");
            System.out.println("【运行程序须知】   \t\t在选择执行功能前，确保当期程序所在目录有 pingan.pdf 、 pinganLogo.png 、 pinganCachet.png 三个文件");
            System.out.println("【打开程序说明】   \t\t输入数字 9 并回车 ");
            System.out.println("【生成Excel文件】  \t\t输入数字 1 并回车 ");
            System.out.println("【生成PDF文件】    \t\t输入数字 2 并回车 ");
            System.out.println("【退出程序】       \t\t输入数字 0 并回车 ");
            System.out.println("--------------------------------------------------------------------------");
            System.out.print(">>> ");

            String input = scanner.nextLine().trim();

            if (input.isEmpty()) {
                System.out.println("命令不能为空！");
                continue;
            }

            switch (input.toLowerCase()) {
                case "1":
                    if (checkFiles(pinganPdfFilePath, pinganLogo, pinganCachet)) {
                        getDataFromPDFtoExcel();
                    } else {
                        System.out.println("【确认相关文件是否存在】");
                    }
                    break;
                case "2":
                    if (checkFiles(pinganPdfFilePath, pinganLogo, pinganCachet)) {
                        createNewPDFfromExcel();
                    } else {
                        System.out.println("【确认相关文件是否存在】");
                    }
                    break;
                case "9":
                    showProgrammerInfo();
                    break;
                case "0":
                    scanner.close();
                    return;
                default:
                    System.out.println("无效命令，请重新输入");
            }
        }
    }
}

