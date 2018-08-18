
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Range;


import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.List;


public class CreateRZ03 {
    static HWPFDocument doc;
    static FileInputStream inputStream;
    static FileOutputStream outputStream;
    static void openFile()throws  Exception{
        String fileName = "/Users/chenkang/Desktop/" +
                            "RZ03 认证任务交接分配单V4.1.doc";
        inputStream = new FileInputStream(new File(fileName));
        doc = new HWPFDocument(inputStream);
        String outFileName = "/Users/chenkang/Desktop/" +
                                "生成的DOC.doc";
        outputStream = new FileOutputStream(outFileName);
    }
    static void fillPosition()throws  Exception{
        Range mainRan =  doc.getRange();
        Range headRan = doc.getHeaderStoryRange();
        String bianhao = "17（05）Y0566" + "RZ03";
        String shenbaojibie = "新三级申报";
        String qiyemingc = "深圳市科陆物联信息技术有限公司";
        headRan.replaceText("$jilubianhao",bianhao);
        mainRan.replaceText("$shenbaojibie",shenbaojibie);
        mainRan.replaceText("$qiyemingcheng",qiyemingc);
        System.out.println("replaceTextFinished");
        doc.write(outputStream);
        outputStream.close();
    }
    static void run()throws  Exception{
        openFile();
        fillPosition();
        inputStream.close();
    }

    public static void main(String args[])throws  Exception{
        run();
    }
}
