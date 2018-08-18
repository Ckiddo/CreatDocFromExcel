import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

public class Ui extends JFrame {
    private JMenu menu = new JMenu("About");
    private JMenuItem[] menuItems = {
            new JMenuItem("About author"),
            new JMenuItem("About Software")
    };
    private JButton buttonWord = new JButton("添加Word模板");
    private JButton buttonExcel = new JButton("添加Excel");
    private JButton buttonCreate = new JButton("生成");
    private JTextField
            wordFileDir = new JTextField(20),
            excelFileDir = new JTextField(20);
    private String wordFileDirS,excelFileDirS;
    private JTextArea consoleLog = new JTextArea(10,30);
    public Ui(){
        setLayout(new FlowLayout());
        for(JMenuItem jMenuItem : menuItems){
            menu.add(jMenuItem);
        }
        JMenuBar mb = new JMenuBar();
        mb.add(menu);
        setJMenuBar(mb);
        buttonWord.addActionListener(new OpenWord());
        buttonExcel.addActionListener(new OpenExcel());
        buttonCreate.addActionListener(new CreateWordFiles());
        menuItems[0].addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                consoleLog.append("作者:陈康\n联系方式:815642475@qq.com\ngithub:https://github.com/Ckiddo\n");
            }
        });
        menuItems[1].addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                consoleLog.append("陈康 于2017/8/18日制作\n版本号1.0\n");
            }
        });
        add(buttonWord);
        add(wordFileDir);
        add(buttonExcel);
        add(excelFileDir);
        add(buttonCreate);
        add(consoleLog);
    }
    class OpenWord implements ActionListener{
        public void actionPerformed(ActionEvent e){
            JFileChooser c = new JFileChooser();
            int rVal = c.showOpenDialog(Ui.this);
            if(rVal == JFileChooser.APPROVE_OPTION){
                wordFileDirS = c.getSelectedFile().getAbsolutePath();
                wordFileDir.setText(wordFileDirS);

            }else{
                wordFileDir.setText("");
            }
        }
    }
    class OpenExcel implements ActionListener{
        public void actionPerformed(ActionEvent e){
            JFileChooser c = new JFileChooser();
            int rVal = c.showOpenDialog(Ui.this);
            if(rVal == JFileChooser.APPROVE_OPTION){
                excelFileDirS = c.getSelectedFile().getAbsolutePath();
                excelFileDir.setText(excelFileDirS);
            }
        }
    }
    class CreateWordFiles implements ActionListener{
        FileInputStream fileInputStreamWord,fileInputStreamExcel;
        HWPFDocument doc;
        HSSFWorkbook excel;
        ArrayList<String> needReplace = new ArrayList<String>();
        ArrayList< ArrayList<String> > myList = new ArrayList<ArrayList<String>>();
        public void openExcel()throws Exception{
            fileInputStreamExcel = new FileInputStream(new File(excelFileDirS));
            excel = new HSSFWorkbook(fileInputStreamExcel);
        }
        public void stractInformation(){
            Sheet sheet1 = excel.getSheetAt(0);
            int rows = sheet1.getPhysicalNumberOfRows();
            int cells = 0;
            if(rows >= 1){
                cells = sheet1.getRow(0).getPhysicalNumberOfCells();
            }
            consoleLog.append("Excel 表格行数:" + rows + " 列数:" + cells + "\n");

            for(int j = 0;j < cells;j++){
                needReplace.add(sheet1.getRow(0).getCell(j).getStringCellValue());

            }
            for(int i = 1;i < rows;i++){
                Row row = sheet1.getRow(i);

                ArrayList<String> aline = new ArrayList<String>();
                String cellValue;
                Cell cell;
                for(int j = 0;j < cells;j++){
                    cell = row.getCell(j);
                    cellValue = cell.getStringCellValue();
                    aline.add(cellValue);
                }
                myList.add(aline);
            }
        }
        public void fillPosition(){
            if(myList.size() <= 0){return;}
            for(ArrayList<String> al : myList){
                Range mainRan,headRan;
                try {
                    fileInputStreamWord = new FileInputStream(new File(wordFileDirS));
                    doc = new HWPFDocument(fileInputStreamWord);
                    mainRan = doc.getRange();
                    headRan = doc.getHeaderStoryRange();
                }catch (Exception exc){
                    consoleLog.append("打开word模板失败\n");
                    break;
                }
                try{
                    for(int i = 1;i < al.size();i++){

                        if(needReplace.get(i).contains("$")){
                            headRan.replaceText(needReplace.get(i),al.get(i));
                        }
                        if(needReplace.get(i).contains("%")) {
                            mainRan.replaceText(needReplace.get(i), al.get(i));
                        }
                    }
                    FileOutputStream fileOutputStream = new FileOutputStream(new File(al.get(0) + ".doc"));
                    doc.write(fileOutputStream);
                    fileOutputStream.close();
                    fileInputStreamWord.close();
                }catch (Exception exc){
                    consoleLog.append("替换或生成文件保存时出错\n");
                }
            }
        }
        public void closeFiles()throws Exception{
            fileInputStreamExcel.close();
        }
        public void actionPerformed(ActionEvent e){
            try{
                openExcel();
            }catch (Exception exc){
                consoleLog.append("open ExcelFile fail\n");
                return;
            }
            try {
                stractInformation();
                fillPosition();
            }catch (Exception exc){
                System.out.println(exc);
                consoleLog.append("StractInfomation fail\n");
            }
            try{
                closeFiles();
            }catch (Exception exc){
                consoleLog.append("close Files fail\n");
            }

        }
    }
    public static void main(String args[]){
        SwingConsole.run(new Ui(),400,350);
    }
}
