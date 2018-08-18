import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

public class CreateArchive {

    static String companyNames[] = {
            "1 深圳市诚业通信技术有限公司,17（04）Y0439",
            "2 深圳市辰赛智能科技有限公司,17（05）Y0538",
            "3 深圳市广和诚信息科技有限公司,17（05）Y0539",
            "4 深圳市小豆科技有限公司,17（05）Y0590",
            "5 深圳市时空数码科技有限公司,17（07）Y0931",
            "6 深圳市统先科技股份有限公司,17（04）Y0438",
            "7 深圳市信德远致科技有限公司,17（04）Y0436",
            "8 深圳市互盟科技股份有限公司,17（05）Y0583",
            "9 深圳市大山智能技术有限公司,17（05）Y0561",
            "10 深圳市中航深亚科技有限公司,17（05）Y0573",
            "11 深圳市欣视景科技股份有限公司,17（05）Y0559",
            "12 深圳市卓讯信息技术有限公司,17（05）Y0540",
            "13 深圳市捷宇通信技术有限公司,17（02）Y0120",
            "14 深圳市威富多媒体有限公司,17（05）Y0585",
            "15 深圳市爱华勘测工程有限公司,16（08）Y0943",
            "16 深圳艾派网络科技股份有限公司,17（04）Y0452",
            "17 深圳航天信息有限公司,17（05）Y0587",
            "18 深圳京柏医疗科技股份有限公司,17（05）Y0584",
            "19 深圳市时讯迪科技有限公司,17（05）Y0565",
            "20 深圳市科陆物联信息技术有限公司,17（05）Y0566",
            "21 深圳市神飞致远技术有限公司,17（05）Y0567",
            "22 深圳市海能达技术服务有限公司,18（01）Y0061",
            "23 深圳市谊润科技有限公司,17（04）Y0435",
            "24 深圳市纬度网络能源有限公司,17（04）Y0433",
            "25 深圳迅停科技发展有限公司,17（06）Y0701",
            "26 深圳市精诚天路科技有限公司,17（04）Y0451",
            "27 深圳市创势互联科技有限公司,17（05）Y0574",
            "28 深圳市东运科技有限公司,17（05）Y0568",
            "29 深圳市阔宏电子科技有限公司,17（05）Y0586",
            "30 中教（深圳）系统集成有限公司,17（05）Y0575",
            "31 深圳市众合汇通工程技术有限公司,17（05）Y0554",
            "32 深圳市青云酒店信息技术有限公司,17（05）Y0556",
            "33 深圳吉英荣科技有限公司,17（05）Y0553",
            "34 深圳市前海信息通信发展有限公司,17（07）Y0942",
            "35 深圳市聚迅科技有限公司,17（05）Y0560",
            "36 深圳市中浦信实业有限公司,17（05）Y0544",
            "37 深圳市数智国兴集成服务有限公司,17（05）Y0552",
            "38 深圳市锐铁计算机技术有限公司,17（02）Y0126",
            "39 深圳市前海网维信息有限公司,18（01）Y0061",
            "40 深圳市广汇源水利勘测设计有限公司,17（03）Y0264",
            "41 深圳市育新文化传播有限公司,17（06）Y0701",
            "42 深圳汉红科技有限公司,17（05）Y0578",
            "43 深圳市振兴光通信股份有限公司,17（06）Y0717",
            "44 深圳市深信信息技术有限公司,17（06）Y0722",
            "45 深圳国辰智能系统有限公司,17（05）Y0541",
            "46 广州市蓝翔信息技术有限公司,17（05）Y0557",
            "47 广州东融电子工程技术有限公司,17（05）Y0579",
            "48 广州迅联通讯科技股份有限公司,17（05）Y0580",
            "49 广州同构信息科技有限公司,17（06）Y0702",
            "50 广州速率信息技术有限公司,17（05）Y0570"
    };

    static String muBanLujing[] = {
            "RZ01 认证申报材料交接单V4.1.doc", // 0
            "RZ02 合同评审表V4.0.doc", // 1
            "RZ03 认证任务交接分配单V4.1.doc", // 2
            "RZ04 集成申报材料问题报告单V4.3.doc", // 3
            "RZ05 集成现场评审计划V4.3-一级.doc", // 4
            "RZ05 集成现场评审计划V4.3-三级.doc", // 5
            "RZ05 集成现场评审计划V4.3-二级.doc", // 6
            "RZ05 集成现场评审计划V4.3-四级.doc", // 7
            "RZ06 集成现场评审确认单V4.3.doc", // 8
            "RZ07 客户信息反馈表v4.0.doc",  // 9
            "RZ08 评审报告评审表V4.0.doc",  //10
            "RZ09 集成企业资质现场评审记录表V4.3-三级—修改 .doc", // 11
            "RZ09 集成企业资质现场评审记录表V4.3-一级—修改.doc",  // 12
            "RZ09 集成企业资质现场评审记录表V4.3-二级—修改.doc",  // 13
            "RZ09 集成企业资质现场评审记录表V4.3-四级—修改.doc", // 14
            "RZ10 记录归档登记表V4.0.doc", // 15
            "RZ11 认证项目质量监督记录表V4.0.doc", //16
            "RZ62 现场评审问题报告V4.0(模版).doc", // 17
            "RZ80 首末次会议记录表V4.0.doc" // 18
    };
    static String shenBaoTime[] = {
            "2017/3/20",
            "2017/3/7",
            "2017/3/1",
            "2017/3/17",
            "2017/3/16",
            "2017/3/21",
            "2017/3/9",
            "2017/3/15",
            "2017/3/17",
            "2017/3/9",
            "2017/2/28",
            "2017/3/17",
            "2017/3/1",
            "2017/3/17",
            "2017/3/25",
            "2017/3/20",
            "2017/3/24",
            "2017/3/2",
            "2017/3/15",
            "2017/3/27",
            "2017/3/18",
            "2017/3/24",
            "2017/3/20",
            "2017/3/30",
            "2017/3/15",
            "2017/3/15",
            "2017/3/15",
            "2017/3/7",
            "2017/3/3",
            "2017/3/15",
            "2017/3/20",
            "2017/3/15",
            "2017/3/15",
            "2017/3/15",
            "2017/3/10",
            "2017/2/17",
            "2017/3/17",
            "2017/3/20",
            "2017/3/1",
            "",
            "2017/3/15",
            "",
            "2017/3/15",
            "2017/3/25",
            "2017/3/30",
            "2017/3/20",
            "2017/3/20",
            "2016/3/13",
            "2017/3/10",
            "2017/3/1"
    };

    static String tijiaoRen[] = {
            "李文忠",
            "李洪柱",
            "郑小爽",
            "刘英",
            "龚春梅",
            "范文风",
            "刘丽娟",
            "张瑶",
            "徐明",
            "邓小飞",
            "郑粉嫦",
            "张丽敏",
            "王雪娟",
            "王振",
            "章巍",
            "张乐",
            "罗思",
            "安丽",
            "李春林",
            "刘芳华",
            "李慈珍",
            "邓爱玲",
            "毛健",
            "潘孟春",
            "梁芳虹",
            "张肖娴",
            "卜伟东",
            "王伟东",
            "曹二招",
            "密忆秋",
            "魏永丰",
            "李晓琳",
            "王仲灵",
            "张福财",
            "梁嘉玲",
            "杨珽",
            "田昌盛",
            "华功",
            "王洁",
            " ",
            "吴丽琼",
            " ",
            "薛洁",
            "王惟",
            "李莉",
            "龙冰",
            "张辉",
            "王小燕",
            "张彦利",
            "李小玲"
    };

    static String dianhua[] = {
            "18928297868",
            "18078158818",
            "15914156533",
            "18938046979",
            "13410104460",
            "18576760467",
            "13823322461",
            "18923899958",
            "13923428062",
            "13923894304",
            "13714701605",
            "13928424127",
            "18927407305",
            "13537538721",
            "13510274176",
            "13538106590",
            "15818757198",
            "18822893802",
            "13424253937",
            "13699808151",
            "13509600377",
            "13826138254",
            "13528717632",
            "15013710829",
            "13714868923",
            "13418506089",
            "18617167947",
            "13538118225",
            "13724347230",
            "13773055539",
            "18666229095",
            "13823332460",
            "13600449961",
            "18665931116",
            "13530162691",
            "18926106999",
            "13427928527",
            "13570814270",
            "15338781666",
            " ",
            "13689595965",
            " ",
            "13825299300",
            "13825299300",
            "18718842552",
            "15920509723",
            "13902278013",
            "18078823383",
            "13763337395",
            "18988991093"
    };
    static String jieshouRen[] = {
            "蔡荟荃",
            "蔡荟荃",
            "蔡荟荃",
            "蔡荟荃",
            "蔡荟荃",
            "蔡荟荃",
            "蔡荟荃",
            "蔡荟荃",
            "蔡荟荃",
            "蔡荟荃",
            "蔡荟荃",
            "蔡荟荃",
            "蔡荟荃",
            "蔡荟荃",
            "蔡荟荃",
            "蔡荟荃",
            "蔡荟荃",
            "蔡荟荃",
            "蔡荟荃",
            "蔡荟荃",
            "蔡荟荃",
            "蔡荟荃",
            "蔡荟荃",
            "蔡荟荃",
            "蔡荟荃",
            "蔡荟荃",
            "蔡荟荃",
            "蔡荟荃",
            "蔡荟荃",
            "蔡荟荃",
            "蔡荟荃",
            "蔡荟荃",
            "蔡荟荃",
            "蔡荟荃",
            "蔡荟荃",
            "蔡荟荃",
            "蔡荟荃",
            "蔡荟荃",
            "蔡荟荃",
            "蔡荟荃",
            "蔡荟荃",
            "蔡荟荃",
            "蔡荟荃",
            "蔡荟荃",
            "蔡荟荃",
            "蔡荟荃",
            "徐勇",
            "徐勇",
            "徐勇",
            "徐勇"
    };

    static String lujing = "/Users/chenkang/Desktop/集成V4.3-2017年2月后使用/";
    static String outputPath = "/Users/chenkang/Desktop/生成的归档/";

    static void createRZ01(){
        for(int i = 0;i < 50;i++){
            String inFileName = muBanLujing[0];
            try{
                FileInputStream inputStream = new FileInputStream(new File(lujing + inFileName));
                HWPFDocument inputDoc = new HWPFDocument(inputStream);
                Range headRan = inputDoc.getHeaderStoryRange();
                Range mainRan = inputDoc.getRange();
                String bianhao = companyNames[i].split(",")[1] + "RZ01";
                String compName = companyNames[i].split(",")[0].split(" ")[1];
                String sbjb = "三级新申报";if(22<=i && i <= 44){sbjb = "四级新申报";}
                String tjsj = shenBaoTime[i];
                String tjr = tijiaoRen[i];
                String jsr = jieshouRen[i];
                headRan.replaceText("$jlbh",bianhao);
                mainRan.replaceText("$sbdw",compName);
                mainRan.replaceText("$sbjb",sbjb);
                mainRan.replaceText("$tjsj",tjsj);
                mainRan.replaceText("$tjr",tjr);
                mainRan.replaceText("$jsr",jsr);
                FileOutputStream outputStream = new FileOutputStream(new File(outputPath + (i+1) + compName + "/" + inFileName));
                inputDoc.write(outputStream);
                outputStream.close();
                inputStream.close();
            }catch (Exception e){
                System.out.println(e);
                System.out.println(i);
            }
        }
    }





    static void createRZ03(){
        for(int i = 0;i < 50;i++){
            String inFileName = muBanLujing[2];
            try{
                FileInputStream inputStream = new FileInputStream(new File(lujing + inFileName));
                HWPFDocument inputDoc = new HWPFDocument(inputStream);
                Range headRan = inputDoc.getHeaderStoryRange();
                Range mainRan = inputDoc.getRange();
                String bianhao = companyNames[i].split(",")[1] + "RZ03";
                String sbjb = "三级新申报";if(22<=i && i <= 44){sbjb = "四级新申报";}
                String compName = companyNames[i].split(",")[0].split(" ")[1];
                headRan.replaceText("$jlbh",bianhao);
                mainRan.replaceText("$sbjb",sbjb);
                mainRan.replaceText("$qymc",compName);
                FileOutputStream outputStream = new FileOutputStream(new File(outputPath + (i+1) + compName + "/" + inFileName));
                inputDoc.write(outputStream);
                outputStream.close();
                inputStream.close();
            }catch (Exception e){
                System.out.println(e);
            }
        }
    }
    static void createRZ05(){
        for(int i = 0;i < 50;i++){
            String inFileName = muBanLujing[5];if(22<=i && i <=44) {inFileName = muBanLujing[7];}
            try{
                FileInputStream inputStream = new FileInputStream(new File((lujing + inFileName)));
                HWPFDocument inputDoc = new HWPFDocument(inputStream);
                Range mainRan = inputDoc.getRange();
                Range headRan = inputDoc.getHeaderStoryRange();
                String compaName = companyNames[i].split(",")[0].split(" ")[1];
                String bianhao = companyNames[i].split(",")[1];
                //System.out.println(bianhao);
                headRan.replaceText("$jlbh",bianhao);
                mainRan.replaceText("$shenqingdanwei",compaName);
                mainRan.replaceText("$qweqwe","cai");
                FileOutputStream outputStream = new FileOutputStream(new File(outputPath + (i+1) + compaName +"/" + inFileName));
                inputDoc.write(outputStream);
                outputStream.close();
                inputStream.close();
            }catch (IOException e){
                System.out.println("error");
                System.out.println(e);
            }

        }
    }
    static void createRZ07(){
        for(int i = 0;i < 50;i++){
            String inFileName = muBanLujing[9];
            try{
                FileInputStream inputStream = new FileInputStream(new File(lujing + inFileName));
                HWPFDocument inputDoc = new HWPFDocument(inputStream);
                Range headRan = inputDoc.getHeaderStoryRange();
                //Range mainRan = inputDoc.getRange();
                String jlbh = companyNames[i].split(",")[1] + "RZ07";
                String companyName = companyNames[i].split(",")[0].split(" ")[1];
                headRan.replaceText("$jlbh",jlbh);
                FileOutputStream outputStream = new FileOutputStream(new File(outputPath + (i+1) + companyName + "/" + inFileName));
                inputDoc.write(outputStream);
                outputStream.close();
                inputStream.close();
            }catch(Exception e){
                System.out.println(e);
            }
        }
    }
    static void createRZ09(){
        for(int i = 0;i < 50;i++){
            String inFileName = muBanLujing[11];if(22 <= i && i <= 44){inFileName = muBanLujing[14];}
            try{
                FileInputStream inputStream = new FileInputStream(new File(lujing + inFileName));
                HWPFDocument inputDoc = new HWPFDocument(inputStream);
                Range mainRan = inputDoc.getRange();
                String jlbh = companyNames[i].split(",")[1] + "RZ09";
                String sqdw = companyNames[i].split(",")[0].split(" ")[1];
                //mainRan.replaceText("$jlbh",jlbh);
                //mainRan.replaceText("$sqdw",sqdw);
                FileOutputStream outputStream = new FileOutputStream(new File(outputPath + (i+1) + sqdw+"/"+ inFileName));
                inputDoc.write(outputStream);
                outputStream.close();
                inputStream.close();

            }catch (Exception e){
                System.out.println(e);
            }
        }
    }

    static void createRZ10(){
        for(int i = 0;i < 50;i++){
            String inFileName = muBanLujing[15];
            try{
                FileInputStream inputStream = new FileInputStream(new File(lujing + inFileName));
                HWPFDocument inputDoc = new HWPFDocument(inputStream);
                Range headRan = inputDoc.getHeaderStoryRange();
                Range mainRan = inputDoc.getRange();
                String bh = companyNames[i].split(",")[1];
                headRan.replaceText("$1",bh);
                mainRan.replaceText("$1",bh);
                String dj = "3";if(22 <= i && i <= 44){dj = "4";}
                mainRan.replaceText("$2",dj);
                String companyName = companyNames[i].split(",")[0].split(" ")[1];
                FileOutputStream outputStream = new FileOutputStream(new File(outputPath + (i+1) + companyName + "/" + inFileName));
                inputDoc.write(outputStream);
                outputStream.close();
                inputStream.close();
            }catch (Exception e){
                System.out.println(e);
            }
        }
    }
    static void createRZ11(){
        for(int i = 0;i < 50;i++){
            String inFileName = muBanLujing[16];
            try{
                FileInputStream inputStream = new FileInputStream(new File(lujing + inFileName));
                HWPFDocument inputDoc = new HWPFDocument(inputStream);
                Range headRan = inputDoc.getHeaderStoryRange();
                Range mainRan = inputDoc.getRange();
                String bh = companyNames[i].split(",")[1];
                String xmmc = "系统集成资质三级评审";
                String rzzzdj = "三级新申报";if(22 <=i && i<= 44){xmmc = "系统集成资质四级评审";rzzzdj = "四级新申报";}
                String xmjf = companyNames[i].split(",")[0].split(" ")[1];
                String xmyf = "北京赛迪认证中心有限公司";
                headRan.replaceText("$1",bh);
                mainRan.replaceText("$xmmc",xmmc);
                mainRan.replaceText("$zzrzdj",rzzzdj);
                mainRan.replaceText("$xmjf",xmjf);
                mainRan.replaceText("$xmyf",xmyf);
                FileOutputStream outputStream = new FileOutputStream(new File(outputPath + (i+1) + xmjf + "/" + inFileName));
                inputDoc.write(outputStream);
                outputStream.close();
                inputStream.close();
            }catch (Exception e){
                System.out.println(e);
            }
        }
    }
    static void createRZ62(){
        for(int i = 0;i < 50;i++){
            String inFileName = muBanLujing[17];
            try{
                FileInputStream inputStream = new FileInputStream(new File(lujing + inFileName));
                HWPFDocument inputDoc = new HWPFDocument(inputStream);
                Range headRan = inputDoc.getHeaderStoryRange();
                Range mainRan = inputDoc.getRange();
                String bh = companyNames[i].split(",")[1];
                String dwmc = companyNames[i].split(",")[0].split(" ")[1];
                String zzjb = "三级";if(22 <= i && i <= 44){zzjb = "四级";}
                String nian = "2017",month = "5",day = "5";

                headRan.replaceText("$1",bh);
                mainRan.replaceText("$dwmc",dwmc);
                mainRan.replaceText("$zzjb",zzjb);
                mainRan.replaceText("$nian",nian);
                mainRan.replaceText("$month",month);
                mainRan.replaceText("$day",day);
                FileOutputStream outputStream = new FileOutputStream(new File(outputPath + (i+1) + dwmc + "/" + inFileName));
                inputDoc.write(outputStream);
                outputStream.close();
                inputStream.close();
            }catch (Exception e){
                System.out.println(e);
            }
        }
    }
    static void createRZ80(){
        for(int i = 0;i < 50;i++){
            String inFileName = muBanLujing[18];
            try{
                FileInputStream inputStream = new FileInputStream(new File(lujing + inFileName));
                HWPFDocument inputDoc = new HWPFDocument(inputStream);
                Range headRan = inputDoc.getHeaderStoryRange();
                Range mainRan = inputDoc.getRange();
                String bh = companyNames[i].split(",")[1];
                String dwmc = companyNames[i].split(",")[0].split(" ")[1];


                headRan.replaceText("$1",bh);
                mainRan.replaceText("$sbdw",dwmc);

                FileOutputStream outputStream = new FileOutputStream(new File(outputPath + (i+1) + dwmc + "/" + inFileName));
                inputDoc.write(outputStream);
                outputStream.close();
                inputStream.close();
            }catch (Exception e){
                System.out.println(e);
            }
        }
    }
    static void createRZ(){
        createRZ01();
        //createRZ03();
        //createRZ05();
        //createRZ07();
        //createRZ09();
        //createRZ10();
        //createRZ11();
        //createRZ62();
        //createRZ80();
    }

    static void run(){
        createRZ();
    }
    public static void main(String args[]){
        run();
    }
}
