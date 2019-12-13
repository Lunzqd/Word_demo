package com.example.demo;

import com.example.utils.CommonUtil;
import com.example.utils.CustomXWPFDocument;
import com.example.utils.NewCfUtils;
import lombok.extern.slf4j.Slf4j;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;

@Slf4j
@SpringBootApplication
public class DemoApplication {
    public static void main(String[] args) {

    }

    public void testDemo() throws IOException {

        Map<String, Object> data = new HashMap<>();
        Map<String, Object> dataInfo = new HashMap<>();
        List<Map<String, Object>> resultData = new ArrayList<>();
        // 模板准备
        // 在线文档需先下载
        String templateUrl = "E:/InsTemplate2.docx";

        SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddHHmmss");
        String date = sdf.format(new Date());

        // 图片准备
        // 在线图片需先下载
        String leftLogoUrl = "E:/leftLogo.png";
        String rightLogoUrl = "E:/rightLogo.png";


        // 顶部logo --左上

        if (CommonUtil.isValidStr(leftLogoUrl)) {
            Map<String, Object> leftLogo = new HashMap<>();
            leftLogo.put("width", 30);
            leftLogo.put("height", 30);
            leftLogo.put("type", "png");
            leftLogo.put("content", NewCfUtils.inputStream2ByteArray(new FileInputStream(leftLogoUrl), true));
            data.put("${leftLogo}", leftLogo);

        } else {
            data.put("${leftLogo}", "");
        }

        // 顶部logo -右上
        if (CommonUtil.isValidStr(rightLogoUrl)) {

            Map<String, Object> rightLogo = new HashMap<>();
            rightLogo.put("width", 180);
            rightLogo.put("height", 30);
            rightLogo.put("type", "png");
            rightLogo.put("content", NewCfUtils.inputStream2ByteArray(new FileInputStream(rightLogoUrl), true));
            data.put("${rightLogo}", rightLogo);

        } else {
            data.put("${rightLogo}", "");
        }


        // 数据准备

        // 纯段落文本
        data.put("${docName}", "一份严肃的文档");

        // 表格文本
        data.put("${docCode}", "编号：2022-01-01-0001");
        //单列，纯文字
        List<String> firstTableData = this.getFirstTableData();
        dataInfo = new HashMap<>();
        dataInfo.put("index", 2);// 序号
        dataInfo.put("columns", 1);// 列数
        dataInfo.put("data", firstTableData);// 数据
        dataInfo.put("fontSize", 10);// 字体大小
        dataInfo.put("startLine", 2);// 第二行开始填入数据
        resultData.add(dataInfo);

        //单列，带图片
        List<String> secTableData = this.getSecTableData();
        dataInfo = new HashMap<>();
        dataInfo.put("index", 3);// 序号
        dataInfo.put("columns", 1);// 列数
        dataInfo.put("data", secTableData);// 数据
        dataInfo.put("fontSize", 10);// 字体大小
        dataInfo.put("startLine", 1);// 第一行开始填入数据
        resultData.add(dataInfo);

        //多列，纯文字
        List<String[]> thirdTableData = this.getThirdTableData();
        dataInfo = new HashMap<>();
        dataInfo.put("index", 4);// 序号
        dataInfo.put("columns", 4);// 列数
        dataInfo.put("fontSize", 10);// 字体大小
        dataInfo.put("data", thirdTableData);// 数据
        dataInfo.put("startLine", 1);// 第一行开始填入数据
        resultData.add(dataInfo);

        //多列，带图片
        List<String[]> forthTableData = this.getForthTableData();
        dataInfo = new HashMap<>();
        dataInfo.put("index", 5);// 序号
        dataInfo.put("columns", 4);// 列数
        dataInfo.put("fontSize", 10);// 字体大小
        dataInfo.put("data", forthTableData);// 数据
        dataInfo.put("startLine", 1);// 第一行开始填入数据
        resultData.add(dataInfo);


        // 生成文档
        CustomXWPFDocument doc = NewCfUtils.changWord(templateUrl, data, resultData);
        String docName = "cfTest_" + date + ".docx";
        String docUrl = "E:/" + docName;
        log.info("[生成文档]-{}", docUrl);
        FileOutputStream outputStream = new FileOutputStream(docUrl);
        doc.write(outputStream);
        outputStream.close();

    }

    private List<String> getFirstTableData() {

        List<String> result = new ArrayList<>();

        result.add("1、蓝教练是女教练，吕教练是男教练，蓝教练不是男教练，吕教练不是女教练。");
        result.add("2、石小四年十四，史肖石年四十。年十四的石小四爱看诗词，年四十的史肖石爱看报纸。");
        result.add("3、九十九头牛，驮着九十九个篓。每篓装着九十九斤油。");
        result.add("4、牛郎恋刘娘,刘娘念牛郎,牛郎牛年恋刘娘,刘娘年年念牛郎,郎恋娘来娘念郎,念娘恋郎,念恋娘郎。");

        return result;
    }

    private List<String> getSecTableData() {
        String pic = "E:/baidu.png";

        List<String> result = new ArrayList<>();
        result.add("姓名：BaiDu--" + pic + "--1");
        result.add("（签章）");
        result.add("地址：xxxxxxxxxxxxxx");
        result.add("电话：131xxxx0001");

        return result;
    }

    private List<String[]> getThirdTableData() {

        List<String[]> result = new ArrayList<>();
        result.add(new String[]{"1", "01", "001", "0001"});
        result.add(new String[]{"2", "02", "002", "0002"});
        result.add(new String[]{"3", "03", "003", "0003"});
        result.add(new String[]{"4", "04", "004", "0004"});
        result.add(new String[]{"5", "05", "005", "0005"});
        result.add(new String[]{"6", "06", "006", "0006"});
        result.add(new String[]{"7", "07", "007", "0007"});
        result.add(new String[]{"8", "08", "008", "0008"});
        result.add(new String[]{"9", "09", "009", "0009"});
        return result;
    }

    private List<String[]> getForthTableData() {

        List<String[]> result = new ArrayList<>();
        String pic01 = "E:/Piggy.png";
        String pic02 = "E:/fighting.png";
        result.add(new String[]{"1--" + pic01 + "--1", "01--" + pic02 + "--2", "001", "1001"});
        result.add(new String[]{"2", "02", "002", "2002"});
        result.add(new String[]{"3", "03", "003", "3003"});
        result.add(new String[]{"4", "04", "004", "4004"});
        result.add(new String[]{"5", "05", "005", "5005"});
        result.add(new String[]{"6", "06", "006", "6006"});
        return result;
    }
}
