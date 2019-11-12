package com.alibaba.excel;

import com.alibaba.excel.parse.ParseData;
import com.alibaba.excel.parse.ParseDataListener;
import com.alibaba.excel.read.metadata.ReadSheet;
import com.alibaba.excel.write.metadata.WriteSheet;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;


public class excelParse {
    public static void main(String[] args) {
        String fileName = "/Users/changle.zhang/Documents/czf.xlsx";
        List<ParseData> sheet1 = new ArrayList<ParseData>();
        ExcelReader excelReader = EasyExcel.read(fileName, ParseData.class, new ParseDataListener(sheet1)).build();
        ReadSheet readSheet = EasyExcel.readSheet(0).build();
        excelReader.read(readSheet);
        // 这里千万别忘记关闭，读的时候会创建临时文件，到时磁盘会崩的
        System.out.println("key map 大小是 " + sheet1.size());
        // 创建sheet1 has map
        final Map<String, List<String>> nameMap = new HashMap<String, List<String>>();
        for (ParseData parseData : sheet1) {
            String key = parseData.getRow4();
            if (nameMap.get(key) == null) {
                List<String> value = new ArrayList<String>();
                value.add(parseData.getRow5());
                nameMap.put(key, value);
            } else {
                nameMap.get(key).add(parseData.getRow5());
            }
        }
        excelReader.finish();

        // 把sheet2 的key list搂到内存中
        List<ParseData> sheet2 = new ArrayList<ParseData>();
        ExcelReader excelReader1 = EasyExcel.read(fileName, ParseData.class, new ParseDataListener(sheet2)).build();
        ReadSheet readSheet1 = EasyExcel.readSheet(1).build();
        excelReader1.read(readSheet1);
        System.out.println("list 大小是" + sheet2.size());
        List<String> keys = new ArrayList<String>();
        for (ParseData parseData : sheet2) {
            keys.add(parseData.getRow2());
        }
        excelReader1.finish();

        //计算
        List<ParseData> result = new ArrayList<ParseData>();
        for (String key : keys) {
            ParseData tmpData = new ParseData();
            if (nameMap.get(key) == null || nameMap.get(key).size() == 0) {
                tmpData.setRow1(key);
                tmpData.setRow2("N/A");
            } else {
                tmpData.setRow1(key);
                tmpData.setRow2(nameMap.get(key).get(0));
                nameMap.get(key).remove(0);
            }
            result.add(tmpData);
        }

        // 写入
        fileName = "/Users/changle.zhang/Documents/output.xlsx";
        // 这里 需要指定写用哪个class去写
        ExcelWriter excelWriter = EasyExcel.write(fileName, ParseData.class).build();
        WriteSheet writeSheet2 = EasyExcel.writerSheet("结果").build();
        excelWriter.write(result, writeSheet2);
        excelWriter.finish();;
    }
}
