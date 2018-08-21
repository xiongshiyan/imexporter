package top.jfunc.common.utils;

import com.jfinal.plugin.activerecord.Record;
import org.junit.Ignore;
import org.junit.Test;
import top.jfunc.common.excel.AppendMoreData;
import top.jfunc.common.excel.ExportUtils;
import top.jfunc.common.fastexcel.FastExcelExporter;

import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author 熊诗言 at 2017/11/23
 */
public class FastExcelTest {
    @Test
    @Ignore
    public void testFastOneSheet() throws Exception{
        String[] headers = new String[]{"列1","列2"};
        String[] columns = new String[]{"sd","sa"};
        List<Map<String,String>> data = new ArrayList<>();
        Map<String,String> map1 = new HashMap<>();map1.put(columns[0],"sd");map1.put(columns[1],"sa");data.add(map1);
        Map<String,String> map2 = new HashMap<>();map2.put(columns[0],"ss");map2.put(columns[1],"sf");data.add(map2);
        FileOutputStream os = new FileOutputStream("C:\\Users\\xiongshiyan\\Desktop\\2.xls");
        FastExcelExporter.data(data).append((AppendMoreData)null).application("1.xls").os(os).version("1.0").sheetNames("sheet1").headers(headers).columns(columns).export();
    }
    @Test
    @Ignore
    public void testFastTwoSheet() throws Exception{
        /*sheet1*/
        String[] headers1 = new String[]{"列11","列12"};
        String[] columns1 = new String[]{"c11","c12"};
        List<Map<String,String>> data1 = new ArrayList<>();
        Map<String,String> map11 = new HashMap<>();map11.put(columns1[0],"sd");map11.put(columns1[1],"sa");data1.add(map11);
        Map<String,String> map12 = new HashMap<>();map12.put(columns1[0],"ss");map12.put(columns1[1],"sf");data1.add(map12);

        /*sheet2*/
        String[] headers2 = new String[]{"列21","列22"};
        String[] columns2 = new String[]{"c21","c22"};
        List<Map<String,String>> data2 = new ArrayList<>();
        Map<String,String> map21 = new HashMap<>();map21.put(columns2[0],"sd");map21.put(columns2[1],"sa");data2.add(map21);
        Map<String,String> map22 = new HashMap<>();map22.put(columns2[0],"ss");map22.put(columns2[1],"sf");data2.add(map22);
        FileOutputStream os = new FileOutputStream("C:\\Users\\xiongshiyan\\Desktop\\2.xls");
        FastExcelExporter.data(data1,data2).append((AppendMoreData)null).application("2.xls").os(os).version("1.0").sheetNames("sheet1","sheet2").headers(headers1,headers2).columns(columns1,columns2).export();
    }

    @Test
    @Ignore
    public void testFastOneSheetAppendData() throws Exception{
        //另外一种可追加数据的excel
        String[] headers = new String[]{"列1","列2"};
        String[] columns = new String[]{"sd","sa"};
        FileOutputStream os = new FileOutputStream("C:\\Users\\xiongshiyan\\Desktop\\3.xls");
//        List<Map<String,String>> data = new ArrayList<>();
//        Map<String,String> map1 = new HashMap<>();map1.put(columns[0],"sd");map1.put(columns[1],"sa");data.add(map1);
//        Map<String,String> map2 = new HashMap<>();map2.put(columns[0],"ss");map2.put(columns[1],"sf");data.add(map2);
        new FastExcelExporter().append(new AppendMoreData() {
            int i = 0;
            @Override
            public List<?> append() {
                List<Map<String,Object>> data = new ArrayList<>();
                int j = i++;
                if(i==9){return null;}
                Map<String,Object> map1 = new HashMap<>();map1.put(columns[0],"sd"+j);map1.put(columns[1],"sa"+j);data.add(map1);
                System.out.println("============================================"+j);
                return data;
            }
        }).application("3.xls").os(os).version("1.0").sheetNames("sheet1").headers(headers).columns(columns).export();
    }

    @Test
    @Ignore
    public void testFastTwoSheetAppendData() throws Exception{
        //另外一种可追加数据的excel
        String[] headers = new String[]{"列1","列2"};
        String[] columns = new String[]{"sd","sa"};
        FileOutputStream os = new FileOutputStream("C:\\Users\\xiongshiyan\\Desktop\\4.xls");
        new FastExcelExporter().append(new AppendMoreData() {//sheet1
                                            int i = 0;
                                            @Override
                                            public List<?> append() {
                                                List<Map<String,Object>> data = new ArrayList<>();
                                                int j = i++;
                                                if(i==10){return null;}
                                                Map<String,Object> map1 = new HashMap<>();map1.put(columns[0],"sd"+j);map1.put(columns[1],"sa"+j);data.add(map1);
                                                System.out.println("============================================"+j);
                                                return data;
                                            }
                                        },
                                        new AppendMoreData() {//sheet2
                                            int i = 0;
                                            @Override
                                            public List<?> append() {
                                                List<Map<String,Object>> data = new ArrayList<>();
                                                int j = i++;
                                                if(i==11){return null;}
                                                Map<String,Object> map1 = new HashMap<>();map1.put(columns[0],"sd"+j);map1.put(columns[1],"sa"+j);data.add(map1);
                                                System.out.println("============================================"+j);
                                                return data;
                                            }
                                        }
                                ).application("3.xls").os(os).version("1.0").sheetNames("sheet1","sheet2").headers(headers,headers).columns(columns,columns).export();
    }


    /**
     * 多次导出到同一个excel,一次10个，100次
     * java.lang.IllegalArgumentException: Invalid row number (1048576) outside allowable range (0..1048575)
     */
    @Test
    @Ignore
    public void testExportMore() throws Exception{
        String[] headers = new String[]{"col1","col2"};
        String[] columns = new String[]{"col1","col2"};
        ExportUtils utils = new ExportUtils(headers);
        int lenIn = 500;
        for (int j = 0; j < 1000; j++) {
            List<Object> dataset = new ArrayList<>(10);
            for (int i = 0; i < lenIn; i++) {
                Record record = new Record();
                record.set("col1",j*lenIn+i);
                record.set("col2","co11strrecord " + (j*lenIn+i));
                dataset.add(record);
            }
            for (int i = 0; i < lenIn; i++) {
                Map<String,Object> map = new HashMap<>(2);
                map.put("col1",j*lenIn+i);
                map.put("col2","co11strmap " + (j*lenIn+i));
                dataset.add(map);
            }
            int rowIndex = j*lenIn*2 + 1;
            utils.exportExcel(columns,dataset,"",rowIndex);
        }

        OutputStream stream = new FileOutputStream("C:\\Users\\xiongshiyan\\Desktop\\xx.xls");
        utils.write(stream);
    }
}
