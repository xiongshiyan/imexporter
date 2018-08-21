/**
 * Copyright (c) 2011-2013, kidzhou 周磊 (zhouleib1412@gmail.com)
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package top.jfunc.common.fastexcel;

import com.jfinal.plugin.activerecord.Model;
import com.jfinal.plugin.activerecord.Record;
import top.jfunc.common.excel.AppendMoreData;

import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Set;

/**
 * FastExcelExporter.data(data).append(append).application(fileName).os(os).version(version).sheetNames(sheetNames).headers(headers).columns(columns).export();
 */
public class FastExcelExporter {

    public static final String VERSION_2003 = "2003";
    private static final int HEADER_ROW = 1;
    private static final int MAX_ROWS = 65535;
    private String version;
    private String[] sheetNames = new String[]{"sheet"};
    private int cellWidth = 8000;
    private int headerRow;
    private String[][] headers;
    private String[][] columns;
    private List<?>[] data;
    private AppendMoreData[] append;
    private String application;
    private OutputStream os;
    public FastExcelExporter(List<?>... data) {
        this.data = data;
    }

    public static FastExcelExporter data(List<?>... data) {
        return new FastExcelExporter(data);
    }

    public static List<List<?>> dice(List<?> num, int chunkSize) {
        int size = num.size();
        int chunk_num = size / chunkSize + (size % chunkSize == 0 ? 0 : 1);
        List<List<?>> result = new ArrayList<>();
        for (int i = 0; i < chunk_num; i++) {
            result.add(new ArrayList<>(num.subList(i * chunkSize, i == chunk_num - 1 ? size : (i + 1) * chunkSize)));
        }
        return result;
    }

    public void export() throws IOException {
        if((data == null || data.length==0)&& (append != null && append.length>0)) {
            List<?>[] firstData= new List<?>[append.length];
            for (int i = 0; i < append.length; i++) {
                firstData[i] = append[i].append();
            }
            set(firstData);
        }
        if(null==data){throw new NullPointerException("data can not be null");}
        if(null==headers){throw new NullPointerException("headers can not be null");}
        if(null==columns){throw new NullPointerException("columns can not be null");}
        if(!(data.length == sheetNames.length && sheetNames.length == headers.length
                && headers.length == columns.length)){
            throw new IllegalArgumentException("data,sheetNames,headers and columns'length should be the same." +
                "(data:" + data.length + ",sheetNames:" + sheetNames.length + ",headers:" + headers.length + ",columns:" + columns.length + ")");}
        if(cellWidth < 0){
            throw new IllegalArgumentException("cellWidth can not be less than 0");}

        Workbook wb = new Workbook(os, application, version);
        
        if (data.length == 0) {
            return;
        }
        for (int i = 0; i < data.length; i++) {
            Worksheet ws = wb.newWorksheet(sheetNames[i] == null ? "Sheet " + i : sheetNames[i]);
            handleHeader(ws, headers[i]);
            List<?> temp = data[i];
            int current = 0;
            do {
                handleData(i, ws, temp, current);
                current += temp.size();
            }while(append != null && (temp = append[i].append()) != null);
        }
        wb.finish();
    }

    private void handleHeader(Worksheet ws, String[] header) {
        if (header.length > 0) {
            if (headerRow <= 0) {
                headerRow = HEADER_ROW;
            }
            headerRow = Math.min(headerRow, MAX_ROWS);
            for (int h = 0, lenH = header.length; h < lenH; h++) {
                ws.value(0, h, header[h]);
            }
        }
    }

    /**
     *
     * @param sheetIndex sheet索引
     * @param ws
     * @param data 数据
     * @param dataIndex 数据行数索引
     */
    private void handleData(int sheetIndex, Worksheet ws, List<?> data, int dataIndex) {
        for (int j = 0, len = data.size(); j < len; j++) {
            Object obj = data.get(j);
            if (obj == null) {
                continue;
            }
            if (obj instanceof Map) {
                processAsMap(columns[sheetIndex], ws,j + headerRow + dataIndex, obj);
            } else if (obj instanceof Model) {
                processAsModel(columns[sheetIndex], ws,j + headerRow + dataIndex, obj);
            } else if (obj instanceof Record) {
                processAsRecord(columns[sheetIndex], ws,j + headerRow + dataIndex, obj);
            } else {
                throw new RuntimeException("Not support type[" + obj.getClass() + "]");
            }
        }
    }

    @SuppressWarnings("unchecked")
    private static void processAsMap(String[] columns, Worksheet ws, int index,Object obj) {
        Map<String, Object> map = (Map<String, Object>) obj;
        if (columns.length == 0) { // show all if column not specified
            Set<String> keys = map.keySet();
            int columnIndex = 0;
            for (String key : keys) {
                ws.value(index,columnIndex, map.get(key));
                columnIndex++;
            }
        } else {
            for (int j = 0, len = columns.length; j < len; j++) {
                ws.value(index,j,map.get(columns[j]) == null ? "" : map.get(columns[j]));
            }
        }
    }

    private static void processAsModel(String[] columns, Worksheet ws, int index, Object obj) {
        Model<?> model = (Model<?>) obj;
        Set<Entry<String, Object>> entries = model._getAttrsEntrySet();
        if (columns.length == 0) { // show all if column not specified
            int columnIndex = 0;
            for (Entry<String, Object> entry : entries) {
                ws.value(index,columnIndex, entry.getValue());
                columnIndex++;
            }
        } else {
            for (int j = 0, len = columns.length; j < len; j++) {
                ws.value(index,j,model.get(columns[j]) == null ? "" : model.get(columns[j]));
            }
        }
    }

    private static void processAsRecord(String[] columns, Worksheet ws, int index, Object obj) {
        Record record = (Record) obj;
        Map<String, Object> map = record.getColumns();
        if (columns.length == 0) { // show all if column not specified
            record.getColumns();
            Set<String> keys = map.keySet();
            int columnIndex = 0;
            for (String key : keys) {
                ws.value(index,columnIndex, record.get(key));
                columnIndex++;
            }
        } else {
            for (int j = 0, len = columns.length; j < len; j++) {
                ws.value(index,j,map.get(columns[j]) == null ? "" : map.get(columns[j]));
            }
        }
    }
    public FastExcelExporter application(String name) {
        this.application = name;
        return this;
    }
    public FastExcelExporter os(OutputStream os) {
        this.os = os;
        return this;
    }
    
    public FastExcelExporter version(String version) {
        this.version = version;
        return this;
    }

    public FastExcelExporter sheetName(String sheetName) {
        this.sheetNames = new String[]{sheetName};
        return this;
    }

    public FastExcelExporter sheetNames(String... sheetName) {
        this.sheetNames = sheetName;
        return this;
    }

    public FastExcelExporter cellWidth(int cellWidth) {
        this.cellWidth = cellWidth;
        return this;
    }

    public FastExcelExporter headerRow(int headerRow) {
        this.headerRow = headerRow;
        return this;
    }

    public FastExcelExporter header(String... header) {
        this.headers = new String[][]{header};
        return this;
    }

    public FastExcelExporter headers(String[]... headers) {
        this.headers = headers;
        return this;
    }

    public FastExcelExporter column(String... column) {
        this.columns = new String[][]{column};
        return this;
    }

    public FastExcelExporter columns(String[]... columns) {
        this.columns = columns;
        return this;
    }

    public FastExcelExporter set(List<?>... data) {
        this.data = data;
        return this;
    }
    public FastExcelExporter append(AppendMoreData... append){
        this.append = append;
        return this;
    }

}
