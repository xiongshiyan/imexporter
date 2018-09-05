package top.jfunc.common.excel;

import com.jfinal.plugin.activerecord.Model;
import com.jfinal.plugin.activerecord.Record;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import top.jfunc.common.datetime.DatetimeUtils;
import top.jfunc.common.db.AppendMore;
import top.jfunc.common.utils.BeanUtil;

import java.io.IOException;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.util.Date;
import java.util.List;
import java.util.Map;

/**
 * 导出excel帮助类
 * @author 熊诗言
 */
public class ExportUtils{

    private static final int ROW_ACCESS_WINDOW_SIZE = 1000;
    private SXSSFWorkbook    workbook;
    private Sheet            sheet;
    private CellStyle        style2;
    private CellStyle        style;

    public ExportUtils(String[] headers){
        workbook = new SXSSFWorkbook(ROW_ACCESS_WINDOW_SIZE);
        sheet = workbook.createSheet();

        sheet.setDefaultColumnWidth((short)20);
        style = workbook.createCellStyle();
        // 设置这些样式
        style.setFillForegroundColor(HSSFColor.WHITE.index);
        style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        style.setBorderRight(HSSFCellStyle.BORDER_THIN);
        style.setBorderTop(HSSFCellStyle.BORDER_THIN);
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        // 生成一个字体
        Font font = workbook.createFont();
        font.setBoldweight(HSSFFont.BOLDWEIGHT_NORMAL);
        // 把字体应用到当前的样式
        style.setFont(font);

        style2 = workbook.createCellStyle();
        style2.setFillForegroundColor(HSSFColor.WHITE.index);
        style2.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        style2.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        style2.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        style2.setBorderRight(HSSFCellStyle.BORDER_THIN);
        style2.setBorderTop(HSSFCellStyle.BORDER_THIN);
        style2.setAlignment(HSSFCellStyle.ALIGN_LEFT);
        style2.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
        // 生成另一个字体
        Font font2 = workbook.createFont();
        font2.setBoldweight(HSSFFont.BOLDWEIGHT_NORMAL);
        // 把字体应用到当前的样式
        style2.setFont(font2);

        // 产生表格标题行
        Row row = sheet.createRow(0);
        for(short i = 0; i < headers.length; i++){
            Cell cell = row.createCell(i);
            cell.setCellStyle(style);
            // XSSFRichTextString text = new XSSFRichTextString(headers[i]);
            cell.setCellType(Cell.CELL_TYPE_STRING);
            cell.setCellValue(headers[i]);
        }
    }

    /**
     * 单个工作薄流 excel文件输出
     * 
     * @param dataset
     *            需要显示的数据集合，集合中一定要放置符合JavaBean风格的类的对象。
     * @param pattern
     *            如果有时间数据，设定输出格式。默认为"yyyy-MM-dd"
     * @param rowIndex 行号 从1开始
     * 
     * @throws IOException IO异常
     */
    public void exportExcel(String[] columns, List<?> dataset, String pattern, int rowIndex) throws IOException{
        // 遍历集合数据，产生数据行
        int index = rowIndex;
        Font font3 = workbook.createFont();
        font3.setColor(HSSFColor.BLUE.index);
        for(int i = 0,size = dataset.size(); i < size; i++){
            Row row = sheet.createRow(index++);
            Object data = dataset.get(i);
            for(int col = 0; col < columns.length; col++){
                Cell cell = row.createCell(col);
                Object value = getValue(data,columns[col]);
                setCell(cell, value,  pattern);
            }
            if(i != 0 && i % ROW_ACCESS_WINDOW_SIZE == 0){
                ((SXSSFSheet)sheet).flushRows();
            }
        }
        ((SXSSFSheet)sheet).flushRows();
    }

    private void setCell(Cell cell, Object value, String pattern) {
        cell.setCellStyle(style2);
        cell.setCellType(Cell.CELL_TYPE_STRING);
        if(value == null){
            cell.setCellType(Cell.CELL_TYPE_STRING);
            cell.setCellValue("");
        }else{
            if(value instanceof Integer){
                cell.setCellType(Cell.CELL_TYPE_NUMERIC);
                cell.setCellValue((int)value);
            } else if(value instanceof Long){
                cell.setCellType(Cell.CELL_TYPE_NUMERIC);
                cell.setCellValue((long)value);
            } else if(value instanceof Double){
                cell.setCellType(Cell.CELL_TYPE_NUMERIC);
                cell.setCellValue((double)value);
            } else if(value instanceof BigDecimal){
                cell.setCellType(Cell.CELL_TYPE_NUMERIC);
                cell.setCellValue(((BigDecimal)value).doubleValue());
            } else if(value instanceof Date){
                cell.setCellType(Cell.CELL_TYPE_STRING);
                String date_str = DatetimeUtils.toStr((Date) value,pattern);
                cell.setCellValue(date_str);
            } else{
                cell.setCellType(Cell.CELL_TYPE_STRING);
                cell.setCellValue(value.toString());
            }
        }
    }

    /**
     * 从map、model、record中获取值
     */
    private Object getValue(Object data, String key){
        Object value = null;

        if(data instanceof Model){
            value = ((Model<?>)data).get(key);
        } else if(data instanceof Record){
            value = ((Record)data).get(key);
        }else if(data instanceof Map){
            value = ((Map)data).get(key);
        }else {//JavaBean风格
            value = BeanUtil.get(data,key);
        }
        return value;
    }

    public void write(OutputStream out){
        try{
            workbook.write(out);
            out.flush();
            workbook.close();
            out.close();
        }
        catch(IOException e){
            e.printStackTrace();
        }
    }

    public SXSSFWorkbook getSXSSFWorkbook(){
        return workbook;
    }

    /**
     * 分页查询数据导出到Excel的工具类
     * @param pageNumber 从那一页开始
     * @param appendMore 获取每页数据的接口
     */
    public static <T> void exportToExcel(String[] headers , String[] columns , int pageNumber , int pageSize , OutputStream outputStream , AppendMore<T> appendMore) throws IOException{
        ExportUtils exportUtils = new ExportUtils(headers);
        int rowIndex = 1;
        List<T> list = appendMore.getList(pageNumber , pageSize);
        while(null != list && list.size() >0 ){
            exportUtils.exportExcel(columns,list,"yyyy-MM-dd HH:mm:ss",rowIndex);
            rowIndex += list.size();
            list = appendMore.getList(++pageNumber , pageSize);
        }
        if( null!= outputStream) {
            exportUtils.write(outputStream);
        }
    }
}
