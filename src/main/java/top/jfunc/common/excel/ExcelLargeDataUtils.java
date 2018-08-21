package top.jfunc.common.excel;

import com.jfinal.plugin.activerecord.Record;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import top.jfunc.common.ProgressNotifier;
import top.jfunc.common.db.DBHelper;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.List;

/**
 * 导出大量数据
 * @author 熊诗言 2017/10/11
 */
public class ExcelLargeDataUtils {
    private static Logger logger = LoggerFactory.getLogger(ExcelLargeDataUtils.class);

    public static void exportData(final DBHelper helper,final String[] headers, final String[] columns, String sql, String exportFilePath) throws Exception {
        exportData(helper,headers,columns,sql,exportFilePath,null);
    }
    public static void exportData(final Connection connection,final String[] headers, final String[] columns, String sql, String exportFilePath) throws Exception {
        exportData(connection,headers,columns,sql,exportFilePath,null);
    }
    /**
     * @param helper:提供数据库连接的
     * @param headers:导出文件列标题
     * @param columns:导出数据对应库字段
     * @param sql:导出SQL查询语句
     * @param exportFilePath:导出文件保存路径
     * @throws Exception Exception
     */
    public static void exportData(final DBHelper helper, final String[] headers, final String[] columns, String sql, String exportFilePath, ProgressNotifier progress) throws Exception {
        long startAll = System.currentTimeMillis();
        File file = new File(exportFilePath);
        if(!file.exists()){
            file.createNewFile();
        }
        FileOutputStream fos = null;
        PreparedStatement pst = null;
        try {
            fos = new FileOutputStream(exportFilePath);
            /* 建立数据库连接 */
            pst = helper.getPrepareStatement(sql);
            /* 查询SQL获取结果 */
            ResultSet resultSet = pst.executeQuery();
            /* 将查询结果集导出到excel文件中 */
            int count = export2Excel(headers, columns, fos, resultSet, progress);
            resultSet.close();
            logger.info(String.format("count=%d,use time=%d", count, (System.currentTimeMillis() - startAll)));
        } catch (SQLException | IOException e) {
            logger.error(e.getMessage() , e);
            if(null!=progress){
                progress.error(e);
            }
            throw new Exception(e);
        } finally {
            close(fos, pst, helper);
        }
    }
    /**
     * @param connection:数据库连接
     * @param headers:导出文件列标题
     * @param columns:导出数据对应库字段
     * @param sql:导出SQL查询语句
     * @param exportFilePath:导出文件保存路径
     * @throws Exception Exception
     */
    public static void exportData(final Connection connection, final String[] headers, final String[] columns, String sql, String exportFilePath, ProgressNotifier progress) throws Exception {
        long startAll = System.currentTimeMillis();
        File file = new File(exportFilePath);
        if(!file.exists()){
            boolean newFile = file.createNewFile();
            if(!newFile){
                logger.error("生成文件失败 : " + file.getAbsolutePath());
            }
        }
        FileOutputStream fos = null;
        PreparedStatement pst = null;
        try {
            fos = new FileOutputStream(exportFilePath);
            /* 建立数据库连接 */
            pst = connection.prepareStatement(sql, ResultSet.TYPE_FORWARD_ONLY, ResultSet.CONCUR_READ_ONLY);
            pst.setFetchSize(Integer.MIN_VALUE);
            pst.setFetchDirection(ResultSet.FETCH_REVERSE);
            /* 查询SQL获取结果 */
            ResultSet resultSet = pst.executeQuery();
            /* 将查询结果集导出到excel文件中 */
            int count = export2Excel(headers, columns, fos, resultSet, progress);
            resultSet.close();
            logger.info(String.format("count=%d,use time=%d", count, (System.currentTimeMillis() - startAll)));
        } catch (Exception e) {
            logger.error(e.getMessage() , e);
            if(null!=progress){
                progress.error(e);
            }
            throw new Exception(e);
        } finally {
            close(fos, pst, null);
        }
    }
    /**
     * @param headers:导出文件列标题
     * @param columns:导出数据对应库字段
     * @param fos:输出文件流
     * @param resultSet:数据查询结果集
     * @throws SQLException SQLException
     * @throws IOException IOException
     */
    private static int export2Excel(final String[] headers, final String[] columns, FileOutputStream fos,
                                    ResultSet resultSet,ProgressNotifier progress) throws SQLException, IOException {
        int count = 0;
        List<Record> list = new ArrayList<Record>();
        int rowIndex = 1;
        ExportUtils exportUtils = new ExportUtils(headers);
        if(null != progress){
            progress.start();
        }
        while (resultSet.next()) {
            Record record = new Record();
            for (int i = 0; i < columns.length; i++) {
                record.set(columns[i], resultSet.getObject(columns[i]));
            }
            list.add(record);
            count++;
            if (count % 5000 == 0) {
                if(null != progress){
                    progress.progressed(count);
                }
                exportUtils.exportExcel(columns, list, "", rowIndex);
                rowIndex += list.size();
                list.clear();
            }
        }
        if (list.size() > 0) {
            if(null != progress){
                progress.progressed(count);
            }
            exportUtils.exportExcel(columns, list, "", rowIndex);
            list.clear();
        }
        exportUtils.write(fos);
        if(null!=progress){
            progress.finish(count);
        }
        return count;
    }


    private static void close(FileOutputStream fos, PreparedStatement pst, DBHelper dbHelper) {
        if (fos != null) {
            try {
                fos.close();
            } catch (IOException e) {
                logger.error("fos关闭异常" , e);
            }
        }
        if (pst != null) {
            try {
                pst.close();
            } catch (SQLException e) {
                logger.error("pst关闭异常" , e);
            }
        }
        if (dbHelper != null) {
            dbHelper.close();
        }
    }


}
