package top.jfunc.common.excel.asyncexport;

import top.jfunc.common.datetime.DatetimeUtils;
import top.jfunc.common.db.DBHelper;
import top.jfunc.common.event.EventKit;
import top.jfunc.common.utils.CommonUtil;

import java.util.Date;
import java.util.HashMap;
import java.util.Map;

/**
 * 用于产生导出excel的文件名等
 * @author 熊诗言
 */
public class ExcelExportKit {
    private ExcelExportKit(){}
    /**
     * 报表导出失败返回
     * @param message 失败消息
     */
    public static Map<String,Object> createDownloadErrorJSONObject(String message){
        Map<String,Object> map = new HashMap<>(2);
        map.put("status", 0);
        map.put("message", "导出失败：" + message);
        return map;
    }
    /**
     * 报表导出开始的返回信息，传入taskId，现在报表的导出异步化
     * @param taskId 任务ID
     */
    public static Map<String,Object> createDownloadBegingJSONObject(String taskId){
        Map<String,Object> map = new HashMap<>(2);
        map.put("status", 1);
        map.put("message", "导出开始,点击去下载,若没有则稍后刷新,也可以到\"报表中心/导出文件下载\",任务号是 " + taskId);
        map.put("taskId", taskId);
        return map;
    }
    /**
     * 获取文件名，根据日期创建导出文件名
     * @return String:文件名
     */
    public static  String generateFileName(String prefix){
        String fileName = prefix + DatetimeUtils.toStr(new Date(),DatetimeUtils.SDF_DATETIME_MS) + ".xlsx";
        return fileName;
    }

    /**
     * 发布一个事件
     * @param sql 查询的SQL语句
     * @param headers 头
     * @param columns 列
     * @param filePath 文件路径
     */
    public static Map<String,Object> postExportEvent(DBHelper helper,String sql, String[] headers, String[] columns, String filePath) {
        return postExportEvent(helper,sql,headers,columns,filePath,null);
    }
    /**
     * 发布一个事件
     * @param sql 查询的SQL语句
     * @param headers 头
     * @param columns 列
     * @param filePath 文件绝对路径
     * @param callback 回调
     */
    public static Map<String,Object> postExportEvent(DBHelper helper,String sql, String[] headers, String[] columns, String filePath, Runnable callback) {
        Map<String,Object> jo = null;
        try{
            String taskId = CommonUtil.getUUID();
            //异步执行导出
            EventKit.post(
                    new ExportExcelEvent(
                            new ExportExcelBean(helper,headers,columns,sql,filePath,callback)));
            jo = createDownloadBegingJSONObject(taskId);

        }
        catch(Exception e){
            e.printStackTrace();
            jo = createDownloadErrorJSONObject(e.getMessage());
        }
        return jo;
    }
}
