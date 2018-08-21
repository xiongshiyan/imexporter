package top.jfunc.common.excel.asyncexport;

import top.jfunc.common.db.DBHelper;

/**
 * 报表导出数据结构
 */
class ExportExcelBean {
    /**
     * 数据库连接信息
     */
    DBHelper helper;
    /**
     * 报表头
     */
    private String[] headers;
    /**
     * 字段名
     */
    private String[] columns;
    /**
     * 查询SQL语句
     */
    private String sql;
    /**
     * 文件输出路径，绝对路径
     */
    private String exportFilePath;
    /**
     * 回调方法
     */
    private Runnable callback;

    public ExportExcelBean(DBHelper helper, String[] headers, String[] columns, String sql, String exportFilePath, Runnable callback) {
        this.helper = helper;
        this.headers = headers;
        this.columns = columns;
        this.sql = sql;
        this.exportFilePath = exportFilePath;
        this.callback = callback;
    }
    public ExportExcelBean(DBHelper helper, String[] headers, String[] columns, String sql, String exportFilePath) {
        this(helper,headers,columns,sql,exportFilePath,null);
    }

    public DBHelper getHelper() { return helper;}

    public String[] getHeaders() {
        return headers;
    }

    public String[] getColumns() {
        return columns;
    }

    public String getSql() {
        return sql;
    }

    public String getExportFilePath() {
        return exportFilePath;
    }

    public Runnable getCallback() { return callback;}
}
