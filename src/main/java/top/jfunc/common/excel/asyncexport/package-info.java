package top.jfunc.common.excel.asyncexport;

/**
 * 该包主要用于异步导出Excel用
 * 1. 将ExportExcelListener加入到事件插件中EventPlugin，确保启动
 * 2. ExcelExportKit.postExportEvent()发送导出事件，还可以传入回调方法，在完成导出之后回调
 * 回调的典型使用方式是保存一条导出记录，让用户下载
 *
 * 虽然能异步导出大的Excel，但是还是有限制的，excel2007开始最大行就到1048576行
 */