package top.jfunc.common.excel.asyncexport;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import top.jfunc.common.event.core.ApplicationListener;
import top.jfunc.common.event.core.EventType;
import top.jfunc.common.event.core.ListenerAttr;
import top.jfunc.common.excel.ExcelLargeDataUtils;

/**
 * ExportExcelEvent事件监听器
 * @author 熊诗言
 */
public class ExportExcelListener implements ApplicationListener<ExportExcelEvent> ,ListenerAttr{
    private static final Logger logger = LoggerFactory.getLogger(ExportExcelListener.class);
    @Override
    public void onApplicationEvent(ExportExcelEvent event) {
        ExportExcelBean bean = (ExportExcelBean)event.getSource();
        String[] headers = bean.getHeaders();
        String[] columns = bean.getColumns();
        String sql = bean.getSql();
        String exportFilePath = bean.getExportFilePath();
        Runnable callback = bean.getCallback();
        logger.info(exportFilePath);

        try {
            /*下载文件*/
            ExcelLargeDataUtils.exportData(bean.getHelper(),headers,columns,sql,exportFilePath);
            /*回调方法*/
            if(null != callback){
                callback.run();
            }
        } catch (Exception e) {
            logger.error(e.getMessage());
        }

    }

    @Override
    public int getOrder() {
        return 1;
    }

    @Override
    public boolean getEnableAsync() {
        return true;
    }

    @Override
    public String getTag() {
        return EventType.DEFAULT_TAG;
    }
}
