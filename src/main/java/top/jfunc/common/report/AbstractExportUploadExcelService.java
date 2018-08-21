package top.jfunc.common.report;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import top.jfunc.common.excel.ExportUtils;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

/**
 * @author xiongshiyan
 * 导出Excel文件
 */
public abstract class AbstractExportUploadExcelService<T> extends AbstractExportUploadService<T> {
    private static final Logger logger = LoggerFactory.getLogger(AbstractExportUploadExcelService.class);


    @Override
    protected void export(String fileName, String filePath) {
        try (OutputStream stream = new FileOutputStream(filePath)){
            ExportUtils.exportToExcel(getHeaders(), getColumns(), 1, PAGE_SIZE, stream, this);
        } catch (IOException e) {
            logger.error(e.getMessage(), e);
            throw new RuntimeException(e);
        }
    }

    /**
     * 文件导出的头
     * @return excel文件头
     */
    protected abstract String[] getHeaders();

    /**
     * 文件导出的列
     * @return excel文件列
     */
    protected abstract String[] getColumns();
}