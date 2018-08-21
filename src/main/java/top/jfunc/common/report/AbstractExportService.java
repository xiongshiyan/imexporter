package top.jfunc.common.report;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import top.jfunc.common.db.AppendMore;

import java.util.List;

/**
 * @author xiongshiyan
 * 导出文件
 */
public abstract class AbstractExportService<T> implements AppendMore<T> {
    private static final Logger logger = LoggerFactory.getLogger(AbstractExportService.class);

    protected static int      PAGE_SIZE      = 5000;
    /**
     * 获取查询数据，不要获取查询的SQL，万一是从别的库查询呢
     * 此方法会被多次调用，传入不同的pageNumber，返回null或者空集合即停止
     * @param pageNumber 页数
     * @param pageSize 每页数
     * @return 查询记录
     */
    @Override
    public abstract List<T> getList(int pageNumber , int pageSize);

    /**
     * 获取导出TEXT文件名
     * @return 文件名
     */
    protected abstract String getFileName();

    /**
     * 本地文件保存绝对路径
     * @param fileName 文件名称
     * @return 绝对路径
     */
    protected abstract String getFilePath(String fileName);


    /**
     * 具体的导出
     * @param fileName 文件名
     * @param filePath 文件路径
     */
    protected abstract void export(String fileName , String filePath);

    /**
     * 导出之后干的事情，比如上传FTP
     */
    protected abstract void afterExport();

    /**
     * 导出并再做什么
     */
    public void exportAndFinish() throws Exception {
        String fileName = getFileName();
        String filePath = getFilePath(fileName);
        logger.info("fileName : " + fileName);
        logger.info("filePath : " + filePath);

        export(fileName , filePath);

        afterExport();
    }

}
