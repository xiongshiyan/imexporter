package top.jfunc.common.report;

import com.jfinal.plugin.activerecord.Record;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStreamWriter;
import java.util.List;

/**
 * @author xiongshiyan
 * 导出TXT文件
 */
public abstract class AbstractExportUploadTxtService extends AbstractExportUploadService<Record> {
    private static final Logger   logger         = LoggerFactory.getLogger(AbstractExportUploadTxtService.class);

    /**
     * record转一行
     * @param item Record数据库记录
     * @return 记录转换成一行
     */
    protected abstract String convertRecordToString(Record item);

    /**
     * 每一行的前缀
     * @param record 如果跟订单相关的话，record就有用，否则忽略即可
     * @return 前缀
     */
    protected abstract String linePrefix(Record record);


    /**
     * 文件的第一行，一般是一些文件头字段说明什么的，返回null就不写
     * @return 第一行
     */
    protected String firstLine(){return null;}


    /**
     * 文件的最后一行，一般是统计数据，返回null就不写
     * @return 最后一行
     */
    protected String lastLine(){return null;}


    /**
     * 文件编码
     * @return 文件编码
     */
    protected String charset(){
        return "GBK";
    }


    @Override
    protected void export(String fileName, String filePath) {
        try {
            exportData(filePath , firstLine() , lastLine());
        } catch (Exception e) {
            logger.error(e.getMessage() , e);
            throw new RuntimeException(e);
        }
    }

    private void exportData(String filePath , String firstLine , String lastLine) throws Exception{
        int i = 1;
        try(OutputStreamWriter osr = new OutputStreamWriter(new FileOutputStream(new File(filePath)), charset())){

            //1.处理第一行
            if(null != firstLine){
                osr.write(firstLine);
                osr.write("\r\n");
            }

            //2.处理数据
            while(true){
                List<Record> list = getList( i++, PAGE_SIZE);
                if(null == list || list.size() == 0){
                    break;
                }
                outputToFile(list, osr);
                if(list.size() < PAGE_SIZE){
                    list.clear();
                    break;
                }
                list.clear();
            }

            //3.处理最后一行
            if(null != lastLine){
                osr.write(lastLine);
                osr.write("\r\n");
            }
        }
    }


    private void outputToFile(List<Record> list, OutputStreamWriter osr){
        for(int i = 0; i < list.size(); i++){
            try{
                Record item = list.get(i);
                osr.write(linePrefix(item) + convertRecordToString(item));
                osr.write("\r\n");
            }
            catch(Exception e){
                logger.error(e.getMessage() , e);
                continue;
            }
        }
    }
}
