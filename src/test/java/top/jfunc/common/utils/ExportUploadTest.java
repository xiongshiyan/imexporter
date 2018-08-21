package top.jfunc.common.utils;

import com.jfinal.plugin.activerecord.Record;
import org.junit.Ignore;
import org.junit.Test;
import top.jfunc.common.ftp.ConnectBean;
import top.jfunc.common.ftp.UploadBean;
import top.jfunc.common.report.AbstractExportService;
import top.jfunc.common.report.AbstractExportUploadExcelService;
import top.jfunc.common.report.AbstractExportUploadTxtService;

import java.util.List;

/**
 * @author xiongshiyan at 2018/5/16
 */
public class ExportUploadTest {
    @Test
    @Ignore
    public void testTxt() throws Exception{
        AbstractExportService service = new AbstractExportUploadTxtService() {
            @Override
            protected String convertRecordToString(Record item) {
                return null;
            }

            @Override
            protected String linePrefix(Record record) {
                return null;
            }

            @Override
            protected ConnectBean getConnectBean() {
                return null;
            }

            @Override
            protected UploadBean getUploadBean() {
                return null;
            }

            @Override
            public List<Record> getList(int pageNumber, int pageSize) {
                return null;
            }

            @Override
            protected String getFileName() {
                return null;
            }

            @Override
            protected String getFilePath(String fileName) {
                return null;
            }
        };
        service.exportAndFinish();
    }

    @Test
    @Ignore
    public void testExcel() throws Exception{
        AbstractExportService<Record> service = new AbstractExportUploadExcelService<Record>() {
            @Override
            protected String[] getHeaders() {
                return new String[0];
            }

            @Override
            protected String[] getColumns() {
                return new String[0];
            }

            @Override
            protected ConnectBean getConnectBean() {
                return null;
            }

            @Override
            protected UploadBean getUploadBean() {
                return null;
            }

            @Override
            public List<Record> getList(int pageNumber, int pageSize) {
                return null;
            }

            @Override
            protected String getFileName() {
                return null;
            }

            @Override
            protected String getFilePath(String fileName) {
                return null;
            }
        };
        service.exportAndFinish();
    }
}
