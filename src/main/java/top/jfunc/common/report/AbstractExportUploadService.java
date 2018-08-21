package top.jfunc.common.report;

import top.jfunc.common.ftp.ConnectBean;
import top.jfunc.common.ftp.Ftp;
import top.jfunc.common.ftp.UploadBean;
import top.jfunc.common.ftp.one.FtpImpl;

import java.io.IOException;

/**
 * 导出并上传
 * @author xiongshiyan
 */
public abstract class AbstractExportUploadService<T> extends AbstractExportService<T> {
    @Override
    protected void afterExport() {
        Ftp ftp = new FtpImpl();
        try {
            ftp.upload(getConnectBean(),getUploadBean());
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * 获取上传连接属性
     * @return 上传文件连接属性
     */
    protected abstract ConnectBean getConnectBean();

    /**
     * 获取上传文件属性
     * @return 上传文件属性
     */
    protected abstract UploadBean getUploadBean();
}
