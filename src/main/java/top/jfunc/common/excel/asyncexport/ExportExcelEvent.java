package top.jfunc.common.excel.asyncexport;

import top.jfunc.common.event.core.ApplicationEvent;

/**
 * 报表导出事件。从1.7.7版本以后引入事件机制，报表导出就异步化了，通过EventKit.post(event)的方式，由ExportExcelListener接收事件
 * @author 熊诗言
 */
class ExportExcelEvent extends ApplicationEvent {
    public ExportExcelEvent(ExportExcelBean source) {
        super(source);
    }
}
