package dcm;

import common.DmsUtils;

import java.math.BigDecimal;

import java.sql.PreparedStatement;
import java.sql.ResultSet;

import java.util.List;

import oracle.adf.share.logging.ADFLogger;

import oracle.jbo.jbotester.load.SimpleDateFormatter;
import oracle.jbo.server.DBTransaction;

import org.apache.commons.lang.ObjectUtils;

import org.hexj.excelhandler.writer.AbstractExcel2007Writer;

public class Excel2007WriterImpl extends AbstractExcel2007Writer {
    private static ADFLogger logger =
        ADFLogger.createADFLogger(Excel2007WriterImpl.class);
    private String sql;
    private int dataStartLine;
    private List<ColumnDef> colsdef;

    public Excel2007WriterImpl(String sql, int dataStartLine,
                               List<ColumnDef> colsdef) {
        this.sql = sql;
        this.dataStartLine = dataStartLine;
        this.colsdef = colsdef;
    }

    public void generate() {
        try {
            DBTransaction dbTransaction =
                (DBTransaction)DmsUtils.getDcmApplicationModule().getTransaction();
            PreparedStatement stat =
                dbTransaction.createPreparedStatement(sql, -1);
            ResultSet rs = stat.executeQuery();
            //电子表格开始
            beginSheet();
            insertRow(dataStartLine - 2);
            for (int i = 0; i < this.colsdef.size(); i++) {
                createCell(i, this.colsdef.get(i).getColumnLabel());
            }
            endRow();
            int n = dataStartLine - 1;
            while (rs.next()) {
                int colInx = 0;
                insertRow(n);
                for (ColumnDef col : this.colsdef) {
                    Object obj = rs.getObject(col.getDbTableCol());
                    if (obj instanceof java.sql.Date) {
                        SimpleDateFormatter format =
                            new SimpleDateFormatter("yyyy-MM-dd hh:mm:ss");
                        obj = format.format((java.sql.Date)obj);
                    } else if (obj instanceof BigDecimal) {
                        obj = ((BigDecimal)obj).doubleValue();
                        obj = ObjectUtils.toString(obj);
                    } else {
                        obj = ObjectUtils.toString(obj);
                    }
                    createCell(colInx, (String)obj);
                    ++colInx;
                }
                ++n;
                //结束行
                endRow();
            }
            rs.close();
            //电子表格结束
            endSheet();
        } catch (Exception e) {
            this.logger.severe(e);
        }
    }
}
