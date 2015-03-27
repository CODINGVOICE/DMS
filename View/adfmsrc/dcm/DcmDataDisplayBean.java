package dcm;

import common.ADFUtils;
import common.DmsUtils;

import common.JSFUtils;

import common.TablePagination;

import dcm.combinantion.CombinationEO;

import dcm.template.TemplateEO;

import dms.login.Person;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

import java.math.BigDecimal;

import java.sql.CallableStatement;
import java.sql.PreparedStatement;

import java.sql.ResultSet;

import java.sql.SQLException;

import java.sql.Statement;

import java.sql.Types;

import java.text.DecimalFormat;
import java.text.SimpleDateFormat;

import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import java.util.UUID;

import javax.faces.context.FacesContext;
import javax.faces.event.ActionEvent;
import javax.faces.event.FacesEvent;
import javax.faces.event.ValueChangeEvent;

import javax.faces.model.SelectItem;

import oracle.adf.share.ADFContext;
import oracle.adf.share.logging.ADFLogger;
import oracle.adf.view.rich.component.rich.RichPopup;
import oracle.adf.view.rich.component.rich.data.RichTable;
import oracle.adf.view.rich.component.rich.input.RichInputFile;
import oracle.adf.view.rich.component.rich.input.RichSelectOneChoice;

import oracle.adf.view.rich.component.rich.output.RichPanelCollection;
import oracle.adf.view.rich.context.AdfFacesContext;

import oracle.adf.view.rich.event.QueryEvent;
import oracle.adf.view.rich.model.FilterableQueryDescriptor;

import oracle.jbo.ApplicationModule;
import oracle.jbo.Key;
import oracle.jbo.Row;
import oracle.jbo.RowIterator;
import oracle.jbo.ViewCriteria;
import oracle.jbo.ViewCriteriaRow;
import oracle.jbo.ViewObject;

import oracle.jbo.jbotester.load.SimpleDateFormatter;
import oracle.jbo.server.DBTransaction;

import org.apache.commons.lang.ObjectUtils;
import org.apache.myfaces.trinidad.event.SelectionEvent;
import org.apache.myfaces.trinidad.event.SortEvent;
import org.apache.myfaces.trinidad.model.CollectionModel;
import org.apache.myfaces.trinidad.model.RowKeySet;
import org.apache.myfaces.trinidad.model.SortCriterion;
import org.apache.myfaces.trinidad.model.UploadedFile;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.hexj.excelhandler.reader.ExcelReaderUtil;

import team.epm.dcm.view.DcmCombinationViewRowImpl;
import team.epm.dcm.view.DcmTemplateColumnViewRowImpl;
import team.epm.dcm.view.DcmTemplateViewRowImpl;

public class DcmDataDisplayBean extends TablePagination{
    private CollectionModel dataModel;
    //模板信息
    private TemplateEO curTempalte;
    //模板列信息
    private List<ColumnDef> colsdef = new ArrayList<ColumnDef>();
    //组合信息
    private CombinationEO curCombiantion;
    //是否增量导入
    private boolean isIncrement = true;
    //是否是2007及以上格式
    private boolean isXlsx = true;
    //组合信息
    private List<ComHeader> templateHeader = new ArrayList<ComHeader>();
    //值集信息
    private Map valueSet=new HashMap();
    //是否可编辑
    private boolean isEditable=true;
    //当前组合是否可编辑
    private boolean curCombinationRecordEditable=true;
    //
    private String curCombiantionRecord;
    //日志
    private static ADFLogger _logger =ADFLogger.createADFLogger(DcmDataDisplayBean.class);
    //页面绑定组件
    private RichInputFile fileInput;
    private Map headerComponents = new LinkedHashMap();
    private RichPanelCollection panelaCollection;
    private Person curUser;
    private RichPopup errorWindow;
    private RichPopup dataImportWnd;
    private RichPopup dataExportWnd;
    private RichPopup templateExportWnd;
    //排序
    private List<SortCriterion> sortCriterions;
    //搜索
    private FilterableQueryDescriptor queryDescriptor=new DcmQueryDescriptor();
    private Map filters;
    //初始化
    public DcmDataDisplayBean() {
        this.curUser =(Person)ADFContext.getCurrent().getSessionScope().get("cur_user");
        this.dataModel = new DcmDataTableModel();
        this.initTemplate();
        this.initCombination();
        this.queryTemplateData();
    }

    public CollectionModel getDataModel() {
        return this.dataModel;
    }
    //值发生改变时更改数据的状态为更新
    public void valueChangeListener(ValueChangeEvent valueChangeEvent) {
        Map rowData = (Map)this.dataModel.getRowData();
        //编辑时仅能选中一行
        if (((DcmDataTableModel)this.dataModel).getSelectedRows().size() > 1) {
            String msg =DmsUtils.getMsg("dcm.msg.can_not_select_multiple_row");
            JSFUtils.addFacesInformationMessage(msg);
            return;
        }
        if (null == rowData.get("OPERATION")) {
            rowData.put("OPERATION", DcmDataTableModel.OPERATE_UPDATE);
        }
    }
    //行选中时设置当前选中行
    public void rowSelectionListener(SelectionEvent selectionEvent) {
        RichTable table = (RichTable)selectionEvent.getSource();
        RowKeySet rks = selectionEvent.getAddedSet();
        if (rks != null) {
            int setSize = rks.size();
            if (setSize == 0) {
                return;
            }
            Object rowKey = rks.iterator().next();
            table.setRowKey(rowKey);
        }
    }
    //新增数据行操作
    public void operation_new() {
        List<Map> modelData = (List<Map>)this.dataModel.getWrappedData();
        Map newRow = new HashMap();
        for (ColumnDef col : this.colsdef) {
            newRow.put(col.getDbTableCol(), null);
        }
        newRow.put("OPERATION", DcmDataTableModel.OPERATE_CREATE);
        modelData.add(0, newRow);
    }
    //删除数据行操作
    public void operation_delete() {
        List<Map> modelData = (List<Map>)this.dataModel.getWrappedData();
        RowKeySet keySet =
            ((DcmDataTableModel)this.dataModel).getSelectedRows();
        for (Object key : keySet) {
            Map rowData = (Map)this.dataModel.getRowData(key);
            //若为新增操作则直接从数据集删除数据
            if (DcmDataTableModel.OPERATE_CREATE.equals(rowData.get("OPERATION"))) {
                modelData.remove(rowData);
            } 
            //若为更新或数据未修改则直接将数据集数据标记为删除
            else if (DcmDataTableModel.OPERATE_UPDATE.equals(rowData.get("OPERATION")) ||
                       null == rowData.get("OPERATION")) {
                rowData.put("OPERATION", DcmDataTableModel.OPERATE_DELETE);
            }
            //已经为删除状态的数据无需做任何处理
        }
    }
    //数据保存操作
    public void operation_save() {
        boolean flag = true;
        String curComRecordId = this.curCombiantionRecord;
        List<Map> modelData = (List<Map>)this.dataModel.getWrappedData();
        StringBuffer sql_insert = new StringBuffer();
        StringBuffer sql_value = new StringBuffer();
        sql_insert.append("INSERT INTO \"").append(this.curTempalte.getTmpTable()).append("\"(");
        sql_insert.append("TEMPLATE_ID,COM_RECORD_ID,ORIGIN_ROWID");
        sql_insert.append(",CREATED_BY,UPDATED_BY,CREATED_AT,UPDATED_AT,SHEET_NAME,ROW_NO,EDIT_TYPE");
        sql_value.append(" VALUES('").append(this.curTempalte.getId()).append("'");
        sql_value.append(",'").append(ObjectUtils.toString(curComRecordId)).append("'");
        sql_value.append(",?,'").append(this.curUser.getId()).append("'");
        sql_value.append(",'").append(this.curUser.getId()).append("'");
        sql_value.append(",SYSDATE,SYSDATE,'NA',?,?");
        for (int i = 0; i < this.colsdef.size(); i++) {
            sql_insert.append(",COLUMN").append(i + 1);
            sql_value.append(",?");
        }
        sql_insert.append(")");
        sql_value.append(")");
        try {
            DBTransaction trans =(DBTransaction)DmsUtils.getDcmApplicationModule().getTransaction();
            this.clearTmpTableAndErrTable(curComRecordId);
            //将数据插入临时表
            PreparedStatement stat =trans.createPreparedStatement(sql_insert.toString() +sql_value.toString(), 0);
            int rowNo=1;
            for (Map rowData : modelData) {
                stat.setInt(2, rowNo);
                rowNo++;
                if (DcmDataTableModel.OPERATE_CREATE.equals(rowData.get("OPERATION"))) {
                    stat.setString(3, "NEW");
                } else if (DcmDataTableModel.OPERATE_DELETE.equals(rowData.get("OPERATION"))) {
                    stat.setString(3, "DELETE");
                } else if (DcmDataTableModel.OPERATE_UPDATE.equals(rowData.get("OPERATION"))) {
                    stat.setString(3, "UPDATE");
                }
                if (null != rowData.get("OPERATION")) {
                    for (int i = 0; i < this.colsdef.size(); i++) {
                        stat.setString(4 + i,(String)rowData.get(this.colsdef.get(i).getDbTableCol()));
                    }
                    stat.setString(1, (String)rowData.get("ROW_ID"));
                    stat.addBatch();
                }
            }   
            stat.executeBatch();
            trans.commit();
            stat.close();
            flag = this.handleData("EDIT", curComRecordId);
            if(flag){ this.queryTemplateData();}
        } catch (Exception e) {
            flag = false;
            if (e.getMessage().length() > 2048) {
                this.writeErrorMsg(e.getMessage().substring(0, 2048),curComRecordId);
            } else {
                this.writeErrorMsg(e.getMessage(),curComRecordId);
            }
            this._logger.severe(e);
        }
        if (!flag) {
            this.showErrorPop();
        }
    }
    //数据重置则直接刷新数据
    public void operation_reset() {
        this.queryTemplateData();
    }
    //数据导入操作
    public void operation_import(ActionEvent actionEvent) throws SQLException {
        this.dataImportWnd.cancel();
        String curComRecordId = this.curCombiantionRecord;
        //组合找不到
        if (this.curTempalte.getCombinationId() != null && curComRecordId == null) {
            JSFUtils.addFacesErrorMessage(DmsUtils.getMsg("dcm.inform.select_correct_combination"));
            return;
        }
        //上传文件为空
        if (null == this.fileInput.getValue()) {
            JSFUtils.addFacesErrorMessage(DmsUtils.getMsg("dcm.plz_select_import_file"));
            return;
        }
        //获取文件上传路径
        String filePath = this.uploadFile();
        this.fileInput.resetValue();
        if (null == filePath) {
            return;
        }
        //读取excel数据到数据库临时表
        if (!this.handleExcel(filePath, curComRecordId)) {
            return;
        }
        //进行数据处理（前置程序、校验和善后程序）
        if (this.handleData(this.isIncrement ? "INCREMENT" : "REPLACE",curComRecordId)) {
            String msg = DmsUtils.getMsg("dcm.inform.data_import_success");
            JSFUtils.addFacesInformationMessage(msg);
        } else {
            //若出现错误则显示错误信息提示框
            this.showErrorPop();
        }
        //刷新数据
        this.queryTemplateData();
    }
    //执行数据校验和数据转移

    private boolean handleData(String mode, String curComRecordId) {
        boolean successFlag = true;
        DBTransaction trans = (DBTransaction)DmsUtils.getDcmApplicationModule().getTransaction();
        try {
            //执行前置程序
            if (this.curTempalte.getPreProgram() != null) {
                CallableStatement prcs =trans.createCallableStatement("{CALl " + this.curTempalte.getPreProgram() +"(?,?,?,?,?)}", 0);
                prcs.setString(1, this.curTempalte.getId());
                prcs.setString(2, curComRecordId);
                prcs.setString(3, this.curUser.getId());
                prcs.setString(4, mode);
                prcs.setString(5, this.curUser.getLocale());
                prcs.execute();
                trans.commit();
                prcs.close();
            }
            //执行校验程序
            ViewObject vo =DmsUtils.getDcmApplicationModule().getDcmValidationQueryView();
            vo.setNamedWhereClauseParam("templateId", this.curTempalte.getId());
            vo.executeQuery();
            vo.reset();
            while (vo.hasNext()) {
                Row row = vo.next();
                CallableStatement cs =trans.createCallableStatement("{CALL " + row.getAttribute("Program") +"(?,?,?,?,?,?,?,?,?)}", 0);
                cs.setString(1, (String)row.getAttribute("ValidationId"));
                cs.setString(2, this.curTempalte.getId());
                cs.setString(3, curComRecordId);
                //获取校验对应的临时表列
                for (int i = 1; i <= this.colsdef.size(); i++) {
                    if (this.colsdef.get(i -1).getDbTableCol().equals((String)row.getAttribute("DbTableCol"))) {
                        cs.setString(4, "COLUMN" + i);
                        break;
                    }
                }
                cs.setString(5, (String)row.getAttribute("DbTableCol"));
                cs.setString(6, mode);
                cs.setString(7, this.curUser.getLocale());
                cs.setString(8, (String)row.getAttribute("Args"));
                cs.registerOutParameter(9, Types.VARCHAR);
                cs.execute();
                //若返回值为N则校验失败
                if ("N".equals(cs.getString(9))) {
                    successFlag = false;
                }
                trans.commit();
                cs.close();
            }
            if(successFlag){
                //若校验通过则执行数据导入
                CallableStatement cs =trans.createCallableStatement("{CALl " + this.curTempalte.getHandleProgram() + "(?,?,?,?,?)}", 0);
                cs.setString(1, this.curTempalte.getId());
                cs.setString(2, curComRecordId);
                cs.setString(3, this.curUser.getId());
                cs.setString(4, mode);
                cs.setString(5, this.curUser.getLocale());
                cs.execute();
                trans.commit();
                cs.close();
                //执行善后程序
                if (this.curTempalte.getAfterProgram() != null) {
                    CallableStatement afcs = trans.createCallableStatement("{CALl " + this.curTempalte.getAfterProgram() +"(?,?,?,?,?)}", 0);
                    afcs.setString(1, this.curTempalte.getId());
                    afcs.setString(2, curComRecordId);
                    afcs.setString(3, this.curUser.getId());
                    afcs.setString(4, mode);
                    afcs.setString(5, this.curUser.getLocale());
                    afcs.execute();
                    trans.commit();
                    afcs.close();
                }
            }
        } catch (Exception e) {
            successFlag = false;
            String msg = e.getMessage();
            if (msg.length() > 2048) {
                msg = msg.substring(0, 2048);
            }
            this.writeErrorMsg(msg, curComRecordId);
            this._logger.severe(e);
        }
        return successFlag;
    }
    //写入数据到错误表
    private void writeErrorMsg(String msg, String curComRecordId) {
        ViewObject vo = ADFUtils.findIterator("DcmErrorViewIterator").getViewObject();
        Row row = vo.createRow();
        row.setAttribute("TemplateId", this.curTempalte.getId());
        row.setAttribute("ComRecordId", curComRecordId);
        row.setAttribute("Msg", msg);
        row.setAttribute("SheetName", "NA");
        row.setAttribute("ValidationId",UUID.randomUUID().toString().replace("-", ""));
        row.setAttribute("Level", "Error");
        vo.insertRow(row);
        vo.getApplicationModule().getTransaction().commit();
    }
    //获取当前的组合
    private String getCurCombinationRecord() {
        String comRecordId = null;
        if (this.curCombiantion != null) {
            String combinationCode  = this.curCombiantion.getCode();
            StringBuffer sql = new StringBuffer();
            sql.append("SELECT ID FROM ").append("\"").append(combinationCode.toUpperCase()).append("\"");
            sql.append(" WHERE 1=1");
            for (ComHeader h : this.templateHeader) {
                sql.append(" AND ");
                sql.append("\"").append(h.getCode()).append("\"='").append(h.getValue()).append("'");
            }
            DBTransaction trans =(DBTransaction)DmsUtils.getDcmApplicationModule().getTransaction();
            Statement stat = trans.createStatement(1);
            try {
                ResultSet rs = stat.executeQuery(sql.toString());
                if (rs.next()) {
                    comRecordId = rs.getString("ID");
                }
                stat.close();
            } catch (SQLException e) {
                this._logger.severe(e);
            }     
        }
        return comRecordId;
    }
    
    //获取当前组合的文本信息用于拼接文件名
    private String getCurComRecordText() {
        String text = "";
        if (this.curCombiantion!=null) {
            for (ComHeader header : this.templateHeader) {
                    for (SelectItem item : header.getValues()) {
                        if (header.getValue().equals(item.getValue())) {
                            text += "_"+item.getLabel();
                        }
                    }
            }
        }
        return text;
    }
    //读取excel数据到临时表
    private boolean handleExcel(String fileName, String curComRecordId) throws SQLException {
        DBTransaction trans =(DBTransaction)DmsUtils.getDcmApplicationModule().getTransaction();
        String combinationRecord = ObjectUtils.toString(curComRecordId);
        //清空已有零时表数据
        this.clearTmpTableAndErrTable(curComRecordId);
        RowReader reader =new RowReader(trans, (int)this.curTempalte.getDataStartLine().getValue(), this.curTempalte.getId(),combinationRecord, this.curTempalte.getTmpTable(),
                          this.colsdef.size(), this.curUser.getId(),this.curTempalte.getName());
        try {
            ExcelReaderUtil.readExcel(reader, fileName, true);
            reader.close();
        } catch (Exception e) {
            this._logger.severe(e);
            JSFUtils.addFacesErrorMessage(DmsUtils.getMsg("dcm.excel_handle_error"));
            return false;
        }
        return true;
    }
    //文件上传

    private String uploadFile() {
        UploadedFile file = (UploadedFile)this.fileInput.getValue();
        if (!(file.getFilename().endsWith(".xls") ||
              file.getFilename().endsWith(".xlsx"))) {
            String msg = DmsUtils.getMsg("dcm.file_format_error");
            JSFUtils.addFacesErrorMessage(msg);
            return null;
        }
        String fileExtension =file.getFilename().substring(file.getFilename().lastIndexOf("."));
        SimpleDateFormat dateFormat = new SimpleDateFormat("yyyyMMddHHmmss");
        String date = dateFormat.format(new Date());
        File dmsBaseDir = new File("DMS/UPLOAD/" + this.curTempalte.getName());
        //如若文件路径不存在则创建文件目录
        if (!dmsBaseDir.exists()) {
            dmsBaseDir.mkdirs();
        }
        String fileName = "DMS/UPLOAD/" + this.curTempalte.getName() + "/";
        if (this.curCombiantion == null) {
            fileName +=this.curTempalte.getName() + "_" + this.curUser.getName() + "_" +date + fileExtension;
        } else {
            fileName +=this.getCurComRecordText() + "_" + this.curUser.getName() +"_" + date + fileExtension;
        }
        try {
            InputStream inputStream = file.getInputStream();
            FileOutputStream outputStream = new FileOutputStream(fileName);
            BufferedInputStream bufferedInputStream =new BufferedInputStream(inputStream);
            BufferedOutputStream bufferedOutputStream =
                new BufferedOutputStream(outputStream);
            byte[] buffer = new byte[10240];
            int bytesRead = 0;
            while ((bytesRead = bufferedInputStream.read(buffer, 0, 10240)) !=-1) {
                bufferedOutputStream.write(buffer, 0, bytesRead);
            }
            bufferedOutputStream.flush();
            if (bufferedOutputStream != null) {
                bufferedOutputStream.close();
            }
            if (bufferedInputStream != null) {
                bufferedInputStream.close();
            }
            file.dispose();
        } catch (IOException e) {
            this._logger.severe(e);
            String msg = DmsUtils.getMsg("dcm.file_upload_error");
            JSFUtils.addFacesErrorMessage(msg);
            return null;
        }
        return (new File(fileName)).getAbsolutePath();
    }
    //数据导出
    public void operation_export(FacesContext facesContext,java.io.OutputStream outputStream) {
        this.dataExportWnd.cancel();
        String type = this.isXlsx ? "xlsx" : "xls";
        try {
            if ("xls".equals(type)) {
                Excel2003WriterImpl writer=new Excel2003WriterImpl(
                                               this.getQuerySql(),
                                               this.curTempalte,
                                               this.colsdef,
                                               outputStream);
                writer.writoToFile();
            } else {
                Excel2007WriterImpl writer=new Excel2007WriterImpl(this.getQuerySql(),
                                                                   (int)this.curTempalte.getDataStartLine().getValue(),
                                                                   this.colsdef);
                writer.process(outputStream, this.curTempalte.getName());
            }
            outputStream.flush();
        } catch (Exception e) {
            this._logger.severe(e);
        }
    }
    //下载模板
    public void operation_download(FacesContext facesContext,java.io.OutputStream outputStream) {
        this.templateExportWnd.cancel();
        try {
            if (this.isXlsx) {
                if (this.curTempalte.getTemplateFile() == null) {
                    XSSFWorkbook wb = new XSSFWorkbook();
                    XSSFSheet sheet = wb.createSheet(this.curTempalte.getName());
                    XSSFRow row = sheet.createRow((int)this.curTempalte.getDataStartLine().getValue() - 2);
                    for (int i = 0; i < this.colsdef.size(); i++) {
                        row.createCell(i).setCellValue(this.colsdef.get(i).getColumnLabel());
                    }
                    wb.write(outputStream);
                    outputStream.flush();
                } else {
                    FileInputStream inputStream =
                        new FileInputStream(this.curTempalte.getTemplateFile() + ".xlsx");
                    byte[] buffer = new byte[10240];
                    int bytesRead = 0;
                    while ((bytesRead = inputStream.read(buffer, 0, 10240)) !=
                           -1) {
                        outputStream.write(buffer, 0, bytesRead);
                    }
                    outputStream.flush();
                }
            } else {
                if (this.curTempalte.getTemplateFile() == null) {
                    // 创建excel2003对象
                    Workbook wb = new HSSFWorkbook();
                    // 创建新的表单
                    Sheet sheet = wb.createSheet(this.curTempalte.getName());
                    // 创建新行
                    org.apache.poi.ss.usermodel.Row headerRow =
                        sheet.createRow((int)this.curTempalte.getDataStartLine().getValue() - 2);
                    for (int i = 0; i < this.colsdef.size(); i++) {
                        headerRow.createCell(i).setCellValue(this.colsdef.get(i).getColumnLabel());
                    }
                    wb.write(outputStream);
                    outputStream.flush();
                } else {
                    FileInputStream inputStream =
                        new FileInputStream(this.curTempalte.getTemplateFile() + ".xls");
                    byte[] buffer = new byte[10240];
                    int bytesRead = 0;
                    while ((bytesRead = inputStream.read(buffer, 0, 10240)) !=
                           -1) {
                        outputStream.write(buffer, 0, bytesRead);
                    }
                    outputStream.flush();
                }
            }
        } catch (Exception e) {
            this._logger.severe(e);
            String msg = DmsUtils.getMsg("dcm.template_download_error");
            JSFUtils.addFacesErrorMessage(msg);
            return;
        }
    }
    //初始化模板和模板列信息

    private void initTemplate() {
        String curTemplateId = (String)ADFContext.getCurrent().getPageFlowScope().get("curTemplateId");
        ViewObject templateView =DmsUtils.getDcmApplicationModule().getDcmTemplateView();
        templateView.executeQuery();
        Row[] rows=templateView.findByKey(new Key(new Object[]{curTemplateId,ADFContext.getCurrent().getLocale().toString()}), 1);
        if(rows.length>0){
            this.curTempalte=new TemplateEO((DcmTemplateViewRowImpl)rows[0]); 
            //只读模板
            if("Y".equals(this.curTempalte.getReadonly())){
                this.isEditable=false;
            }
            //对该模板仅有只读权限
            if(this.isEditable){
                this.isEditable=!this.isReadOnly();
            }
            //不支持增量导入则设置默认覆盖导入
            if(!this.curTempalte.getHandleMode().contains("I")){
               this.isIncrement=false; 
            }
            templateView.setCurrentRow(rows[0]);
            DcmTemplateViewRowImpl row=(DcmTemplateViewRowImpl)rows[0];
            RowIterator itr=row.getDcmTemplateColumnView();
            while(itr.hasNext()){
                Row colRow = itr.next();
                ColumnDef colDef = new ColumnDef((DcmTemplateColumnViewRowImpl)colRow);
                this.colsdef.add(colDef);
                //获取值列表
                if(colDef.getValueSetId()!=null&&this.valueSet.get(colDef.getValueSetId())==null){
                    this.valueSet.put(colDef.getValueSetId(), 
                                      this.fetchValueList(colDef.getValueSetId()));            
                }
            }
            ((DcmDataTableModel)this.dataModel).setColsdef(this.colsdef);
        }else{
            this._logger.severe(DmsUtils.getMsg("dcm.template_not_found"));
            throw new RuntimeException(DmsUtils.getMsg("dcm.template_not_found")+":tempateId:"+curTemplateId);
        }
    }
    //获取值列表
    private List<SelectItem> fetchValueList(String vsId){
        List<SelectItem> list=new ArrayList<SelectItem>();
        list.add(new SelectItem());
        Row[] vsRows=DmsUtils.getDmsApplicationModule().getDmsValueSetView()
            .findByKey(new Key(new Object[]{vsId,ADFContext.getCurrent().getLocale().toString()}), 1);
        if(vsRows.length>0){
            String vsCode=(String)vsRows[0].getAttribute("Source");
            StringBuffer sql=new StringBuffer();
            sql.append("SELECT T.CODE, T.MEANING FROM \"").append(vsCode)
            .append("\" T WHERE T.LOCALE = '").append(ADFContext.getCurrent().getLocale()).append("'  ORDER BY T.IDX ");
            Statement stmt= DmsUtils.getDmsApplicationModule().getDBTransaction().createStatement(DBTransaction.DEFAULT);
            try {
                ResultSet rs = stmt.executeQuery(sql.toString());
                while(rs.next()){
                    SelectItem itm=new SelectItem();
                    itm.setLabel(rs.getString("MEANING"));
                    itm.setValue(rs.getString("MEANING"));
                    list.add(itm);
                }
            } catch (SQLException e) {
                this._logger.severe(e);
            }
        }
        return list;
    }
    //查询数据
    private void queryTemplateData() {
        List<Map> data = new ArrayList<Map>();
        DBTransaction dbTransaction =(DBTransaction)DmsUtils.getDcmApplicationModule().getTransaction();
        String countSql=this.getCountSql(this.getQuerySql());
        PreparedStatement cstat =dbTransaction.createPreparedStatement(countSql, -1);
        try {
            ResultSet crs = cstat.executeQuery();
            if(crs.next()){
                this.setTotalCount(crs.getInt(1));
            }
            crs.close();
            cstat.close();
        } catch (SQLException e) {
            this._logger.severe(e);
        }
        String sql =this.getPaginationSql(this.getQuerySql());
        PreparedStatement stat =dbTransaction.createPreparedStatement(sql, -1);
        ResultSet rs = null;
        try {
            rs = stat.executeQuery();
            while (rs.next()) {
                Map row = new HashMap();
                for (ColumnDef col : this.colsdef) {
                    Object obj=rs.getObject(col.getDbTableCol().toUpperCase());
                    if(obj instanceof java.sql.Date){
                        SimpleDateFormatter format=new SimpleDateFormatter("yyyy-MM-dd hh:mm:ss");
                        obj=format.format((java.sql.Date)obj);
                    }else if(obj instanceof BigDecimal){
                        obj=((BigDecimal)obj).doubleValue();
                        DecimalFormat fmt=new DecimalFormat("#.################");
                        obj=fmt.format((Double)obj);
                    }
                    else{
                        obj=ObjectUtils.toString(obj);
                    }
                    row.put(col.getDbTableCol(),obj);
                }
                row.put("ROW_ID", rs.getString("ROW_ID"));
                data.add(row);
            }
            rs.close();
        } catch (SQLException e) {
            JSFUtils.addFacesErrorMessage(DmsUtils.getMsg("dcm.query_data_error"));
            this._logger.severe(e);
        }
        this.dataModel.setWrappedData(data);
    }
    //获取数据查询语句
    private String getQuerySql() {
        StringBuffer sql_select = new StringBuffer();
        StringBuffer sql_from = new StringBuffer();
        StringBuffer sql_where = new StringBuffer();
        if (null != this.curTempalte.getDbView()) {
            sql_select.append("SELECT ROW_ID");
            sql_from.append(" ").append("FROM ").append(this.curTempalte.getDbView());
        } else {
            sql_select.append("SELECT ROWID ROW_ID");
            sql_from.append(" ").append("FROM ").append(this.curTempalte.getDbTable());
        }
        for (ColumnDef col : this.colsdef) {
            sql_select.append(",\"").append(col.getDbTableCol()).append("\"");
        }
        if (this.curCombiantion != null) {
            sql_where.append(" WHERE COM_RECORD_ID='").append(this.curCombiantionRecord).append("'");
        } else {
            sql_where.append(" WHERE COM_RECORD_ID IS NULL");
        }
        if(this.filters!=null){
            for(Object key:this.filters.keySet()){
                if(!ObjectUtils.toString(this.filters.get(key)).trim().equals("")){
                    String fv=ObjectUtils.toString(this.filters.get(key));
                    if("NULL".equals(fv.toUpperCase())){
                        sql_where.append(" AND \"").append(key).append("\" IS NULL");
                    }else{
                        sql_where.append(" AND UPPER(\"").append(key).append("\") LIKE UPPER('");
                        if(fv.startsWith("%")){
                            sql_where.append("%");
                            fv=fv.substring(1);
                        }
                        if(fv.endsWith("%")){
                            fv=fv.substring(0,fv.lastIndexOf("%"));
                            sql_where.append(fv);
                            sql_where.append("%");
                        }else{
                            sql_where.append(fv);
                        } 
                        sql_where.append("')");
                    }
                }
            }
        }
        if(this.sortCriterions==null){
            sql_where.append(" ORDER BY IDX");
        }else{
            sql_where.append(" ORDER BY ");
            for(int i=0;i<this.sortCriterions.size();i++){
                SortCriterion s=this.sortCriterions.get(i);
                if(i>0){
                    sql_where.append(",");
                }
                sql_where.append("\"").append(s.getProperty()).append("\"").append(s.isAscending() ? " ASC":" DESC");
            }
        }
        return sql_select.toString() + sql_from.toString() +
            sql_where.toString();
    }
    
    //初始化组合信息
    private void initCombination() {
        if (null != this.curTempalte.getCombinationId()) {
            //初始化组合基本信息
            ViewObject cVo =DmsUtils.getDcmApplicationModule().getDcmCombinationView();
            Row rows[]=cVo.findByKey(new Key(new Object[]{this.curTempalte.getCombinationId(),ADFContext.getCurrent().getLocale().toString()}), 1);
            if (rows.length>0) {
                this.curCombiantion = new CombinationEO((DcmCombinationViewRowImpl)rows[0]);;
            }
            //初始化组合头信息
            ViewObject vo =DmsUtils.getDcmApplicationModule().getDcmComVsQueryView();
            vo.setNamedWhereClauseParam("combinationId", this.curTempalte.getCombinationId());
            vo.executeQuery();
            vo.reset();
            while (vo.hasNext()) {
                Row row = vo.next();
                ComHeader header = new ComHeader();
                header.setName((String)row.getAttribute("Name"));
                header.setIsAuthority((String)row.getAttribute("IsAuthority"));
                header.setSrcTable((String)row.getAttribute("Source"));
                header.setValueSetId((String)row.getAttribute("ValueSetId"));
                header.setCode((String)row.getAttribute("Code"));
                this.initHeaderValueList(header);
                this.setDefaultHeaderValue(header);
                this.templateHeader.add(header);
            }
            this.curCombiantionRecord=this.getCurCombinationRecord();
            this.setCurCombinationRecordEditable();
        }
    }
    private void setCurCombinationRecordEditable(){
        if(this.curCombiantionRecord==null){
            this.curCombinationRecordEditable=false;
        }else{
            boolean flag=true;
            StringBuffer sql=new StringBuffer();
            sql.append("SELECT T.STATUS FROM DCM_TEMPLATE_COMBINATION T WHERE T.TEMPLATE_ID='")
                .append(this.curTempalte.getId()).append("' AND T.COM_RECORD_ID='")
                .append(this.curCombiantionRecord).append("'");
            DBTransaction dbTransaction =(DBTransaction)DmsUtils.getDcmApplicationModule().getTransaction();
            Statement stat=dbTransaction.createStatement(DBTransaction.DEFAULT);
            try {
                ResultSet rs=stat.executeQuery(sql.toString());
                if(rs.next()){
                    String status=rs.getString("STATUS");
                    if("OPEN".equals(status)){
                        for(ComHeader h:this.templateHeader){
                            StringBuffer sql0=new StringBuffer();
                            sql0.append("SELECT T.ENABLED FROM \"").append(h.getSrcTable())
                                .append("\" T WHERE T.CODE='")
                                .append(h.getValue()).append("' AND T.LOCALE='")
                                .append(ADFContext.getCurrent().getLocale()).append("'");
                            ResultSet rst=stat.executeQuery(sql0.toString());
                            if(rst.next()){
                                String enabled=rst.getString("ENABLED");
                                if("N".equals(enabled)){
                                    flag=false;
                                    break;
                                }
                            }else{
                                flag=false;
                                break;
                            }
                            rst.close();
                        }
                    }else{
                        flag=false;
                    }         
                }else{
                    flag=false;
                }
                rs.close();
                stat.close();
            } catch (SQLException e) {
                this._logger.severe(e);
            }
            this.curCombinationRecordEditable=flag;
        }
    }
    private void initHeaderValueList(ComHeader header){
        List<SelectItem> values = new ArrayList<SelectItem>();
        DBTransaction dbTransaction =(DBTransaction)DmsUtils.getDcmApplicationModule().getTransaction();
        StringBuffer sql = new StringBuffer();
        sql.append("SELECT V.CODE,V.MEANING FROM \"");
        sql.append(header.getSrcTable()).append("\" V");
        sql.append(" WHERE V.LOCALE='").append(this.curUser.getLocale()).append("'");
        if ("Y".equals(header.getIsAuthority())) {
            sql.append(" AND EXISTS(SELECT 1 FROM ");
            sql.append(" DMS_USER_VALUE_V T");
            sql.append(" WHERE T.USER_ID = '").append(this.curUser.getId()).append("'");
            sql.append(" AND T.VALUE_SET_ID = '").append(header.getValueSetId()).append("'");
            sql.append(" AND T.VALUE_ID=V.CODE)");
        }
        sql.append(" ORDER BY V.IDX");
        PreparedStatement stat =
            dbTransaction.createPreparedStatement(sql.toString(), -1);
        try {
            ResultSet rs = stat.executeQuery();
            while (rs.next()) {
                SelectItem item = new SelectItem();
                item.setLabel(rs.getString("MEANING"));
                item.setValue(rs.getString("CODE"));
                values.add(item);
            }
        } catch (SQLException e) {
            this._logger.severe(e);
        }   
        header.setValues(values);
    }
    private void setDefaultHeaderValue(ComHeader header){
        DBTransaction dbTransaction =(DBTransaction)DmsUtils.getDcmApplicationModule().getTransaction();
        StringBuffer sql = new StringBuffer();
        sql.append("SELECT V.CODE,V.MEANING FROM \"");
        sql.append(header.getSrcTable()).append("\" V");
        sql.append(" WHERE V.LOCALE='").append(this.curUser.getLocale()).append("'");
        sql.append(" AND EXISTS(SELECT 1 FROM \"").append(this.curCombiantion.getCode()).append("\" T,DCM_TEMPLATE_COMBINATION A ");
        sql.append(" WHERE T.\"").append(header.getCode()).append("\"=V.CODE");
        sql.append(" AND T.ID=A.COM_RECORD_ID");
        sql.append(" AND A.TEMPLATE_ID='").append(this.curTempalte.getId()).append("'");
        sql.append(" AND A.STATUS='OPEN'");
        for (ComHeader h : this.templateHeader) {
            if (!h.equals(header)) {
                sql.append("AND T.\"").append(h.getCode()).append("\"='");
                sql.append(h.getValue()).append("'");
            }else{
                break;
            }
        }
        sql.append(")");
        if ("Y".equals(header.getIsAuthority())) {
            sql.append(" AND EXISTS(SELECT 1 FROM ");
            sql.append(" DMS_USER_VALUE_V T");
            sql.append(" WHERE T.USER_ID = '").append(this.curUser.getId()).append("'");
            sql.append(" AND T.VALUE_SET_ID = '").append(header.getValueSetId()).append("'");
            sql.append(" AND T.VALUE_ID=V.CODE)");
        }
        sql.append(" ORDER BY V.IDX");
        PreparedStatement stat =
            dbTransaction.createPreparedStatement(sql.toString(), -1);
        try {
            ResultSet rs = stat.executeQuery();
            if (rs.next()) {
                header.setValue(rs.getString("CODE"));
            }
        } catch (SQLException e) {
            this._logger.severe(e);
        }    
    }

    public String getCombinationId() {
        return curTempalte.getCombinationId();
    }

    public void setTemplateHeader(List<ComHeader> templateHeader) {
        this.templateHeader = templateHeader;
    }

    public List<ComHeader> getTemplateHeader() {
        return templateHeader;
    }

    public void setFileInput(RichInputFile fileInput) {
        this.fileInput = fileInput;
    }

    public RichInputFile getFileInput() {
        return fileInput;
    }
    //选择不同的组合时的处理逻辑
    public void headerSelectChangeListener(ValueChangeEvent valueChangeEvent) {
        RichSelectOneChoice header =
            (RichSelectOneChoice)valueChangeEvent.getSource();
        int i = 0;
        for (Object key : this.headerComponents.keySet()) {
            //找到当前表头
            if (header.equals(this.headerComponents.get(key))) {
                this.templateHeader.get(i).setValue((String)valueChangeEvent.getNewValue());
            }
            i++;
        }
        this.curCombiantionRecord=this.getCurCombinationRecord();
        this.setCurCombinationRecordEditable();
        this.queryTemplateData();
        AdfFacesContext adfFacesContext = AdfFacesContext.getCurrentInstance();
        adfFacesContext.addPartialTarget(this.panelaCollection);
    }

    public Map getHeaderComponents() {
        return headerComponents;
    }

    public void setPanelaCollection(RichPanelCollection panelaCollection) {
        this.panelaCollection = panelaCollection;
    }

    public RichPanelCollection getPanelaCollection() {
        return panelaCollection;
    }

    public String getTemplateName() {
        return this.curTempalte.getName();
    }

    public void setIsIncrement(boolean isIncrement) {
        this.isIncrement = isIncrement;
    }

    public boolean isIsIncrement() {
        return isIncrement;
    }

    public void setIsXlsx(boolean isXlsx) {
        this.isXlsx = isXlsx;
    }

    public boolean isIsXlsx() {
        return isXlsx;
    }
    //获取导出数据时的文件名

    public String getExportDataExcelName() {
        if (this.isXlsx) {
            return this.curTempalte.getName() + this.getCurComRecordText() + ".xlsx";
        } else {
            return this.curTempalte.getName() + this.getCurComRecordText() + ".xls";
        }
    }
    //获取模板文件名

    public String getExportTemplateExcelName() {
        if (this.isXlsx) {
            return this.curTempalte.getName() + ".xlsx";
        } else {
            return this.curTempalte.getName() + ".xls";
        }
    }

    public void showErrors(ActionEvent actionEvent) {
        this.showErrorPop();
    }
    //显示错误信息窗口
    private void showErrorPop() {
        ViewObject vo =ADFUtils.findIterator("DcmErrorViewIterator").getViewObject();
        ViewCriteria viewCriteria = vo.createViewCriteria();
        ViewCriteriaRow vr = viewCriteria.createViewCriteriaRow();
        if (this.curCombiantion == null) {
            vr.setAttribute("ComRecordId", " is null");
        } else {
            vr.setAttribute("ComRecordId", "='" + this.curCombiantionRecord + "'");
        }
        vr.setAttribute("TemplateId", "='" + this.curTempalte.getId() + "'");
        viewCriteria.addRow(vr);
        vo.applyViewCriteria(viewCriteria);
        vo.executeQuery();
        vo.getViewCriteriaManager().setApplyViewCriteriaNames(null);
        RichPopup.PopupHints ph = new RichPopup.PopupHints();
        this.errorWindow.show(ph);
    }

    public void setErrorWindow(RichPopup errorWindow) {
        this.errorWindow = errorWindow;
    }

    public RichPopup getErrorWindow() {
        return errorWindow;
    }
    private void clearTmpTableAndErrTable(String comRecordId) throws SQLException {
        String clearTmpTableSql="DELETE FROM \"" +this.curTempalte.getTmpTable() 
                                +"\" WHERE TEMPLATE_ID='" +
                                this.curTempalte.getId() +"' AND COM_RECORD_ID " 
                                +(comRecordId == null ?
                                "IS NULL" :("='" + comRecordId + "'"));
        String clearErrTableSql="DELETE FROM DCM_ERROR WHERE TEMPLATE_ID='" 
                                +this.curTempalte.getId() +"' AND COM_RECORD_ID " 
                                +(comRecordId == null ?
                                "IS NULL" :("='" + comRecordId + "'"));
        ApplicationModule dcmApplicationModule =DmsUtils.getDcmApplicationModule();
        //清空临时表相应数据
        dcmApplicationModule.getTransaction().executeCommand(clearTmpTableSql);
        //清空错误表相应数据
        dcmApplicationModule.getTransaction().executeCommand(clearErrTableSql);
        dcmApplicationModule.getTransaction().commit();
    }

    public void sortListener(SortEvent sortEvent) {
        this.sortCriterions=sortEvent.getSortCriteria();
        this.queryTemplateData();
    }
    public FilterableQueryDescriptor getQueryDescriptor(){
        return this.queryDescriptor;
    }

    public void queryListener(QueryEvent queryEvent) {
        DcmQueryDescriptor descriptor = (DcmQueryDescriptor)queryEvent.getDescriptor();
        this.filters=descriptor.getFilterCriteria();
        this.setCurPage(1);
        this.queryTemplateData();
    }

    public Map getValueSet() {
        return valueSet;
    }

    public void nextPage(ActionEvent event) {
        this.setCurPage(this.getCurPage()+1);
        this.queryTemplateData();
    }

    public void prePage(ActionEvent event) {
        this.setCurPage(this.getCurPage()-1);
        this.queryTemplateData();
    }

    public void refreshPage(ActionEvent event) {
    }

    public void gotoPage(ValueChangeEvent event) {
        this.setCurPage((Integer)event.getNewValue());
        this.queryTemplateData();
        AdfFacesContext.getCurrentInstance().addPartialTarget(this.panelaCollection);
    }

    public void firstPage(ActionEvent event) {
        this.setCurPage(1);
        this.queryTemplateData();
    }

    public void lastPage(ActionEvent event) {
        this.setCurPage(this.getTotalPage());
        this.queryTemplateData();
    }

    public void setDataImportWnd(RichPopup dataImportWnd) {
        this.dataImportWnd = dataImportWnd;
    }

    public RichPopup getDataImportWnd() {
        return dataImportWnd;
    }

    public void setDataExportWnd(RichPopup dataExportWnd) {
        this.dataExportWnd = dataExportWnd;
    }

    public RichPopup getDataExportWnd() {
        return dataExportWnd;
    }

    public void setTemplateExportWnd(RichPopup templateExportWnd) {
        this.templateExportWnd = templateExportWnd;
    }

    public RichPopup getTemplateExportWnd() {
        return templateExportWnd;
    }
    public boolean getIsEditable(){
        if(this.curCombiantion==null){
            return this.isEditable;
        }else{
            return this.isEditable&&this.curCombiantionRecord!=null&&this.curCombinationRecordEditable;
        }
    }
    private boolean isReadOnly(){
        boolean readonly=false;
        StringBuffer sql=new StringBuffer();
        sql.append("SELECT 1 ")
           .append("  FROM DCM_USER_TEMPLATE_V T ")
           .append(" WHERE T.READ_ONLY='Y' AND T.USER_ID = '").append(this.curUser.getId()).append("' ")
           .append("   AND T.TEMPLATE_ID = '").append(this.curTempalte.getId()).append("'");
        Statement stmt=DmsUtils.getDcmApplicationModule().getDBTransaction().createStatement(DBTransaction.DEFAULT);
        try {
            ResultSet rs = stmt.executeQuery(sql.toString());
            if(rs.next()){
                readonly=true;
            }
            rs.close();
            stmt.close();
        } catch (SQLException e) {
            this._logger.severe(e);
        }
        return readonly;
    }
    public boolean getCanIncrementImport(){
        return this.curTempalte.getHandleMode().contains("I") ? true:false;
    }
    public boolean getCanReplaceImport(){
        return this.curTempalte.getHandleMode().contains("R") ? true:false;
    }
    public boolean getIsReplaceDefault(){
        return !this.getCanIncrementImport();
    }
    public void setIsReplaceDefault(boolean flag){
       
    }
}
