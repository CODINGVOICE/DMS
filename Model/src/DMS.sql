/*==============================================================*/
/* DBMS name:      ORACLE Version 11g                           */
/* Created on:     2014/7/31 9:48:05                            */
/*==============================================================*/


drop table DCM_COMBINATION cascade constraints;

drop table DCM_COM_VS cascade constraints;

drop table DCM_ERROR cascade constraints;

drop table DCM_JOB cascade constraints;

drop table DCM_ROLE_TEMPLATE cascade constraints;

drop table DCM_TEMPLATE_CAT cascade constraints;

drop table DCM_TEMPLATE cascade constraints;

drop table DCM_TEMPLATE_COLUMN cascade constraints;

drop table DCM_TEMPLATE_COMBINATION cascade constraints;

drop table DCM_TEMPLATE_VALIDATION cascade constraints;

drop table DCM_VALIDATION cascade constraints;

drop table DMS_AUDIT_MSG cascade constraints;

drop table DMS_FUNCTION cascade constraints;

drop table DMS_GROUP cascade constraints;

drop table DMS_GROUP_ROLE cascade constraints;

drop table DMS_MENU cascade constraints;

drop table DMS_PROPERTY cascade constraints;

drop table DMS_ROLE cascade constraints;

drop table DMS_ROLE_FUNCTION cascade constraints;

drop table DMS_ROLE_MENU cascade constraints;

drop table DMS_GROUP_VALUE cascade constraints;

drop table DMS_USER cascade constraints;

drop table DMS_USER_GROUP cascade constraints;

drop table DMS_USER_KEY cascade constraints;

drop table DMS_VALUE_SET cascade constraints;

drop table DMS_LOOKUP cascade constraints;

drop table DCM_TEMPTABLE10 cascade constraints;

drop table DCM_TEMPTABLE20 cascade constraints;

drop table DCM_TEMPTABLE30 cascade constraints;

drop table DCM_TEMPTABLE40 cascade constraints;

drop table DCM_TEMPTABLE50 cascade constraints;

drop table DCM_TEMPTABLE80 cascade constraints;

drop table DCM_TEMPTABLE100 cascade constraints;

drop sequence DCM_SEQ;
drop VIEW DMS_USER_VALUE_V;
drop VIEW DMS_USER_FUNCTION_V;
drop VIEW DCM_USER_TEMPLATE_V;

/*==============================================================*/
/* Sequence: DCM_SEQ                                            */
/*==============================================================*/
create sequence DCM_SEQ
minvalue 10000
maxvalue 99999999999999999999999999999999
start with 10000
increment by 1
cache 20;
/*==============================================================*/
/* Table: DCM_COMBINATION                                       */
/*==============================================================*/
create table DCM_COMBINATION 
(
   ID                   VARCHAR2(32)         not null,
   NAME          VARCHAR2(100)        not null,
   CODE              VARCHAR2(26)     not null,
   LOCALE        VARCHAR2(10)         not null,
   CREATED_AT           DATE,
   UPDATED_AT           DATE,
   UPDATED_BY           VARCHAR2(32),
   CREATED_BY           VARCHAR2(32),
   constraint PK_DCM_COMBINATION primary key (ID,LOCALE)
);

comment on table DCM_COMBINATION is
'模版组合';

comment on column DCM_COMBINATION.ID is
'标识ID';

comment on column DCM_COMBINATION.NAME is
'组合名称';

comment on column DCM_COMBINATION.CODE is
'组合编码';

comment on column DCM_COMBINATION.LOCALE is
'语言编码';

comment on column DCM_COMBINATION.CREATED_AT is
'创建时间';

comment on column DCM_COMBINATION.UPDATED_AT is
'更新时间';

comment on column DCM_COMBINATION.UPDATED_BY is
'更新者';

comment on column DCM_COMBINATION.CREATED_BY is
'创建者';


/*==============================================================*/
/* Table: DCM_COM_VS                                       */
/*==============================================================*/
create table DCM_COM_VS 
(
   COMBINATION_ID       VARCHAR2(32)         not null,
   VALUE_SET_ID         VARCHAR2(32)     NOT NULL,
   SEQ          NUMBER,
   IS_AUTHORITY         VARCHAR2(10),
   CREATED_AT           DATE,
   UPDATED_AT           DATE,
   UPDATED_BY           VARCHAR2(32),
   CREATED_BY           VARCHAR2(32),
   constraint PK_DCM_COM_VS primary key (COMBINATION_ID,VALUE_SET_ID)
);

comment on table DCM_COM_VS is
'组合值集对应关系';

comment on column DCM_COM_VS.COMBINATION_ID is
'组合ID';

comment on column DCM_COM_VS.VALUE_SET_ID is
'值集ID';

comment on column DCM_COM_VS.SEQ is
'序号';

comment on column DCM_COM_VS.IS_AUTHORITY is
'是否作权限控制';


comment on column DCM_COM_VS.CREATED_AT is
'创建时间';

comment on column DCM_COM_VS.UPDATED_AT is
'更新时间';

comment on column DCM_COM_VS.UPDATED_BY is
'更新者';

comment on column DCM_COM_VS.CREATED_BY is
'创建者';

/*==============================================================*/
/* Table: DCM_TEMPLATE_CAT                                       */
/*==============================================================*/
create table DCM_TEMPLATE_CAT
(
   ID                   VARCHAR2(32)         not null,
   NAME                 VARCHAR2(100)        not null,
   LOCALE               VARCHAR2(10)         not null,
   P_ID                 VARCHAR2(32),
   CREATED_AT           DATE,
   UPDATED_AT           DATE,
   UPDATED_BY           VARCHAR2(32),
   CREATED_BY           VARCHAR2(32),
   constraint PK_DCM_TEMPLATE_CAT primary key (ID,LOCALE)
);

comment on table DCM_TEMPLATE_CAT is
'模版分类';

comment on column DCM_TEMPLATE_CAT.ID is
'标识ID';

comment on column DCM_TEMPLATE_CAT.NAME is
'分类名称';

comment on column DCM_TEMPLATE_CAT.P_ID is
'父标识ID';

comment on column DCM_TEMPLATE_CAT.LOCALE is
'语言编码';

comment on column DCM_TEMPLATE_CAT.CREATED_AT is
'创建时间';

comment on column DCM_TEMPLATE_CAT.UPDATED_AT is
'更新时间';

comment on column DCM_TEMPLATE_CAT.UPDATED_BY is
'更新者';

comment on column DCM_TEMPLATE_CAT.CREATED_BY is
'创建者';

/*==============================================================*/
/* Table: DCM_ROLE_TEMPLATE                                     */
/*==============================================================*/
create table DCM_ROLE_TEMPLATE 
(
   ID                   VARCHAR2(32)         not null,
   ROLE_ID              VARCHAR2(32),
   TEMPLATE_ID          VARCHAR2(32),
   CREATED_AT           DATE,
   UPDATED_AT           DATE,
   UPDATED_BY           VARCHAR2(32),
   CREATED_BY           VARCHAR2(32),
   READ_ONLY            VARCHAR2(10),
   constraint PK_DCM_ROLE_TEMPLATE primary key (ID)
);

comment on table DCM_ROLE_TEMPLATE is
'角色模版对应表';

comment on column DCM_ROLE_TEMPLATE.ID is
'标识ID';

comment on column DCM_ROLE_TEMPLATE.ROLE_ID is
'角色ID';

comment on column DCM_ROLE_TEMPLATE.TEMPLATE_ID is
'模版ID';

comment on column DCM_ROLE_TEMPLATE.CREATED_AT is
'创建时间';

comment on column DCM_ROLE_TEMPLATE.UPDATED_AT is
'更新时间';

comment on column DCM_ROLE_TEMPLATE.UPDATED_BY is
'更新者';

comment on column DCM_ROLE_TEMPLATE.CREATED_BY is
'创建者';

comment on column DCM_ROLE_TEMPLATE.READ_ONLY is
'只读标识';

/*==============================================================*/
/* Table: DCM_TEMPLATE                                          */
/*==============================================================*/
create table DCM_TEMPLATE 
(
   ID                   VARCHAR2(32)         not null,
   LOCALE               VARCHAR2(10)         not null,
   NAME                 VARCHAR2(300),
   CREATED_AT           DATE,
   UPDATED_AT           DATE,
   UPDATED_BY           VARCHAR2(32),
   CREATED_BY           VARCHAR2(32),
   READONLY             VARCHAR2(10),
   SEQ                  INTEGER,
   DESCRIPTION          VARCHAR2(1000),
   DB_TABLE             VARCHAR2(100),
   DB_VIEW              VARCHAR2(100),
   TMP_TABLE            VARCHAR2(100),
   PRE_PROGRAM          VARCHAR2(200),
   HANDLE_PROGRAM       VARCHAR2(200),
   AFTER_PROGRAM        VARCHAR2(200),
   HANDLE_MODE          VARCHAR2(10),
   TEMPLATE_FILE        VARCHAR2(300),
   DATA_START_LINE    NUMBER,
   COMBINATION_ID       VARCHAR2(32),
   CATEGORY_ID          VARCHAR2(32),
   constraint PK_DCM_TEMPLATE primary key (ID, LOCALE)
);

comment on table DCM_TEMPLATE is
'模版';

comment on column DCM_TEMPLATE.ID is
'标识ID';

comment on column DCM_TEMPLATE.LOCALE is
'语言';

comment on column DCM_TEMPLATE.NAME is
'模版名称';

comment on column DCM_TEMPLATE.CREATED_AT is
'创建时间';

comment on column DCM_TEMPLATE.UPDATED_AT is
'更新时间';

comment on column DCM_TEMPLATE.UPDATED_BY is
'更新者';

comment on column DCM_TEMPLATE.CREATED_BY is
'创建者';

comment on column DCM_TEMPLATE.READONLY is
'只读';

comment on column DCM_TEMPLATE.SEQ is
'序号';

comment on column DCM_TEMPLATE.DESCRIPTION is
'描述';

comment on column DCM_TEMPLATE.DB_TABLE is
'数据库表';

comment on column DCM_TEMPLATE.DB_VIEW is
'数据库视图';

comment on column DCM_TEMPLATE.TMP_TABLE is
'临时表';

comment on column DCM_TEMPLATE.PRE_PROGRAM is
'预处理程序';

comment on column DCM_TEMPLATE.HANDLE_PROGRAM is
'处理程序';

comment on column DCM_TEMPLATE.AFTER_PROGRAM is
'善后程序';

comment on column DCM_TEMPLATE.HANDLE_MODE is
'处理模式';

comment on column DCM_TEMPLATE.TEMPLATE_FILE is
'模版文件路径';

comment on column DCM_TEMPLATE.DATA_START_LINE is
'数据起始行';

comment on column DCM_TEMPLATE.COMBINATION_ID is
'组合ID';

comment on column DCM_TEMPLATE.CATEGORY_ID is
'模版分类ID';

/*==============================================================*/
/* Table: DCM_TEMPLATE_COLUMN                                   */
/*==============================================================*/
create table DCM_TEMPLATE_COLUMN 
(
   ID                   VARCHAR2(32)         not null,
   LOCALE               VARCHAR2(10)         not null,
   COLUMN_LABEL         VARCHAR2(300),
   DB_TABLE_COL         VARCHAR2(100),
   CREATED_AT           DATE,
   UPDATED_AT           DATE,
   UPDATED_BY           VARCHAR2(32),
   CREATED_BY           VARCHAR2(32),
   IS_PK                VARCHAR2(10),
   READONLY             VARCHAR2(10),
   DATA_TYPE            VARCHAR2(100),
   VISIBLE              VARCHAR2(10),
   SEQ                  INTEGER,
   TEMPLATE_ID          VARCHAR2(32),
   VALUE_SET_ID         VARCHAR2(32),
   constraint PK_DCM_TEMPLATE_COLUMN primary key (ID, LOCALE)
);

comment on table DCM_TEMPLATE_COLUMN is
'模版列';

comment on column DCM_TEMPLATE_COLUMN.ID is
'标识ID';

comment on column DCM_TEMPLATE_COLUMN.LOCALE is
'语言';

comment on column DCM_TEMPLATE_COLUMN.COLUMN_LABEL is
'列文本';

comment on column DCM_TEMPLATE_COLUMN.DB_TABLE_COL is
'数据库表列';

comment on column DCM_TEMPLATE_COLUMN.CREATED_AT is
'创建时间';

comment on column DCM_TEMPLATE_COLUMN.UPDATED_AT is
'更新时间';

comment on column DCM_TEMPLATE_COLUMN.UPDATED_BY is
'更新者';

comment on column DCM_TEMPLATE_COLUMN.CREATED_BY is
'创建者';

comment on column DCM_TEMPLATE_COLUMN.IS_PK is
'主键标识';

comment on column DCM_TEMPLATE_COLUMN.READONLY is
'只读';

comment on column DCM_TEMPLATE_COLUMN.DATA_TYPE is
'数据类型';

comment on column DCM_TEMPLATE_COLUMN.VISIBLE is
'是否可见';

comment on column DCM_TEMPLATE_COLUMN.SEQ is
'序号';

comment on column DCM_TEMPLATE_COLUMN.TEMPLATE_ID is
'模版ID';

comment on column DCM_TEMPLATE_COLUMN.VALUE_SET_ID is
'值集ID';

/*==============================================================*/
/* Table: DCM_TEMPLATE_COMBINATION                              */
/*==============================================================*/
create table DCM_TEMPLATE_COMBINATION 
(
   ID                   VARCHAR2(32)         not null,
   TEMPLATE_ID          VARCHAR2(32),
   COM_RECORD_ID       VARCHAR2(32),
   STATUS               VARCHAR2(10),
   CREATED_AT           DATE,
   UPDATED_AT           DATE,
   UPDATED_BY           VARCHAR2(32),
   CREATED_BY           VARCHAR2(32),
   constraint PK_DCM_TEMPLATE_COMBINATION primary key (ID)
);

comment on table DCM_TEMPLATE_COMBINATION is
'模版组合控制';

comment on column DCM_TEMPLATE_COMBINATION.ID is
'标识ID';

comment on column DCM_TEMPLATE_COMBINATION.TEMPLATE_ID is
'模版ID';

comment on column DCM_TEMPLATE_COMBINATION.COM_RECORD_ID is
'组合记录ID';

comment on column DCM_TEMPLATE_COMBINATION.STATUS is
'状态';

comment on column DCM_TEMPLATE_COMBINATION.CREATED_AT is
'创建时间';

comment on column DCM_TEMPLATE_COMBINATION.UPDATED_AT is
'更新时间';

comment on column DCM_TEMPLATE_COMBINATION.UPDATED_BY is
'更新者';

comment on column DCM_TEMPLATE_COMBINATION.CREATED_BY is
'创建者';

/*==============================================================*/
/* Table: DCM_ERROR                                             */
/*==============================================================*/
create table DCM_ERROR 
(
   ID                   VARCHAR2(32)         not null,
   TEMPLATE_ID          VARCHAR2(32)         not null,
   COM_RECORD_ID        VARCHAR2(32),
   SHEET_NAME           VARCHAR2(300),     
   ROW_NUM              INTEGER,
   MSG                  VARCHAR2(3000),
   ELEVEL               VARCHAR2(10),
   LOCALE               VARCHAR2(10)         not null,
   VALIDATION_ID        VARCHAR2(32),
   CREATED_AT           DATE,
   UPDATED_AT           DATE,
   UPDATED_BY           VARCHAR2(32),
   CREATED_BY           VARCHAR2(32),
   constraint PK_DCM_ERROR primary key (ID, LOCALE)
);

comment on table DCM_ERROR is
'错误信息';

comment on column DCM_ERROR.TEMPLATE_ID is
'模版ID';

comment on column DCM_ERROR.COM_RECORD_ID is
'组合ID';

comment on column DCM_ERROR.SHEET_NAME is
'工作簿名称';

comment on column DCM_ERROR.ROW_NUM is
'行号';

comment on column DCM_ERROR.MSG is
'错误消息';

comment on column DCM_ERROR.ELEVEL is
'错误级别';

comment on column DCM_ERROR.LOCALE is
'语言';

comment on column DCM_ERROR.VALIDATION_ID is
'校验程序ID';

comment on column DCM_ERROR.CREATED_AT is
'创建时间';

comment on column DCM_ERROR.UPDATED_AT is
'更新时间';

comment on column DCM_ERROR.UPDATED_BY is
'更新者';

comment on column DCM_ERROR.CREATED_BY is
'创建者';

/*==============================================================*/
/* Table: DCM_JOB                                               */
/*==============================================================*/
create table DCM_JOB 
(
   ID                   VARCHAR2(32)         not null,
   TEMPLATE_ID          VARCHAR2(32),
   COM_RECORD_ID        VARCHAR2(32),
   "MODE"               VARCHAR2(10),
   FILE_PATH            VARCHAR2(300),
   CREATED_AT           DATE,
   UPDATED_AT           DATE,
   UPDATED_BY           VARCHAR2(32),
   CREATED_BY           VARCHAR2(32),
   STATUS               VARCHAR2(10),
   constraint PK_DCM_JOB primary key (ID)
);
comment on table DCM_JOB is
'导入任务';

comment on column DCM_JOB.ID is
'标识ID';

comment on column DCM_JOB.TEMPLATE_ID is
'模版ID';

comment on column DCM_JOB.COM_RECORD_ID is
'组合ID';

comment on column DCM_JOB."MODE" is
'导入模式';

comment on column DCM_JOB.FILE_PATH is
'文件路径';

comment on column DCM_JOB.CREATED_AT is
'创建时间';

comment on column DCM_JOB.UPDATED_AT is
'更新时间';

comment on column DCM_JOB.UPDATED_BY is
'更新者';

comment on column DCM_JOB.CREATED_BY is
'创建者';

comment on column DCM_JOB.STATUS is
'状态';

/*==============================================================*/
/* Table: DCM_TEMPLATE_VALIDATION                               */
/*==============================================================*/
create table DCM_TEMPLATE_VALIDATION 
(
   ID                   VARCHAR2(32)         not null,
   COLUMN_ID            VARCHAR2(32),
   VALIDATION_ID        VARCHAR2(32),
   SEQ                  INTEGER,
   ARGS                 VARCHAR2(1000),
   CREATED_AT           DATE,
   UPDATED_AT           DATE,
   UPDATED_BY           VARCHAR2(32),
   CREATED_BY           VARCHAR2(32),
   constraint PK_DCM_TEMPLATE_VALIDATION primary key (ID)
);

comment on table DCM_TEMPLATE_VALIDATION is
'模版校验程序对应表';

comment on column DCM_TEMPLATE_VALIDATION.ID is
'标识ID';

comment on column DCM_TEMPLATE_VALIDATION.COLUMN_ID is
'模版列ID';

comment on column DCM_TEMPLATE_VALIDATION.VALIDATION_ID is
'校验程序ID';

comment on column DCM_TEMPLATE_VALIDATION.SEQ is
'序号';

comment on column DCM_TEMPLATE_VALIDATION.CREATED_AT is
'创建时间';

comment on column DCM_TEMPLATE_VALIDATION.UPDATED_AT is
'更新时间';

comment on column DCM_TEMPLATE_VALIDATION.UPDATED_BY is
'更新者';

comment on column DCM_TEMPLATE_VALIDATION.CREATED_BY is
'创建者';

comment on column DCM_TEMPLATE_VALIDATION.ARGS is
'参数';

/*==============================================================*/
/* Table: DCM_VALIDATION                                        */
/*==============================================================*/
create table DCM_VALIDATION 
(
   ID                   VARCHAR2(32)         not null,
   LOCALE               VARCHAR2(10)         not null,
   NAME                 VARCHAR2(100),
   PROGRAM              VARCHAR2(300),
   DESCRIPTION          VARCHAR2(1000),
   CREATED_AT           DATE,
   UPDATED_AT           DATE,
   UPDATED_BY           VARCHAR2(32),
   CREATED_BY           VARCHAR2(32),
   constraint PK_DCM_VALIDATION primary key (ID, LOCALE)
);

comment on table DCM_VALIDATION is
'校验程序';

comment on column DCM_VALIDATION.ID is
'标识ID';

comment on column DCM_VALIDATION.LOCALE is
'语言';

comment on column DCM_VALIDATION.NAME is
'校验程序名';

comment on column DCM_VALIDATION.PROGRAM is
'校验程序';

comment on column DCM_VALIDATION.DESCRIPTION is
'描述';

comment on column DCM_VALIDATION.CREATED_AT is
'创建时间';

comment on column DCM_VALIDATION.UPDATED_AT is
'更新时间';

comment on column DCM_VALIDATION.UPDATED_BY is
'更新者';

comment on column DCM_VALIDATION.CREATED_BY is
'创建者';

/*==============================================================*/
/* Table: DMS_AUDIT_MSG                                         */
/*==============================================================*/
create table DMS_AUDIT_MSG 
(
   ID                   VARCHAR2(32)         not null,
   LOCALE               VARCHAR2(10)         not null,
   CATEGORY             VARCHAR2(100),
   MSG                  VARCHAR2(4000),
   UPDATED_BY           VARCHAR2(32),
   CREATED_BY           VARCHAR2(32),
   CREATED_AT           DATE,
   UPDATED_AT           DATE,
   constraint PK_DMS_AUDIT_MSG primary key (ID, LOCALE)
);

comment on table DMS_AUDIT_MSG is
'用户追踪';

comment on column DMS_AUDIT_MSG.ID is
'标识ID';

comment on column DMS_AUDIT_MSG.LOCALE is
'语言';

comment on column DMS_AUDIT_MSG.CATEGORY is
'分类';

comment on column DMS_AUDIT_MSG.MSG is
'消息';

comment on column DMS_AUDIT_MSG.UPDATED_BY is
'更新者';

comment on column DMS_AUDIT_MSG.CREATED_BY is
'创建者';

comment on column DMS_AUDIT_MSG.CREATED_AT is
'创建时间';

comment on column DMS_AUDIT_MSG.UPDATED_AT is
'更新时间';

/*==============================================================*/
/* Table: DMS_FUNCTION                                          */
/*==============================================================*/
create table DMS_FUNCTION 
(
   ID                   VARCHAR2(32)         not null,
   LOCALE               VARCHAR2(10)         not null,
   DESCRIPTION          VARCHAR2(1000),
   NAME                 VARCHAR2(100),
   CREATED_AT           DATE,
   UPDATED_AT           DATE,
   UPDATED_BY           VARCHAR2(32),
   CREATED_BY           VARCHAR2(32),
   ACTION               VARCHAR2(100),
   PARAMETERS           VARCHAR2(1000),
   CATEGORY             VARCHAR2(100),
   SEQ                  INTEGER,
   P_FUNCTION_ID        VARCHAR2(32),
   constraint PK_DMS_FUNCTION primary key (ID,LOCALE)
);

comment on table DMS_FUNCTION is
'功能表';

comment on column DMS_FUNCTION.ID is
'标识ID';

comment on column DMS_FUNCTION.LOCALE is
'语言';

comment on column DMS_FUNCTION.DESCRIPTION is
'描述';

comment on column DMS_FUNCTION.NAME is
'功能名';

comment on column DMS_FUNCTION.CREATED_AT is
'创建时间';

comment on column DMS_FUNCTION.UPDATED_AT is
'更新时间';

comment on column DMS_FUNCTION.UPDATED_BY is
'更新者';

comment on column DMS_FUNCTION.CREATED_BY is
'创建者';

comment on column DMS_FUNCTION.ACTION is
'动作';

comment on column DMS_FUNCTION.PARAMETERS is
'参数';

comment on column DMS_FUNCTION.CATEGORY is
'分类';

comment on column DMS_FUNCTION.SEQ is
'序号';

comment on column DMS_FUNCTION.P_FUNCTION_ID is
'父功能ID';

/*==============================================================*/
/* Table: DMS_GROUP                                             */
/*==============================================================*/
create table DMS_GROUP 
(
   ID                   VARCHAR2(32)         not null,
   LOCALE               VARCHAR2(10)         not null,
   NAME                 VARCHAR2(100),
   ENABLE_FLAG          VARCHAR2(10),
   CREATED_AT           DATE,
   UPDATED_AT           DATE,
   UPDATED_BY           VARCHAR2(32),
   CREATED_BY           VARCHAR2(32),
   constraint PK_DMS_GROUP primary key (ID, LOCALE)
);

comment on table DMS_GROUP is
'用户组';

comment on column DMS_GROUP.ID is
'标识ID';

comment on column DMS_GROUP.LOCALE is
'语言';

comment on column DMS_GROUP.NAME is
'组名称';

comment on column DMS_GROUP.ENABLE_FLAG is
'有效标识';

comment on column DMS_GROUP.CREATED_AT is
'创建时间';

comment on column DMS_GROUP.UPDATED_AT is
'更新时间';

comment on column DMS_GROUP.UPDATED_BY is
'更新者';

comment on column DMS_GROUP.CREATED_BY is
'创建者';

/*==============================================================*/
/* Table: DMS_GROUP_ROLE                                        */
/*==============================================================*/
create table DMS_GROUP_ROLE 
(
   ID                   VARCHAR2(32)         not null,
   GROUP_ID              VARCHAR2(32),
   ROLE_ID              VARCHAR2(32),
   CREATED_AT           DATE,
   UPDATED_AT           DATE,
   UPDATED_BY           VARCHAR2(32),
   CREATED_BY           VARCHAR2(32),
   constraint PK_DMS_GROUP_ROLE primary key (ID)
);

comment on table DMS_GROUP_ROLE is
'用户组角色对应表';

comment on column DMS_GROUP_ROLE.ID is
'标识ID';

comment on column DMS_GROUP_ROLE.GROUP_ID is
'用户组ID';

comment on column DMS_GROUP_ROLE.ROLE_ID is
'角色ID';

comment on column DMS_GROUP_ROLE.CREATED_AT is
'创建时间';

comment on column DMS_GROUP_ROLE.UPDATED_AT is
'更新时间';

comment on column DMS_GROUP_ROLE.UPDATED_BY is
'更新者';

comment on column DMS_GROUP_ROLE.CREATED_BY is
'创建者';

/*==============================================================*/
/* Table: DMS_MENU                                              */
/*==============================================================*/
create table DMS_MENU 
(
   ID                   VARCHAR2(32)         not null,
   LOCALE               VARCHAR2(10)         not null,
   SEQ                  INTEGER,
   LABEL                VARCHAR2(150),
   ENABLE_FLAG          VARCHAR2(10),
   P_ID                 VARCHAR2(32),
   FUNCTION_ID         VARCHAR2(32),
   CREATED_AT           DATE,
   UPDATED_AT           DATE,
   UPDATED_BY           VARCHAR2(32),
   CREATED_BY           VARCHAR2(32),
   constraint PK_DMS_MENU primary key (ID, LOCALE)
);

comment on table DMS_MENU is
'菜单';

comment on column DMS_MENU.ID is
'标识ID';

comment on column DMS_MENU.LOCALE is
'语言';

comment on column DMS_MENU.SEQ is
'序号';

comment on column DMS_MENU.LABEL is
'显示文本';

comment on column DMS_MENU.ENABLE_FLAG is
'有效标识';

comment on column DMS_MENU.P_ID is
'父菜单ID';

comment on column DMS_MENU.FUNCTION_ID is
'功能ID';

comment on column DMS_MENU.CREATED_AT is
'创建时间';

comment on column DMS_MENU.UPDATED_AT is
'更新时间';

comment on column DMS_MENU.UPDATED_BY is
'更新者';

comment on column DMS_MENU.CREATED_BY is
'创建者';

/*==============================================================*/
/* Table: DMS_PROPERTY                                          */
/*==============================================================*/
create table DMS_PROPERTY 
(
   ID                   VARCHAR2(32)         not null,
   CKEY                  VARCHAR2(100),
   CVALUE                VARCHAR2(100),
   ENABLE_FLAG          VARCHAR2(10),
   CDESC               VARCHAR2(300),
   CREATED_AT           DATE,
   UPDATED_AT           DATE,
   UPDATED_BY           VARCHAR2(32),
   CREATED_BY           VARCHAR2(32),
   constraint PK_DMS_PROPERTY primary key (ID)
);

comment on table DMS_PROPERTY is
'属性表';

comment on column DMS_PROPERTY.ID is
'标识ID';

comment on column DMS_PROPERTY.CKEY is
'属性';

comment on column DMS_PROPERTY.CVALUE is
'属性值';

comment on column DMS_PROPERTY.ENABLE_FLAG is
'有效标识';

comment on column DMS_PROPERTY."CDESC" is
'描述';

comment on column DMS_PROPERTY.CREATED_AT is
'创建时间';

comment on column DMS_PROPERTY.UPDATED_AT is
'更新时间';

comment on column DMS_PROPERTY.UPDATED_BY is
'更新者';

comment on column DMS_PROPERTY.CREATED_BY is
'创建者';

/*==============================================================*/
/* Table: DMS_ROLE                                              */
/*==============================================================*/
create table DMS_ROLE 
(
   ID                   VARCHAR2(32)         not null,
   LOCALE               VARCHAR2(10)         not null,
   ROLE_NAME            VARCHAR2(100),
   ENABLE_FLAG          VARCHAR2(10),
   CREATED_AT           DATE,
   UPDATED_AT           DATE,
   UPDATED_BY           VARCHAR2(32),
   CREATED_BY           VARCHAR2(32),
   constraint PK_DMS_ROLE primary key (ID, LOCALE)
);

comment on table DMS_ROLE is
'角色';

comment on column DMS_ROLE.ID is
'标识ID';

comment on column DMS_ROLE.LOCALE is
'语言';

comment on column DMS_ROLE.ROLE_NAME is
'角色名';

comment on column DMS_ROLE.ENABLE_FLAG is
'有效标识';

comment on column DMS_ROLE.CREATED_AT is
'创建时间';

comment on column DMS_ROLE.UPDATED_AT is
'更新时间';

comment on column DMS_ROLE.UPDATED_BY is
'更新者';

comment on column DMS_ROLE.CREATED_BY is
'创建者';

/*==============================================================*/
/* Table: DMS_ROLE_FUNCTION                                     */
/*==============================================================*/
create table DMS_ROLE_FUNCTION 
(
   ID     VARCHAR2(32)         not null,
   ROLE_ID              VARCHAR2(32),
   FUNCTION_ID          VARCHAR2(32),
   CREATED_AT           DATE,
   UPDATED_AT           DATE,
   UPDATED_BY           VARCHAR2(32),
   CREATED_BY           VARCHAR2(32),
   constraint PK_DMS_ROLE_FUNCTION primary key (ID)
);

comment on table DMS_ROLE_FUNCTION is
'角色功能对应表';

comment on column DMS_ROLE_FUNCTION.ID is
'ID';

comment on column DMS_ROLE_FUNCTION.ROLE_ID is
'角色ID';

comment on column DMS_ROLE_FUNCTION.FUNCTION_ID is
'功能ID';

comment on column DMS_ROLE_FUNCTION.CREATED_AT is
'创建时间';

comment on column DMS_ROLE_FUNCTION.UPDATED_AT is
'更新时间';

comment on column DMS_ROLE_FUNCTION.UPDATED_BY is
'更新者';

comment on column DMS_ROLE_FUNCTION.CREATED_BY is
'创建者';

/*==============================================================*/
/* Table: DMS_ROLE_MENU                                         */
/*==============================================================*/
create table DMS_ROLE_MENU 
(
   ID                   VARCHAR2(32)         not null,
   ROLE_ID              VARCHAR2(32),
   MENU_ID              VARCHAR2(32),
   CREATED_AT           DATE,
   UPDATED_AT           DATE,
   UPDATED_BY           VARCHAR2(32),
   CREATED_BY           VARCHAR2(32),
   constraint PK_DMS_ROLE_MENU primary key (ID)
);

comment on table DMS_ROLE_MENU is
'角色菜单对应表';

comment on column DMS_ROLE_MENU.ID is
'标识ID';

comment on column DMS_ROLE_MENU.ROLE_ID is
'角色ID';

comment on column DMS_ROLE_MENU.MENU_ID is
'菜单ID';

comment on column DMS_ROLE_MENU.CREATED_AT is
'创建时间';

comment on column DMS_ROLE_MENU.UPDATED_AT is
'更新时间';

comment on column DMS_ROLE_MENU.UPDATED_BY is
'更新者';

comment on column DMS_ROLE_MENU.CREATED_BY is
'创建者';

/*==============================================================*/
/* Table: DMS_GROUP_VALUE                                        */
/*==============================================================*/
create table DMS_GROUP_VALUE 
(
   ID                   VARCHAR2(32)         not null,
   GROUP_ID              VARCHAR2(32),
   VALUE_SET_ID         VARCHAR2(32),
   VALUE_ID             VARCHAR2(100),
   UPDATED_BY           VARCHAR2(32),
   CREATED_BY           VARCHAR2(32),
   CREATED_AT           DATE,
   UPDATED_AT           DATE,
   constraint PK_DMS_GROUP_VALUE primary key (ID)
);

comment on table DMS_GROUP_VALUE is
'用户组和值对应表';

comment on column DMS_GROUP_VALUE.ID is
'标识ID';

comment on column DMS_GROUP_VALUE.GROUP_ID is
'角色ID';

comment on column DMS_GROUP_VALUE.VALUE_SET_ID is
'值集ID';

comment on column DMS_GROUP_VALUE.VALUE_ID is
'值ID';

comment on column DMS_GROUP_VALUE.UPDATED_BY is
'更新者';

comment on column DMS_GROUP_VALUE.CREATED_BY is
'创建者';

comment on column DMS_GROUP_VALUE.CREATED_AT is
'创建时间';

comment on column DMS_GROUP_VALUE.UPDATED_AT is
'更新时间';

/*==============================================================*/
/* Table: DMS_USER                                              */
/*==============================================================*/
create table DMS_USER 
(
   ID                   VARCHAR2(32)         not null,
   ACC                  VARCHAR2(100),
   PWD                  VARCHAR2(100),
   LOCK_FLAG            VARCHAR2(10),
   ENABLE_FLAG          VARCHAR2(10),
   RETRY_COUNT          INTEGER,
   LOCK_TIME            DATE,
   UPDATED_AT           DATE,
   CREATED_AT           DATE,
   CREATED_BY           VARCHAR2(32),
   UPDATED_BY           VARCHAR2(32),
   NAME                 VARCHAR2(100),
   SEX                  VARCHAR2(10),
   MAIL                 VARCHAR2(100),
   LOCALE               VARCHAR2(32),
   PHONE                VARCHAR2(100),
   OTHER_INFO           VARCHAR2(1000),
   constraint PK_DMS_USER primary key (ID)
);
alter table DMS_USER
  add constraint UK_DMS_USER_ACC unique (ACC);
comment on table DMS_USER is
'用户表';

comment on column DMS_USER.ID is
'标识ID';

comment on column DMS_USER.ACC is
'帐号';

comment on column DMS_USER.PWD is
'密码';

comment on column DMS_USER.LOCK_FLAG is
'锁定标识';

comment on column DMS_USER.ENABLE_FLAG is
'有效标识';

comment on column DMS_USER.RETRY_COUNT is
'重试次数';

comment on column DMS_USER.LOCK_TIME is
'锁定时间';

comment on column DMS_USER.UPDATED_AT is
'更新时间';

comment on column DMS_USER.CREATED_AT is
'创建时间';

comment on column DMS_USER.CREATED_BY is
'创建者';

comment on column DMS_USER.UPDATED_BY is
'更新者';

comment on column DMS_USER.NAME is
'姓名';

comment on column DMS_USER.SEX is
'性别';

comment on column DMS_USER.MAIL is
'邮箱';

comment on column DMS_USER.LOCALE is
'语言';

comment on column DMS_USER.PHONE is
'电话';

comment on column DMS_USER.OTHER_INFO is
'其它信息';

/*==============================================================*/
/* Table: DMS_USER_GROUP                                        */
/*==============================================================*/
create table DMS_USER_GROUP 
(
   ID                   VARCHAR2(32)         not null,
   USER_ID              VARCHAR2(32),
   GROUP_ID              VARCHAR2(32),
   CREATED_AT           DATE,
   UPDATED_AT           DATE,
   UPDATED_BY           VARCHAR2(32),
   CREATED_BY           VARCHAR2(32),
   constraint PK_DMS_USER_GROUP primary key (ID)
);

comment on table DMS_USER_GROUP is
'用户用户组对应表';

comment on column DMS_USER_GROUP.ID is
'标识ID';

comment on column DMS_USER_GROUP.USER_ID is
'用户ID';

comment on column DMS_USER_GROUP.GROUP_ID is
'用户组ID';

comment on column DMS_USER_GROUP.CREATED_AT is
'创建时间';

comment on column DMS_USER_GROUP.UPDATED_AT is
'更新时间';

comment on column DMS_USER_GROUP.UPDATED_BY is
'更新者';

comment on column DMS_USER_GROUP.CREATED_BY is
'创建者';

/*==============================================================*/
/* Table: DMS_USER_KEY                                          */
/*==============================================================*/
create table DMS_USER_KEY 
(
   ID                   VARCHAR2(32)         not null,
   USER_ID              VARCHAR2(32),
   KEY                  VARCHAR2(100),
   CREATED_AT           DATE,
   constraint PK_DMS_USER_KEY primary key (ID)
);

comment on table DMS_USER_KEY is
'用户临时凭证表';

comment on column DMS_USER_KEY.ID is
'标识ID';

comment on column DMS_USER_KEY.USER_ID is
'用户ID';

comment on column DMS_USER_KEY.KEY is
'用户KEY';

comment on column DMS_USER_KEY.CREATED_AT is
'创建时间';

/*==============================================================*/
/* Table: DMS_VALUE_SET                                         */
/*==============================================================*/
create table DMS_VALUE_SET 
(
   ID                   VARCHAR2(32)         not null,
   LOCALE               VARCHAR2(10)         not null,
   NAME                 VARCHAR2(100),
   CODE                 VARCHAR2(100)        not null,
   SOURCE               VARCHAR2(100),
   CREATED_AT           DATE,
   UPDATED_AT           DATE,
   UPDATED_BY           VARCHAR2(32),
   CREATED_BY           VARCHAR2(32),
   constraint PK_DMS_VALUE_SET primary key (ID, LOCALE)
);

comment on table DMS_VALUE_SET is
'值集';

comment on column DMS_VALUE_SET.ID is
'标识ID';

comment on column DMS_VALUE_SET.LOCALE is
'语言';

comment on column DMS_VALUE_SET.NAME is
'值集名';

comment on column DMS_VALUE_SET.code is
'值集编码';

comment on column DMS_VALUE_SET.SOURCE is
'值集源';

comment on column DMS_VALUE_SET.CREATED_AT is
'创建时间';

comment on column DMS_VALUE_SET.UPDATED_AT is
'更新时间';

comment on column DMS_VALUE_SET.UPDATED_BY is
'更新者';

comment on column DMS_VALUE_SET.CREATED_BY is
'创建者';

/*==============================================================*/
/* Table: DMS_LOOKUP                                         */
/*==============================================================*/
create table DMS_LOOKUP
(
   ID                   VARCHAR2(32)         not null,
   LOCALE               VARCHAR2(10)         not null,
   LOOKUP_TYPE          VARCHAR2(100)        NOT NULL,
   CODE                 VARCHAR2(100)        not null,
   MEANING              VARCHAR2(100)        not null,
   IDX                  NUMBER,
   CREATED_AT           DATE,
   UPDATED_AT           DATE,
   UPDATED_BY           VARCHAR2(32),
   CREATED_BY           VARCHAR2(32),
   constraint PK_DMS_LOOKUP primary key (LOOKUP_TYPE, LOCALE,CODE)
);

comment on table DMS_LOOKUP is
'字典表';

comment on column DMS_LOOKUP.ID is
'标识ID';

comment on column DMS_LOOKUP.LOCALE is
'语言';

comment on column DMS_LOOKUP.LOOKUP_TYPE is
'分类';

comment on column DMS_LOOKUP.CODE is
'编码';

comment on column DMS_LOOKUP.MEANING is
'含义';

comment on column DMS_LOOKUP.IDX is
'序号';

comment on column DMS_LOOKUP.CREATED_AT is
'创建时间';

comment on column DMS_LOOKUP.UPDATED_AT is
'更新时间';

comment on column DMS_LOOKUP.UPDATED_BY is
'更新者';

comment on column DMS_LOOKUP.CREATED_BY is
'创建者';

/*==============================================================*/
/* View: DMS_USER_VALUE_V                                       */
/*==============================================================*/
CREATE VIEW DMS_USER_VALUE_V AS
SELECT DISTINCT UG.USER_ID, GV.VALUE_SET_ID, GV.VALUE_ID
  FROM DMS_GROUP_VALUE GV, DMS_USER_GROUP UG, DMS_GROUP GRP
 WHERE GV.GROUP_ID = UG.GROUP_ID
   AND GRP.ENABLE_FLAG = 'Y'
   AND GRP.ID = UG.GROUP_ID
 ORDER BY GV.VALUE_ID;

/*==============================================================*/
/* View: DMS_USER_FUNCTION_V                                    */
/*==============================================================*/
CREATE VIEW DMS_USER_FUNCTION_V AS
SELECT DISTINCT UG.USER_ID, RF.FUNCTION_ID
  FROM DMS_ROLE_FUNCTION RF,
       DMS_USER_GROUP    UG,
       DMS_GROUP_ROLE    GR,
       DMS_GROUP         GRP,
       DMS_ROLE          RL
 WHERE RF.ROLE_ID = GR.ROLE_ID
   AND GR.GROUP_ID = UG.GROUP_ID
   AND UG.GROUP_ID = GRP.ID
   AND RF.ROLE_ID = RL.ID
   AND GRP.ENABLE_FLAG = 'Y'
   AND RL.ENABLE_FLAG = 'Y';

/*==============================================================*/
/* View: DCM_USER_TEMPLATE_V                                    */
/*==============================================================*/
CREATE VIEW DCM_USER_TEMPLATE_V AS
SELECT DISTINCT UG.USER_ID, RT.TEMPLATE_ID, RT.READ_ONLY
  FROM DCM_ROLE_TEMPLATE RT,
       DMS_GROUP_ROLE    GR,
       DMS_USER_GROUP    UG,
       DMS_ROLE          RL,
       DMS_GROUP         GRP
 WHERE RT.ROLE_ID = GR.ROLE_ID
   AND GR.GROUP_ID = UG.GROUP_ID
   AND RT.ROLE_ID = RL.ID
   AND UG.GROUP_ID = GRP.ID
   AND RL.ENABLE_FLAG = 'Y'
   AND GRP.ENABLE_FLAG = 'Y'
 ORDER BY RT.READ_ONLY;
   
/*==============================================================*/
/* Table: DCM_TEMPTABLE10                                         */
/*==============================================================*/
CREATE TABLE DCM_TEMPTABLE10(
  TEMPLATE_ID   VARCHAR2(32),
  COM_RECORD_ID VARCHAR2(32),
  SHEET_NAME    VARCHAR2(300),
  ROW_NO        NUMBER,
  EDIT_TYPE     VARCHAR2(10),
  IDX           NUMBER,
  COLUMN1       VARCHAR2(300),
  COLUMN2       VARCHAR2(300),
  COLUMN3       VARCHAR2(300),
  COLUMN4       VARCHAR2(300),
  COLUMN5       VARCHAR2(300),
  COLUMN6       VARCHAR2(300),
  COLUMN7       VARCHAR2(300),
  COLUMN8       VARCHAR2(300),
  COLUMN9       VARCHAR2(300),
  COLUMN10      VARCHAR2(300),
  ORIGIN_ROWID  VARCHAR2(32),
  CREATED_BY    VARCHAR2(32),
  UPDATED_BY    VARCHAR2(32),
  CREATED_AT    DATE,
  UPDATED_AT    DATE
);

/*==============================================================*/
/* Table: DCM_TEMPTABLE20                                         */
/*==============================================================*/
CREATE TABLE DCM_TEMPTABLE20(
  TEMPLATE_ID   VARCHAR2(32),
  COM_RECORD_ID VARCHAR2(32),
  SHEET_NAME    VARCHAR2(300),
  ROW_NO        NUMBER,
  EDIT_TYPE     VARCHAR2(10),
  IDX           NUMBER,
  COLUMN1       VARCHAR2(300),
  COLUMN2       VARCHAR2(300),
  COLUMN3       VARCHAR2(300),
  COLUMN4       VARCHAR2(300),
  COLUMN5       VARCHAR2(300),
  COLUMN6       VARCHAR2(300),
  COLUMN7       VARCHAR2(300),
  COLUMN8       VARCHAR2(300),
  COLUMN9       VARCHAR2(300),
  COLUMN10      VARCHAR2(300),
  COLUMN11      VARCHAR2(300),
  COLUMN12      VARCHAR2(300),
  COLUMN13      VARCHAR2(300),
  COLUMN14      VARCHAR2(300),
  COLUMN15      VARCHAR2(300),
  COLUMN16      VARCHAR2(300),
  COLUMN17      VARCHAR2(300),
  COLUMN18      VARCHAR2(300),
  COLUMN19      VARCHAR2(300),
  COLUMN20      VARCHAR2(300),
  ORIGIN_ROWID  VARCHAR2(32),
  CREATED_BY    VARCHAR2(32),
  UPDATED_BY    VARCHAR2(32),
  CREATED_AT    DATE,
  UPDATED_AT    DATE
);

/*==============================================================*/
/* Table: DCM_TEMPTABLE30                                         */
/*==============================================================*/
CREATE TABLE DCM_TEMPTABLE30(
  TEMPLATE_ID   VARCHAR2(32),
  COM_RECORD_ID VARCHAR2(32),
  SHEET_NAME    VARCHAR2(300),
  ROW_NO        NUMBER,
  EDIT_TYPE     VARCHAR2(10),
  IDX           NUMBER,
  COLUMN1       VARCHAR2(300),
  COLUMN2       VARCHAR2(300),
  COLUMN3       VARCHAR2(300),
  COLUMN4       VARCHAR2(300),
  COLUMN5       VARCHAR2(300),
  COLUMN6       VARCHAR2(300),
  COLUMN7       VARCHAR2(300),
  COLUMN8       VARCHAR2(300),
  COLUMN9       VARCHAR2(300),
  COLUMN10      VARCHAR2(300),
  COLUMN11      VARCHAR2(300),
  COLUMN12      VARCHAR2(300),
  COLUMN13      VARCHAR2(300),
  COLUMN14      VARCHAR2(300),
  COLUMN15      VARCHAR2(300),
  COLUMN16      VARCHAR2(300),
  COLUMN17      VARCHAR2(300),
  COLUMN18      VARCHAR2(300),
  COLUMN19      VARCHAR2(300),
  COLUMN20      VARCHAR2(300),
  COLUMN21      VARCHAR2(300),
  COLUMN22      VARCHAR2(300),
  COLUMN23      VARCHAR2(300),
  COLUMN24      VARCHAR2(300),
  COLUMN25      VARCHAR2(300),
  COLUMN26      VARCHAR2(300),
  COLUMN27      VARCHAR2(300),
  COLUMN28      VARCHAR2(300),
  COLUMN29      VARCHAR2(300),
  COLUMN30      VARCHAR2(300),
  ORIGIN_ROWID  VARCHAR2(32),
  CREATED_BY    VARCHAR2(32),
  UPDATED_BY    VARCHAR2(32),
  CREATED_AT    DATE,
  UPDATED_AT    DATE
);

/*==============================================================*/
/* Table: DCM_TEMPTABLE40                                         */
/*==============================================================*/
CREATE TABLE DCM_TEMPTABLE40(
  TEMPLATE_ID   VARCHAR2(32),
  COM_RECORD_ID VARCHAR2(32),
  SHEET_NAME    VARCHAR2(300),
  ROW_NO        NUMBER,
  EDIT_TYPE     VARCHAR2(10),
  IDX           NUMBER,
  COLUMN1       VARCHAR2(300),
  COLUMN2       VARCHAR2(300),
  COLUMN3       VARCHAR2(300),
  COLUMN4       VARCHAR2(300),
  COLUMN5       VARCHAR2(300),
  COLUMN6       VARCHAR2(300),
  COLUMN7       VARCHAR2(300),
  COLUMN8       VARCHAR2(300),
  COLUMN9       VARCHAR2(300),
  COLUMN10      VARCHAR2(300),
  COLUMN11      VARCHAR2(300),
  COLUMN12      VARCHAR2(300),
  COLUMN13      VARCHAR2(300),
  COLUMN14      VARCHAR2(300),
  COLUMN15      VARCHAR2(300),
  COLUMN16      VARCHAR2(300),
  COLUMN17      VARCHAR2(300),
  COLUMN18      VARCHAR2(300),
  COLUMN19      VARCHAR2(300),
  COLUMN20      VARCHAR2(300),
  COLUMN21      VARCHAR2(300),
  COLUMN22      VARCHAR2(300),
  COLUMN23      VARCHAR2(300),
  COLUMN24      VARCHAR2(300),
  COLUMN25      VARCHAR2(300),
  COLUMN26      VARCHAR2(300),
  COLUMN27      VARCHAR2(300),
  COLUMN28      VARCHAR2(300),
  COLUMN29      VARCHAR2(300),
  COLUMN30      VARCHAR2(300),
  COLUMN31      VARCHAR2(300),
  COLUMN32      VARCHAR2(300),
  COLUMN33      VARCHAR2(300),
  COLUMN34      VARCHAR2(300),
  COLUMN35      VARCHAR2(300),
  COLUMN36      VARCHAR2(300),
  COLUMN37      VARCHAR2(300),
  COLUMN38      VARCHAR2(300),
  COLUMN39      VARCHAR2(300),
  COLUMN40      VARCHAR2(300),
  ORIGIN_ROWID  VARCHAR2(32),
  CREATED_BY    VARCHAR2(32),
  UPDATED_BY    VARCHAR2(32),
  CREATED_AT    DATE,
  UPDATED_AT    DATE
);

/*==============================================================*/
/* Table: DCM_TEMPTABLE50                                         */
/*==============================================================*/
CREATE TABLE DCM_TEMPTABLE50(
  TEMPLATE_ID   VARCHAR2(32),
  COM_RECORD_ID VARCHAR2(32),
  SHEET_NAME    VARCHAR2(300),
  ROW_NO        NUMBER,
  EDIT_TYPE     VARCHAR2(10),
  IDX           NUMBER,
  COLUMN1       VARCHAR2(300),
  COLUMN2       VARCHAR2(300),
  COLUMN3       VARCHAR2(300),
  COLUMN4       VARCHAR2(300),
  COLUMN5       VARCHAR2(300),
  COLUMN6       VARCHAR2(300),
  COLUMN7       VARCHAR2(300),
  COLUMN8       VARCHAR2(300),
  COLUMN9       VARCHAR2(300),
  COLUMN10      VARCHAR2(300),
  COLUMN11      VARCHAR2(300),
  COLUMN12      VARCHAR2(300),
  COLUMN13      VARCHAR2(300),
  COLUMN14      VARCHAR2(300),
  COLUMN15      VARCHAR2(300),
  COLUMN16      VARCHAR2(300),
  COLUMN17      VARCHAR2(300),
  COLUMN18      VARCHAR2(300),
  COLUMN19      VARCHAR2(300),
  COLUMN20      VARCHAR2(300),
  COLUMN21      VARCHAR2(300),
  COLUMN22      VARCHAR2(300),
  COLUMN23      VARCHAR2(300),
  COLUMN24      VARCHAR2(300),
  COLUMN25      VARCHAR2(300),
  COLUMN26      VARCHAR2(300),
  COLUMN27      VARCHAR2(300),
  COLUMN28      VARCHAR2(300),
  COLUMN29      VARCHAR2(300),
  COLUMN30      VARCHAR2(300),
  COLUMN31      VARCHAR2(300),
  COLUMN32      VARCHAR2(300),
  COLUMN33      VARCHAR2(300),
  COLUMN34      VARCHAR2(300),
  COLUMN35      VARCHAR2(300),
  COLUMN36      VARCHAR2(300),
  COLUMN37      VARCHAR2(300),
  COLUMN38      VARCHAR2(300),
  COLUMN39      VARCHAR2(300),
  COLUMN40      VARCHAR2(300),
  COLUMN41      VARCHAR2(300),
  COLUMN42      VARCHAR2(300),
  COLUMN43      VARCHAR2(300),
  COLUMN44      VARCHAR2(300),
  COLUMN45      VARCHAR2(300),
  COLUMN46      VARCHAR2(300),
  COLUMN47      VARCHAR2(300),
  COLUMN48      VARCHAR2(300),
  COLUMN49      VARCHAR2(300),
  COLUMN50      VARCHAR2(300),
  ORIGIN_ROWID  VARCHAR2(32),
  CREATED_BY    VARCHAR2(32),
  UPDATED_BY    VARCHAR2(32),
  CREATED_AT    DATE,
  UPDATED_AT    DATE
);

/*==============================================================*/
/* Table: DCM_TEMPTABLE80                                         */
/*==============================================================*/
CREATE TABLE DCM_TEMPTABLE80(
  TEMPLATE_ID   VARCHAR2(32),
  COM_RECORD_ID VARCHAR2(32),
  SHEET_NAME    VARCHAR2(300),
  ROW_NO        NUMBER,
  EDIT_TYPE     VARCHAR2(10),
  IDX           NUMBER,
  COLUMN1       VARCHAR2(300),
  COLUMN2       VARCHAR2(300),
  COLUMN3       VARCHAR2(300),
  COLUMN4       VARCHAR2(300),
  COLUMN5       VARCHAR2(300),
  COLUMN6       VARCHAR2(300),
  COLUMN7       VARCHAR2(300),
  COLUMN8       VARCHAR2(300),
  COLUMN9       VARCHAR2(300),
  COLUMN10      VARCHAR2(300),
  COLUMN11      VARCHAR2(300),
  COLUMN12      VARCHAR2(300),
  COLUMN13      VARCHAR2(300),
  COLUMN14      VARCHAR2(300),
  COLUMN15      VARCHAR2(300),
  COLUMN16      VARCHAR2(300),
  COLUMN17      VARCHAR2(300),
  COLUMN18      VARCHAR2(300),
  COLUMN19      VARCHAR2(300),
  COLUMN20      VARCHAR2(300),
  COLUMN21      VARCHAR2(300),
  COLUMN22      VARCHAR2(300),
  COLUMN23      VARCHAR2(300),
  COLUMN24      VARCHAR2(300),
  COLUMN25      VARCHAR2(300),
  COLUMN26      VARCHAR2(300),
  COLUMN27      VARCHAR2(300),
  COLUMN28      VARCHAR2(300),
  COLUMN29      VARCHAR2(300),
  COLUMN30      VARCHAR2(300),
  COLUMN31      VARCHAR2(300),
  COLUMN32      VARCHAR2(300),
  COLUMN33      VARCHAR2(300),
  COLUMN34      VARCHAR2(300),
  COLUMN35      VARCHAR2(300),
  COLUMN36      VARCHAR2(300),
  COLUMN37      VARCHAR2(300),
  COLUMN38      VARCHAR2(300),
  COLUMN39      VARCHAR2(300),
  COLUMN40      VARCHAR2(300),
  COLUMN41      VARCHAR2(300),
  COLUMN42      VARCHAR2(300),
  COLUMN43      VARCHAR2(300),
  COLUMN44      VARCHAR2(300),
  COLUMN45      VARCHAR2(300),
  COLUMN46      VARCHAR2(300),
  COLUMN47      VARCHAR2(300),
  COLUMN48      VARCHAR2(300),
  COLUMN49      VARCHAR2(300),
  COLUMN50      VARCHAR2(300),
  COLUMN51      VARCHAR2(300),
  COLUMN52      VARCHAR2(300),
  COLUMN53      VARCHAR2(300),
  COLUMN54      VARCHAR2(300),
  COLUMN55      VARCHAR2(300),
  COLUMN56      VARCHAR2(300),
  COLUMN57      VARCHAR2(300),
  COLUMN58      VARCHAR2(300),
  COLUMN59      VARCHAR2(300),
  COLUMN60      VARCHAR2(300),
  COLUMN61      VARCHAR2(300),
  COLUMN62      VARCHAR2(300),
  COLUMN63      VARCHAR2(300),
  COLUMN64      VARCHAR2(300),
  COLUMN65      VARCHAR2(300),
  COLUMN66      VARCHAR2(300),
  COLUMN67      VARCHAR2(300),
  COLUMN68      VARCHAR2(300),
  COLUMN69      VARCHAR2(300),
  COLUMN70      VARCHAR2(300),
  COLUMN71      VARCHAR2(300),
  COLUMN72      VARCHAR2(300),
  COLUMN73      VARCHAR2(300),
  COLUMN74      VARCHAR2(300),
  COLUMN75      VARCHAR2(300),
  COLUMN76      VARCHAR2(300),
  COLUMN77      VARCHAR2(300),
  COLUMN78      VARCHAR2(300),
  COLUMN79      VARCHAR2(300),
  COLUMN80      VARCHAR2(300),
  ORIGIN_ROWID  VARCHAR2(32),
  CREATED_BY    VARCHAR2(32),
  UPDATED_BY    VARCHAR2(32),
  CREATED_AT    DATE,
  UPDATED_AT    DATE
);
/*==============================================================*/
/* Table: DCM_TEMPTABLE100                                         */
/*==============================================================*/
CREATE TABLE DCM_TEMPTABLE100(
  TEMPLATE_ID   VARCHAR2(32),
  COM_RECORD_ID VARCHAR2(32),
  SHEET_NAME    VARCHAR2(300),
  ROW_NO        NUMBER,
  EDIT_TYPE     VARCHAR2(10),
  IDX           NUMBER,
  COLUMN1       VARCHAR2(300),
  COLUMN2       VARCHAR2(300),
  COLUMN3       VARCHAR2(300),
  COLUMN4       VARCHAR2(300),
  COLUMN5       VARCHAR2(300),
  COLUMN6       VARCHAR2(300),
  COLUMN7       VARCHAR2(300),
  COLUMN8       VARCHAR2(300),
  COLUMN9       VARCHAR2(300),
  COLUMN10      VARCHAR2(300),
  COLUMN11      VARCHAR2(300),
  COLUMN12      VARCHAR2(300),
  COLUMN13      VARCHAR2(300),
  COLUMN14      VARCHAR2(300),
  COLUMN15      VARCHAR2(300),
  COLUMN16      VARCHAR2(300),
  COLUMN17      VARCHAR2(300),
  COLUMN18      VARCHAR2(300),
  COLUMN19      VARCHAR2(300),
  COLUMN20      VARCHAR2(300),
  COLUMN21      VARCHAR2(300),
  COLUMN22      VARCHAR2(300),
  COLUMN23      VARCHAR2(300),
  COLUMN24      VARCHAR2(300),
  COLUMN25      VARCHAR2(300),
  COLUMN26      VARCHAR2(300),
  COLUMN27      VARCHAR2(300),
  COLUMN28      VARCHAR2(300),
  COLUMN29      VARCHAR2(300),
  COLUMN30      VARCHAR2(300),
  COLUMN31      VARCHAR2(300),
  COLUMN32      VARCHAR2(300),
  COLUMN33      VARCHAR2(300),
  COLUMN34      VARCHAR2(300),
  COLUMN35      VARCHAR2(300),
  COLUMN36      VARCHAR2(300),
  COLUMN37      VARCHAR2(300),
  COLUMN38      VARCHAR2(300),
  COLUMN39      VARCHAR2(300),
  COLUMN40      VARCHAR2(300),
  COLUMN41      VARCHAR2(300),
  COLUMN42      VARCHAR2(300),
  COLUMN43      VARCHAR2(300),
  COLUMN44      VARCHAR2(300),
  COLUMN45      VARCHAR2(300),
  COLUMN46      VARCHAR2(300),
  COLUMN47      VARCHAR2(300),
  COLUMN48      VARCHAR2(300),
  COLUMN49      VARCHAR2(300),
  COLUMN50      VARCHAR2(300),
  COLUMN51      VARCHAR2(300),
  COLUMN52      VARCHAR2(300),
  COLUMN53      VARCHAR2(300),
  COLUMN54      VARCHAR2(300),
  COLUMN55      VARCHAR2(300),
  COLUMN56      VARCHAR2(300),
  COLUMN57      VARCHAR2(300),
  COLUMN58      VARCHAR2(300),
  COLUMN59      VARCHAR2(300),
  COLUMN60      VARCHAR2(300),
  COLUMN61      VARCHAR2(300),
  COLUMN62      VARCHAR2(300),
  COLUMN63      VARCHAR2(300),
  COLUMN64      VARCHAR2(300),
  COLUMN65      VARCHAR2(300),
  COLUMN66      VARCHAR2(300),
  COLUMN67      VARCHAR2(300),
  COLUMN68      VARCHAR2(300),
  COLUMN69      VARCHAR2(300),
  COLUMN70      VARCHAR2(300),
  COLUMN71      VARCHAR2(300),
  COLUMN72      VARCHAR2(300),
  COLUMN73      VARCHAR2(300),
  COLUMN74      VARCHAR2(300),
  COLUMN75      VARCHAR2(300),
  COLUMN76      VARCHAR2(300),
  COLUMN77      VARCHAR2(300),
  COLUMN78      VARCHAR2(300),
  COLUMN79      VARCHAR2(300),
  COLUMN80      VARCHAR2(300),
  COLUMN81      VARCHAR2(300),
  COLUMN82      VARCHAR2(300),
  COLUMN83      VARCHAR2(300),
  COLUMN84      VARCHAR2(300),
  COLUMN85      VARCHAR2(300),
  COLUMN86      VARCHAR2(300),
  COLUMN87      VARCHAR2(300),
  COLUMN88      VARCHAR2(300),
  COLUMN89      VARCHAR2(300),
  COLUMN90      VARCHAR2(300),
  COLUMN91      VARCHAR2(300),
  COLUMN92      VARCHAR2(300),
  COLUMN93      VARCHAR2(300),
  COLUMN94      VARCHAR2(300),
  COLUMN95      VARCHAR2(300),
  COLUMN96      VARCHAR2(300),
  COLUMN97      VARCHAR2(300),
  COLUMN98      VARCHAR2(300),
  COLUMN99      VARCHAR2(300),
  COLUMN100      VARCHAR2(300),
  ORIGIN_ROWID  VARCHAR2(32),
  CREATED_BY    VARCHAR2(32),
  UPDATED_BY    VARCHAR2(32),
  CREATED_AT    DATE,
  UPDATED_AT    DATE
);
CREATE OR REPLACE PACKAGE DCM_COMMON IS
  /**************************************************
  *数据预处理程序，用于在数据校验前进行数据加工
  *P_TEMPLATE_ID 模版ID
  *P_COM_RECORD_ID 组合ID
  *P_USER_ID  用户ID
  *P_HANDLE_MODE 处理模式，REPLACE表示替换，INCREMENT表示增量
  *P_LOCALE 语言编码
  ***************************************************/
  PROCEDURE PRE_PROGRAM(
       P_TEMPLATE_ID   IN VARCHAR2,
       P_COM_RECORD_ID IN VARCHAR2,
       P_USER_ID       IN VARCHAR2,
       P_HANDLE_MODE   IN VARCHAR2,
       P_LOCALE        IN VARCHAR2
  );
  /**************************************************
  *数据处理程序，用于将数据从零时表传输到正式表
  *P_TEMPLATE_ID 模版ID
  *P_COM_RECORD_ID 组合ID
  *P_USER_ID  用户ID
  *P_HANDLE_MODE 处理模式，REPLACE表示替换，INCREMENT表示增量
  *P_LOCALE 语言编码
  ***************************************************/
  PROCEDURE HANDLE_PROGRAM(
       P_TEMPLATE_ID   IN VARCHAR2,
       P_COM_RECORD_ID IN VARCHAR2,
       P_USER_ID       IN VARCHAR2,
       P_HANDLE_MODE   IN VARCHAR2,
       P_LOCALE        IN VARCHAR2
  );
  /*****************************************************
  *获取模版信息
  *P_TEMPLATE_ID 模版ID
  *P_TEMP_TALBE  临时表
  *P_SRC_TABLE   原表
  *P_LOCALE    语言编码
  ******************************************************/
  PROCEDURE FETCH_TEMPLATE_INFO(
    P_TEMPLATE_ID     IN  VARCHAR2,
    P_LOCALE          IN  VARCHAR2,
    P_TEMP_TALBE      OUT VARCHAR2,
    P_SRC_TABLE       OUT VARCHAR2
  );
  /*****************************************************
  *空值校验
  *P_VALIDATION_ID 校验程序ID
  *P_TEMPLATE_ID 模版ID
  *P_COM_RECORD_ID 组合ID
  *P_TEMP_COL 校验程序对应临时表中的列
  *P_SRC_COL  校验程序对应目标表中的列
  *P_MODE     当前数据处理模式（REPLACE,INCREMENT,EDIT）
  *P_LOCALE    语言编码
  *P_ARGS      校验程序参数
  *P_FLAG      校验结果标识(Y OR N)
  ******************************************************/
  PROCEDURE VALIDATE_NULL(
    P_VALIDATION_ID   IN  VARCHAR,
    P_TEMPLATE_ID     IN  VARCHAR2,
    P_COM_RECORD_ID   IN  VARCHAR2,
    P_TEMP_COL        IN  VARCHAR2,
    P_SRC_COL         IN  VARCHAR2,
    P_MODE            IN  VARCHAR2,
    P_LOCALE          IN  VARCHAR2,
    P_ARGS            IN  VARCHAR2,
    P_FLAG            OUT VARCHAR2
  );
  /**************************************************
  *善后处理程序，用于在数据导入到真实表后进行操作
  *P_TEMPLATE_ID 模版ID
  *P_COM_RECORD_ID 组合ID
  *P_USER_ID  用户ID
  *P_HANDLE_MODE 处理模式，REPLACE表示替换，INCREMENT表示增量
  *P_LOCALE 语言编码
  ***************************************************/
  PROCEDURE AFTER_PROGRAM(
       P_TEMPLATE_ID   IN VARCHAR2,
       P_COM_RECORD_ID IN VARCHAR2,
       P_USER_ID       IN VARCHAR2,
       P_HANDLE_MODE   IN VARCHAR2,
       P_LOCALE        IN VARCHAR2
  );
END DCM_COMMON;
/
CREATE OR REPLACE PACKAGE BODY DCM_COMMON IS
  CURSOR L_TEMPLATE_COL_CURSOR(
       P_TEMPLATE_ID IN VARCHAR2,
       P_LOCALE      IN VARCHAR2)
  IS
       SELECT T.DB_TABLE_COL,T.IS_PK,T.READONLY,T.DATA_TYPE
       FROM DCM_TEMPLATE_COLUMN T
       WHERE T.TEMPLATE_ID=P_TEMPLATE_ID
             AND T.LOCALE=P_LOCALE
       ORDER BY T.SEQ;
  /**************************************************
  *数据预处理程序，用于在数据校验前进行数据加工
  *P_TEMPLATE_ID 模版ID
  *P_COM_RECORD_ID 组合ID
  *P_USER_ID  用户ID
  *P_HANDLE_MODE 处理模式，REPLACE表示替换，INCREMENT表示增量
  *P_LOCALE 语言编码
  ***************************************************/
  PROCEDURE PRE_PROGRAM(
       P_TEMPLATE_ID   IN VARCHAR2,
       P_COM_RECORD_ID IN VARCHAR2,
       P_USER_ID       IN VARCHAR2,
       P_HANDLE_MODE   IN VARCHAR2,
       P_LOCALE        IN VARCHAR2
  )IS
  BEGIN
    DBMS_OUTPUT.PUT_LINE('PLEASE FLOW THIS DEMO TO IMPLEMENT THE ACTRAL LOGIC');
  END PRE_PROGRAM;
  /**************************************************
  *数据处理程序，用于将数据从零时表传输到正式表
  *P_TEMPLATE_ID 模版ID
  *P_COM_RECORD_ID 组合ID
  *P_USER_ID  用户ID
  *P_HANDLE_MODE 处理模式，REPLACE表示替换，INCREMENT表示增量,EDIT表示页面编辑
  *P_LOCALE 语言编码
  ***************************************************/
  PROCEDURE HANDLE_PROGRAM(
      P_TEMPLATE_ID   IN VARCHAR2,
      P_COM_RECORD_ID IN VARCHAR2,
      P_USER_ID       IN VARCHAR2,
      P_HANDLE_MODE   IN VARCHAR2,
      P_LOCALE        IN VARCHAR2
  )
  IS
  L_TEMP_TALBE VARCHAR2(30);
  L_TABLE      VARCHAR2(30);
  --L_SQL_SEG1,L_SQL_SEG2用于拼接INSERT语句
  L_SQL_SEG1 VARCHAR2(2000);
  L_SQL_SEG2 VARCHAR2(2000);
  --L_SQL_SEG3,L_SQL_SEG4用于拼接UPDATE语句
  L_SQL_SEG3 VARCHAR2(2000);
  L_SQL_SEG4 VARCHAR2(2000);
  --L_SQL_SEG5用于拼接DELETE语句(仅用于编辑模式)
  L_SQL_SEG5 VARCHAR2(2000);
  L_PK_DATA  VARCHAR2(1000);
  L_COLS_IDX NUMBER :=0;
  BEGIN
    DCM_COMMON.FETCH_TEMPLATE_INFO(P_TEMPLATE_ID,P_LOCALE,L_TEMP_TALBE,L_TABLE);
    IF P_HANDLE_MODE='REPLACE' THEN
      IF P_COM_RECORD_ID IS NULL THEN
        EXECUTE IMMEDIATE 'DELETE FROM "'||L_TABLE||'"';
      ELSE
        EXECUTE IMMEDIATE 'DELETE FROM "'||L_TABLE||'" WHERE COM_RECORD_ID='''||P_COM_RECORD_ID||'''';
      END IF;
    END IF;
    --INSERT
    L_SQL_SEG1:='INSERT INTO "'||L_TABLE||'"(';
    L_SQL_SEG1:=L_SQL_SEG1||'COM_RECORD_ID,CREATED_BY,UPDATED_BY,CREATED_AT,UPDATED_AT,IDX';
    L_SQL_SEG2:=' SELECT '''||P_COM_RECORD_ID||''','''||P_USER_ID||''','''||P_USER_ID||'''';
    L_SQL_SEG2:=L_SQL_SEG2||',SYSDATE,SYSDATE,T.IDX';
    --UPDATE
    L_SQL_SEG3:='UPDATE "'||L_TABLE||'" A SET(UPDATED_BY,UPDATED_AT,IDX';
    L_SQL_SEG4:='(SELECT '''||P_USER_ID||''',SYSDATE,T.IDX';
    --拼接字段列表
    FOR T_COL IN L_TEMPLATE_COL_CURSOR(P_TEMPLATE_ID,P_LOCALE) LOOP
      L_COLS_IDX:=L_COLS_IDX+1;
      IF T_COL.READONLY='N' THEN
        L_SQL_SEG1:=L_SQL_SEG1||',"'||T_COL.DB_TABLE_COL||'"';
        L_SQL_SEG3:=L_SQL_SEG3||',"'||T_COL.DB_TABLE_COL||'"';
        IF T_COL.DATA_TYPE='NUMBER' THEN
           L_SQL_SEG2:=L_SQL_SEG2||',TO_NUMBER(T.COLUMN'||L_COLS_IDX||')';
           L_SQL_SEG4:=L_SQL_SEG4||',TO_NUMBER(T.COLUMN'||L_COLS_IDX||')';
        ELSIF T_COL.DATA_TYPE='DATE' THEN
           L_SQL_SEG2:=L_SQL_SEG2||',TO_DATE(T.COLUMN'||L_COLS_IDX||',''YYYY-MM-DD HH24:MI:SS'')';
           L_SQL_SEG4:=L_SQL_SEG4||',TO_DATE(T.COLUMN'||L_COLS_IDX||',''YYYY-MM-DD HH24:MI:SS'')';
        ELSE
           L_SQL_SEG2:=L_SQL_SEG2||',T.COLUMN'||L_COLS_IDX;
           L_SQL_SEG4:=L_SQL_SEG4||',T.COLUMN'||L_COLS_IDX;
        END IF;
      END IF; 
      IF T_COL.IS_PK='Y' THEN
        L_PK_DATA:=L_PK_DATA||' AND A."'||T_COL.DB_TABLE_COL||'"=T.COLUMN'||L_COLS_IDX;
      END IF;
    END LOOP;
    --INSERT
    L_SQL_SEG1:=L_SQL_SEG1||')';
    L_SQL_SEG2:=L_SQL_SEG2||' FROM "'||L_TEMP_TALBE||'" T';
    L_SQL_SEG2:=L_SQL_SEG2||' WHERE T.TEMPLATE_ID='''||P_TEMPLATE_ID||''''; 
    --UPDATE 
    L_SQL_SEG3:=L_SQL_SEG3||')='; 
    L_SQL_SEG4:=L_SQL_SEG4||' FROM "'||L_TEMP_TALBE||'" T';
    L_SQL_SEG4:=L_SQL_SEG4||' WHERE T.TEMPLATE_ID='''||P_TEMPLATE_ID||''''; 
    
    IF P_COM_RECORD_ID IS NOT NULL THEN
       L_SQL_SEG2:=L_SQL_SEG2||' AND T.COM_RECORD_ID='''||P_COM_RECORD_ID||'''';
       L_SQL_SEG4:=L_SQL_SEG4||' AND T.COM_RECORD_ID='''||P_COM_RECORD_ID||'''';
    ELSE
       L_SQL_SEG2:=L_SQL_SEG2||' AND T.COM_RECORD_ID IS NULL';
       L_SQL_SEG4:=L_SQL_SEG4||' AND T.COM_RECORD_ID IS NULL';
    END IF;
    --增量模式且存在主键
    IF P_HANDLE_MODE='INCREMENT' AND L_PK_DATA IS NOT NULL THEN
       --仅插入不存在的数据
       L_SQL_SEG2:=L_SQL_SEG2||' AND NOT EXISTS (SELECT 1 FROM "'||L_TABLE||'" A WHERE 1=1';
       L_SQL_SEG2:=L_SQL_SEG2||L_PK_DATA;
       --根据主键更新已经存在的数据
       L_SQL_SEG4:=L_SQL_SEG4||L_PK_DATA||')';
       L_SQL_SEG4:=L_SQL_SEG4||' WHERE EXISTS(SELECT 1 FROM "'||L_TEMP_TALBE||'" T WHERE';
       IF P_COM_RECORD_ID IS NOT NULL THEN
          L_SQL_SEG2:=L_SQL_SEG2||' AND A.COM_RECORD_ID='''||P_COM_RECORD_ID||''')';
          L_SQL_SEG4:=L_SQL_SEG4||' T.COM_RECORD_ID='''||P_COM_RECORD_ID||'''';
          L_SQL_SEG4:=L_SQL_SEG4||' AND T.TEMPLATE_ID='''||P_TEMPLATE_ID||'''';
          L_SQL_SEG4:=L_SQL_SEG4||L_PK_DATA;
          L_SQL_SEG4:=L_SQL_SEG4||') AND A.COM_RECORD_ID='''||P_COM_RECORD_ID||'''';
       ELSE
          L_SQL_SEG2:=L_SQL_SEG2||'  AND A.COM_RECORD_ID IS NULL)';
          L_SQL_SEG4:=L_SQL_SEG4||'  T.COM_RECORD_ID IS NULL';
          L_SQL_SEG4:=L_SQL_SEG4||'  AND T.TEMPLATE_ID='''||P_TEMPLATE_ID||'''';
          L_SQL_SEG4:=L_SQL_SEG4||L_PK_DATA;
          L_SQL_SEG4:=L_SQL_SEG4||') AND A.COM_RECORD_ID IS NULL';       
       END IF;
    END IF;
    --若为编辑模式则根据ROWID更新数据,仅更新标记为UPDATE的数据,插入时仅插入标记为NEW的数据
    IF P_HANDLE_MODE='EDIT' THEN
      --INSERT
      L_SQL_SEG2:=L_SQL_SEG2||' AND T.EDIT_TYPE=''NEW''';
      --UPDATE
      L_SQL_SEG4:=L_SQL_SEG4||' AND T.ORIGIN_ROWID IS NOT NULL AND T.ORIGIN_ROWID=A.ROWID) WHERE ';
      L_SQL_SEG4:=L_SQL_SEG4||' EXISTS (SELECT 1 FROM "'||L_TEMP_TALBE||'" T WHERE ';
      L_SQL_SEG4:=L_SQL_SEG4||' T.TEMPLATE_ID='''||P_TEMPLATE_ID||'''';
      --DELETE
      L_SQL_SEG5:='DELETE FROM "'||L_TABLE||'" A WHERE';
      IF P_COM_RECORD_ID IS NOT NULL THEN
        L_SQL_SEG4:=L_SQL_SEG4||' AND T.COM_RECORD_ID='''||P_COM_RECORD_ID||'''';
        L_SQL_SEG4:=L_SQL_SEG4||' AND T.ORIGIN_ROWID=A.ROWID AND T.EDIT_TYPE=''UPDATE'')';
        L_SQL_SEG4:=L_SQL_SEG4||' AND A.COM_RECORD_ID='''||P_COM_RECORD_ID||'''';
        --DELETE
        L_SQL_SEG5:=L_SQL_SEG5||' A.COM_RECORD_ID='''||P_COM_RECORD_ID||''' AND EXISTS(';
        L_SQL_SEG5:=L_SQL_SEG5||' SELECT 1 FROM "'||L_TEMP_TALBE||'" T WHERE T.TEMPLATE_ID='''||P_TEMPLATE_ID||'''';
        L_SQL_SEG5:=L_SQL_SEG5||' AND T.COM_RECORD_ID='''||P_COM_RECORD_ID||'''';
        L_SQL_SEG5:=L_SQL_SEG5||' AND T.ORIGIN_ROWID=A.ROWID AND T.EDIT_TYPE=''DELETE'')';
      ELSE
        L_SQL_SEG4:=L_SQL_SEG4||' AND T.COM_RECORD_ID IS NULL';
        L_SQL_SEG4:=L_SQL_SEG4||' AND T.ORIGIN_ROWID=A.ROWID AND T.EDIT_TYPE=''UPDATE'')';
        L_SQL_SEG4:=L_SQL_SEG4||' AND A.COM_RECORD_ID IS NULL';
        --DELETE
        L_SQL_SEG5:=L_SQL_SEG5||' A.COM_RECORD_ID IS NULL AND EXISTS(';
        L_SQL_SEG5:=L_SQL_SEG5||' SELECT 1 FROM "'||L_TEMP_TALBE||'" T WHERE T.TEMPLATE_ID='''||P_TEMPLATE_ID||'''';
        L_SQL_SEG5:=L_SQL_SEG5||' AND T.COM_RECORD_ID IS NULL';
        L_SQL_SEG5:=L_SQL_SEG5||' AND T.ORIGIN_ROWID=A.ROWID AND T.EDIT_TYPE=''DELETE'')';
      END IF;
    END IF;
    --若为编辑模式（编辑模式存在数据的删除操作，且删除应该在插入和更新前）
    IF P_HANDLE_MODE='EDIT' THEN
      EXECUTE IMMEDIATE (L_SQL_SEG5);
    END IF;
    --执行数据插入操作（三种模式均存在数据插入操作）
    EXECUTE IMMEDIATE (L_SQL_SEG1||L_SQL_SEG2);
    --若为增量模式且模版定义了主键或者为编辑模式则执行更新操作
    IF (P_HANDLE_MODE='INCREMENT' AND L_PK_DATA IS NOT NULL) OR P_HANDLE_MODE='EDIT' THEN
      EXECUTE IMMEDIATE (L_SQL_SEG3||L_SQL_SEG4);
    END IF;
  END HANDLE_PROGRAM;
  /*****************************************************
  *获取模版信息
  *P_TEMPLATE_ID 模版ID
  *P_TEMP_TALBE  临时表
  *P_SRC_TABLE   原表
  *P_LOCALE    语言编码
  ******************************************************/
  PROCEDURE FETCH_TEMPLATE_INFO(
    P_TEMPLATE_ID     IN  VARCHAR2,
    P_LOCALE          IN  VARCHAR2,
    P_TEMP_TALBE      OUT VARCHAR2,
    P_SRC_TABLE       OUT VARCHAR2
  )
  IS
  BEGIN
    SELECT T.TMP_TABLE,T.DB_TABLE INTO P_TEMP_TALBE,P_SRC_TABLE 
    FROM DCM_TEMPLATE T 
    WHERE T.ID=P_TEMPLATE_ID
          AND T.LOCALE=P_LOCALE;
  END FETCH_TEMPLATE_INFO;
  /*****************************************************
  *空值校验
  *P_VALIDATION_ID 校验程序ID
  *P_TEMPLATE_ID 模版ID
  *P_COM_RECORD_ID 组合ID
  *P_TEMP_COL 校验程序对应临时表中的列
  *P_SRC_COL  校验程序对应目标表中的列
  *P_MODE     当前数据处理模式（REPLACE,INCREMENT,EDIT）
  *P_LOCALE    语言编码
  *P_ARGS      校验程序参数
  *P_FLAG      校验结果标识(Y,N)
  ******************************************************/
  PROCEDURE VALIDATE_NULL(P_VALIDATION_ID IN VARCHAR,
                          P_TEMPLATE_ID   IN VARCHAR2,
                          P_COM_RECORD_ID IN VARCHAR2,
                          P_TEMP_COL      IN VARCHAR2,
                          P_SRC_COL       IN VARCHAR2,
                          P_MODE          IN VARCHAR2,
                          P_LOCALE        IN VARCHAR2,
                          P_ARGS          IN VARCHAR2,
                          P_FLAG          OUT VARCHAR2) IS
    L_TABLE        VARCHAR2(32);
    L_TMP_TABLE    VARCHAR2(32);
    L_SQL_EN       VARCHAR2(2000);
    L_SQL_ZHS      VARCHAR2(2000);
    L_COL_ZHS_TEXT VARCHAR2(100);
    L_COL_EN_TEXT  VARCHAR2(100);
  BEGIN
    P_FLAG := 'Y';
    DCM_COMMON.FETCH_TEMPLATE_INFO(P_TEMPLATE_ID,
                                   P_LOCALE,
                                   L_TMP_TABLE,
                                   L_TABLE);
    SELECT T.COLUMN_LABEL
      INTO L_COL_ZHS_TEXT
      FROM DCM_TEMPLATE_COLUMN T
     WHERE T.TEMPLATE_ID = P_TEMPLATE_ID
       AND T.LOCALE = 'zh_CN'
       AND T.DB_TABLE_COL = P_SRC_COL;
    SELECT T.COLUMN_LABEL
      INTO L_COL_EN_TEXT
      FROM DCM_TEMPLATE_COLUMN T
     WHERE T.TEMPLATE_ID = P_TEMPLATE_ID
       AND T.LOCALE = 'en'
       AND T.DB_TABLE_COL = P_SRC_COL;
    --中文
    L_SQL_ZHS := 'INSERT INTO DCM_ERROR(ID,TEMPLATE_ID,COM_RECORD_ID,SHEET_NAME';
    L_SQL_ZHS := L_SQL_ZHS ||
                 ',ROW_NUM,MSG,ELEVEL,LOCALE,VALIDATION_ID,CREATED_AT,UPDATED_AT,CREATED_BY,UPDATED_BY)';
    L_SQL_ZHS := L_SQL_ZHS ||
                 ' SELECT DCM_SEQ.NEXTVAL,T.TEMPLATE_ID,'''||P_COM_RECORD_ID||''',T.SHEET_NAME';
    L_SQL_ZHS := L_SQL_ZHS || ',T.ROW_NO';
    L_SQL_ZHS := L_SQL_ZHS || ',''' || L_COL_ZHS_TEXT || '不能为空''' ||
                 ',''错误'',''zh_CN''';
    L_SQL_ZHS := L_SQL_ZHS || ',''' || P_VALIDATION_ID || '''';
    L_SQL_ZHS := L_SQL_ZHS || ',SYSDATE,SYSDATE,T.CREATED_BY,T.UPDATED_BY';
    L_SQL_ZHS := L_SQL_ZHS || ' FROM ' || L_TMP_TABLE ||
                 ' T WHERE T.TEMPLATE_ID=''' || P_TEMPLATE_ID || '''';
    IF P_COM_RECORD_ID IS NOT NULL THEN
      L_SQL_ZHS := L_SQL_ZHS || ' AND T.COM_RECORD_ID=''' ||
                   P_COM_RECORD_ID || '''';
    ELSE
      L_SQL_ZHS := L_SQL_ZHS || ' AND T.COM_RECORD_ID IS NULL';
    END IF;
    L_SQL_ZHS := L_SQL_ZHS || ' AND T.' || P_TEMP_COL || ' IS NULL';
    --英文
    L_SQL_EN := 'INSERT INTO DCM_ERROR(ID,TEMPLATE_ID,COM_RECORD_ID,SHEET_NAME';
    L_SQL_EN := L_SQL_EN ||
                ',ROW_NUM,MSG,ELEVEL,LOCALE,VALIDATION_ID,CREATED_AT,UPDATED_AT,CREATED_BY,UPDATED_BY)';
    L_SQL_EN := L_SQL_EN ||
                ' SELECT DCM_SEQ.NEXTVAL,T.TEMPLATE_ID,'''||P_COM_RECORD_ID||''',T.SHEET_NAME';
    L_SQL_EN := L_SQL_EN || ',T.ROW_NO';
    L_SQL_EN := L_SQL_EN || ',''' || L_COL_EN_TEXT || ' can not be blank''' ||
                ',''Error'',''en''';
    L_SQL_EN := L_SQL_EN || ',''' || P_VALIDATION_ID || '''';
    L_SQL_EN := L_SQL_EN || ',SYSDATE,SYSDATE,T.CREATED_BY,T.UPDATED_BY';
    L_SQL_EN := L_SQL_EN || ' FROM ' || L_TMP_TABLE ||
                ' T WHERE T.TEMPLATE_ID=''' || P_TEMPLATE_ID || '''';
    IF P_COM_RECORD_ID IS NOT NULL THEN
      L_SQL_EN := L_SQL_EN || ' AND T.COM_RECORD_ID=''' || P_COM_RECORD_ID || '''';
    ELSE
      L_SQL_EN := L_SQL_EN || ' AND T.COM_RECORD_ID IS NULL';
    END IF;
    L_SQL_EN := L_SQL_EN || ' AND T.' || P_TEMP_COL || ' IS NULL';
   
    EXECUTE IMMEDIATE L_SQL_ZHS;
    IF SQL%FOUND THEN
      P_FLAG := 'N';
      EXECUTE IMMEDIATE L_SQL_EN;
    END IF;
  END VALIDATE_NULL;
  
  /**************************************************
  *善后处理程序，用于在数据导入到真实表后进行操作
  *P_TEMPLATE_ID 模版ID
  *P_COM_RECORD_ID 组合ID
  *P_USER_ID  用户ID
  *P_HANDLE_MODE 处理模式，REPLACE表示替换，INCREMENT表示增量
  *P_LOCALE 语言编码
  ***************************************************/
  PROCEDURE AFTER_PROGRAM(
       P_TEMPLATE_ID   IN VARCHAR2,
       P_COM_RECORD_ID IN VARCHAR2,
       P_USER_ID       IN VARCHAR2,
       P_HANDLE_MODE   IN VARCHAR2,
       P_LOCALE        IN VARCHAR2
  )IS
  BEGIN
    DBMS_OUTPUT.PUT_LINE('PLEASE FLOW THIS DEMO TO IMPLEMENT THE ACTRAL LOGIC');
  END AFTER_PROGRAM;
END DCM_COMMON;
/
drop table ODI11_AGENT cascade constraints;

drop table ODI11_SCENE cascade constraints;

drop table ODI11_PARAMETER cascade constraints;

drop table ODI11_SCENE_PARAM cascade constraints;

drop table ODI11_SCENE_EXEC cascade constraints;

drop table ODI11_SCENE_CAT cascade constraints;

drop table ODI11_WORKREP cascade constraints;

drop table ODI11_ROLE_SCENE cascade constraints;

drop table ODI11_SCENE_LOG cascade constraints;

drop VIEW ODI11_USER_SCENE_V;

/*==============================================================*/
/* Table: ODI11_AGENT                                           */
/*==============================================================*/
create table ODI11_AGENT 
(
   ID                   VARCHAR2(32)         not null,
   AGENT_ALIAS          VARCHAR2(100)        not null,
   AGENT_NAME           VARCHAR2(100)        not null,
   AGENT_HOST           VARCHAR2(100)        not null,
   AGENT_CONTEXT        VARCHAR2(100)        not null,
   AGENT_PORT           VARCHAR2(10)         not null,
   PROTOCOL             VARCHAR2(10)         not null,
   LOCALE               VARCHAR2(10)         not null,
   CREATED_AT           DATE,
   UPDATED_AT           DATE,
   UPDATED_BY           VARCHAR2(32),
   CREATED_BY           VARCHAR2(32),
   constraint PK_ODI11_AGENT primary key (ID,LOCALE)
);

comment on table ODI11_AGENT is
'ODI11G代理';

comment on column ODI11_AGENT.ID is
'标识ID';

comment on column ODI11_AGENT.AGENT_ALIAS is
'代理别名';

comment on column ODI11_AGENT.AGENT_NAME is
'代理名称';

comment on column ODI11_AGENT.AGENT_HOST is
'代理主机';

comment on column ODI11_AGENT.AGENT_CONTEXT is
'代理上下文';

comment on column ODI11_AGENT.AGENT_PORT is
'代理端口';

comment on column ODI11_AGENT.PROTOCOL is
'代理协议';

comment on column ODI11_AGENT.LOCALE is
'多语言标识';

comment on column ODI11_AGENT.CREATED_AT is
'创建时间';

comment on column ODI11_AGENT.UPDATED_AT is
'更新时间';

comment on column ODI11_AGENT.UPDATED_BY is
'更新者';

comment on column ODI11_AGENT.CREATED_BY is
'创建者';

/*==============================================================*/
/* Table: ODI11_SCENE                                           */
/*==============================================================*/
create table ODI11_SCENE
(
   ID                   VARCHAR2(32)         not null,
   SCENE_ALIAS          VARCHAR2(100)        not null,
   SCENE_NAME           VARCHAR2(100)        not null,
   SCENE_VERSION        VARCHAR2(100)        not null,
   AGENT_ID             VARCHAR2(32)         ,
   WORKREP_ID           VARCHAR2(32)         not null,    
   SCENE_CONTEXT        VARCHAR2(100)        not null,
   CAT_ID               VARCHAR2(32)         not null,
   LOCALE               VARCHAR2(10)         not null,
   CREATED_AT           DATE,
   UPDATED_AT           DATE,
   UPDATED_BY           VARCHAR2(32),
   CREATED_BY           VARCHAR2(32),
   constraint PK_ODI11_SCENE primary key (ID,LOCALE)
);

comment on table ODI11_SCENE is
'ODI11G场景';

comment on column ODI11_SCENE.ID is
'ID标识';

comment on column ODI11_SCENE.SCENE_ALIAS is
'场景别名';

comment on column ODI11_SCENE.SCENE_NAME is
'场景名称';

comment on column ODI11_SCENE.AGENT_ID is
'代理ID';

comment on column ODI11_SCENE.WORKREP_ID is
'资料库ID';

comment on column ODI11_SCENE.SCENE_VERSION is
'场景版本';

comment on column ODI11_SCENE.SCENE_CONTEXT is
'执行上下文';

comment on column ODI11_SCENE.CAT_ID is
'分类ID';

comment on column ODI11_SCENE.LOCALE is
'语言标识';

comment on column ODI11_SCENE.CREATED_AT is
'创建时间';

comment on column ODI11_SCENE.UPDATED_AT is
'更新时间';

comment on column ODI11_SCENE.UPDATED_BY is
'更新者';

comment on column ODI11_SCENE.CREATED_BY is
'创建者';

/*==============================================================*/
/* Table: ODI11_PARAMETER                                       */
/*==============================================================*/
create table ODI11_PARAMETER
(
   ID                   VARCHAR2(32)         not null,
   P_ALIAS              VARCHAR2(100)        not null,
   P_NAME               VARCHAR2(100)        not null,
   VALUE_SET_ID         VARCHAR2(32)                 ,
   LOCALE               VARCHAR2(10)         not null,
   CREATED_AT           DATE,
   UPDATED_AT           DATE,
   UPDATED_BY           VARCHAR2(32),
   CREATED_BY           VARCHAR2(32),
   constraint PK_ODI11_PARAMETER primary key (ID,LOCALE)
);

comment on table ODI11_PARAMETER is
'ODI11G参数';

comment on column ODI11_PARAMETER.ID is
'ID标识';

comment on column ODI11_PARAMETER.P_ALIAS is
'参数别名';

comment on column ODI11_PARAMETER.P_NAME is
'参数名称';

comment on column ODI11_PARAMETER.VALUE_SET_ID is
'值集编码';

comment on column ODI11_PARAMETER.LOCALE is
'多语言标识';

comment on column ODI11_PARAMETER.CREATED_AT is
'创建时间';

comment on column ODI11_PARAMETER.UPDATED_AT is
'更新时间';

comment on column ODI11_PARAMETER.UPDATED_BY is
'更新者';

comment on column ODI11_PARAMETER.CREATED_BY is
'创建者';

/*==============================================================*/
/* Table: ODI11_SCENE_PARAM                                     */
/*==============================================================*/
create table ODI11_SCENE_PARAM
(
   ID                   VARCHAR2(32)         not null,
   SCENE_ID             VARCHAR2(32)         not null,
   PARAMETER_ID         VARCHAR2(32)         not null,
   IDX                  NUMBER,
   CREATED_AT           DATE,
   UPDATED_AT           DATE,
   UPDATED_BY           VARCHAR2(32),
   CREATED_BY           VARCHAR2(32),
   constraint PK_ODI11_SCENE_PARAM primary key (ID)
);

comment on table ODI11_SCENE_PARAM is
'ODI11G场景参数';

comment on column ODI11_SCENE_PARAM.ID is
'ID标识';

comment on column ODI11_SCENE_PARAM.SCENE_ID is
'场景ID';

comment on column ODI11_SCENE_PARAM.PARAMETER_ID is
'参数ID';

comment on column ODI11_SCENE_PARAM.IDX is
'序号';

comment on column ODI11_SCENE_PARAM.CREATED_AT is
'创建时间';

comment on column ODI11_SCENE_PARAM.UPDATED_AT is
'更新时间';

comment on column ODI11_SCENE_PARAM.UPDATED_BY is
'更新者';

comment on column ODI11_SCENE_PARAM.CREATED_BY is
'创建者';
/*==============================================================*/
/* Table: ODI11_SCENE_CAT                                       */
/*==============================================================*/
create table ODI11_SCENE_CAT
(
   ID                   VARCHAR2(32)         not null,
   P_ID                 VARCHAR2(32),
   CAT_NAME             VARCHAR2(150)        not null,
   LOCALE               VARCHAR2(10)         not null,
   IDX                  NUMBER,
   CREATED_AT           DATE,
   UPDATED_AT           DATE,
   UPDATED_BY           VARCHAR2(32),
   CREATED_BY           VARCHAR2(32),
   constraint PK_ODI11_SCENE_CAT primary key (ID,LOCALE)
);

comment on table ODI11_SCENE_CAT is
'ODI11G场景分类';

comment on column ODI11_SCENE_CAT.ID is
'ID标识';

comment on column ODI11_SCENE_CAT.CAT_NAME is
'分类名称';

comment on column ODI11_SCENE_CAT.LOCALE is
'多语言标识';

comment on column ODI11_SCENE_CAT.IDX is
'序号';

comment on column ODI11_SCENE_CAT.CREATED_AT is
'创建时间';

comment on column ODI11_SCENE_CAT.UPDATED_AT is
'更新时间';

comment on column ODI11_SCENE_CAT.UPDATED_BY is
'更新者';

comment on column ODI11_SCENE_CAT.CREATED_BY is
'创建者';

/*==============================================================*/
/* Table: ODI11_SCENE_CAT                                       */
/*==============================================================*/
create table ODI11_SCENE_EXEC
(
   ID                   VARCHAR2(32)         not null,
   SCENE_ID             VARCHAR2(32)         not null,
   PARAMS               VARCHAR2(1000),
   EXEC_STATUS          VARCHAR2(100),
   SESSION_NUM          VARCHAR2(100), 
   LOG_TEXT             VARCHAR2(1000), 
   HAS_EXCEPTION        CHAR(1),
   CREATED_AT           DATE,
   UPDATED_AT           DATE,
   UPDATED_BY           VARCHAR2(32),
   CREATED_BY           VARCHAR2(32),
   constraint PK_ODI11_SCENE_EXEC primary key (ID)
);

comment on table ODI11_SCENE_EXEC is
'ODI11场景执行';

comment on column ODI11_SCENE_EXEC.ID is
'ID标识';

comment on column ODI11_SCENE_EXEC.SCENE_ID is
'场景ID';

comment on column ODI11_SCENE_EXEC.PARAMS is
'参数';

comment on column ODI11_SCENE_EXEC.EXEC_STATUS is
'执行状态';

comment on column ODI11_SCENE_EXEC.SESSION_NUM is
'会话ID';

comment on column ODI11_SCENE_EXEC.LOG_TEXT is
'执行日志';

comment on column ODI11_SCENE_EXEC.CREATED_AT is
'创建时间';

comment on column ODI11_SCENE_EXEC.UPDATED_AT is
'更新时间';

comment on column ODI11_SCENE_EXEC.UPDATED_BY is
'更新者';

comment on column ODI11_SCENE_EXEC.CREATED_BY is
'创建者';

comment on column ODI11_SCENE_EXEC.HAS_EXCEPTION is
'有异常数据';

/*==============================================================*/
/* Table: ODI11_WORKREP                                         */
/*==============================================================*/
create table ODI11_WORKREP
(
   ID                   VARCHAR2(32)         not null,
   REP_ALIAS            VARCHAR2(100)        not null,
   REP_NAME             VARCHAR2(100)        not null,
   LOGIN_USER           VARCHAR2(100)        not null,
   LOGIN_PWD            VARCHAR2(255)        not null,
   LOCALE               VARCHAR2(10)         not null,      
   CREATED_AT           DATE,
   UPDATED_AT           DATE,
   UPDATED_BY           VARCHAR2(32),
   CREATED_BY           VARCHAR2(32),
   constraint PK_ODI11_WORKREP primary key (ID,LOCALE)
);

comment on table ODI11_WORKREP is
'工作资料库';

comment on column ODI11_WORKREP.ID is
'ID标识';

comment on column ODI11_WORKREP.REP_ALIAS is
'资料库别名';

comment on column ODI11_WORKREP.REP_NAME is
'资料库名称';

comment on column ODI11_WORKREP.LOGIN_USER is
'登录用户';

comment on column ODI11_WORKREP.LOGIN_PWD is
'登录密码';

comment on column ODI11_WORKREP.LOCALE is
'多语言标识';

comment on column ODI11_WORKREP.CREATED_AT is
'创建时间';

comment on column ODI11_WORKREP.UPDATED_AT is
'更新时间';

comment on column ODI11_WORKREP.UPDATED_BY is
'更新者';

comment on column ODI11_WORKREP.CREATED_BY is
'创建者';

/*==============================================================*/
/* Table: ODI11_ROLE_SCENE                                          */
/*==============================================================*/
create table ODI11_ROLE_SCENE 
(
   ID                   VARCHAR2(32)         not null,
   ROLE_ID              VARCHAR2(32)         not null,
   SCENE_ID             VARCHAR2(32)         not null,
   CREATED_AT           DATE,
   UPDATED_AT           DATE,
   UPDATED_BY           VARCHAR2(32),
   CREATED_BY           VARCHAR2(32),
   constraint PK_ODI11_ROLE_SCENE primary key (ID)
);

comment on table ODI11_ROLE_SCENE is
'ODI11G角色场景权限';

comment on column ODI11_ROLE_SCENE.ID is
'标识ID';

comment on column ODI11_ROLE_SCENE.ROLE_ID is
'角色ID';

comment on column ODI11_ROLE_SCENE.SCENE_ID is
'场景ID';

comment on column ODI11_ROLE_SCENE.CREATED_AT is
'创建时间';

comment on column ODI11_ROLE_SCENE.UPDATED_AT is
'更新时间';

comment on column ODI11_ROLE_SCENE.UPDATED_BY is
'更新者';

comment on column ODI11_ROLE_SCENE.CREATED_BY is
'创建者';
/*==============================================================*/
/* Table: ODI11_SCENE_LOG                                       */
/*==============================================================*/
create table ODI11_SCENE_LOG
(
  ID          VARCHAR2(32) not null,
  SESSION_NUM VARCHAR2(32) not null,
  MSG         VARCHAR2(4000),
  constraint ODI11_SCENE_LOG_PK primary key (ID)
);
comment on table ODI11_SCENE_LOG is
'ODI11G接口运行日志';
comment on column ODI11_SCENE_LOG.ID is
'ID标识';
comment on column ODI11_SCENE_LOG.SESSION_NUM is
'接口运行ID';
comment on column ODI11_SCENE_LOG.MSG is
'日志信息';

/*==============================================================*/
/* View: ODI11_USER_SCENE_V                                     */
/*==============================================================*/
CREATE VIEW ODI11_USER_SCENE_V AS
SELECT DISTINCT UG.USER_ID, RS.SCENE_ID
  FROM ODI11_ROLE_SCENE RS,
       DMS_GROUP_ROLE   GR,
       DMS_USER_GROUP   UG,
       DMS_GROUP        GRP,
       DMS_ROLE         RL
 WHERE RS.ROLE_ID = GR.ROLE_ID
   AND GR.GROUP_ID = UG.GROUP_ID
   AND RS.ROLE_ID = RL.ID
   AND UG.GROUP_ID = GRP.ID
   AND RL.ENABLE_FLAG = 'Y'
   AND GRP.ENABLE_FLAG = 'Y';