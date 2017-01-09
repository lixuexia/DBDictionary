using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;

namespace DBDictionary_OpenXml
{
    public class DBStructure
    {
        public static readonly string DBConn = System.Configuration.ConfigurationManager.ConnectionStrings["DBCONN"].ConnectionString;
        public static DataTable GetTables()
        {
            using (SqlConnection MyConn = new SqlConnection(DBConn))
            {
                SqlCommand MyCmd = new SqlCommand(
@"select t.id,t.name,d.[value] desctxt from sysobjects t left outer join sys.extended_properties d 
on t.id=d.major_id and d.minor_id=0
where t.xtype='U' and t.name not like '%.del' order by t.name asc", MyConn);
                SqlDataAdapter MyAdp = new SqlDataAdapter(MyCmd);
                DataTable DT = new DataTable();
                MyAdp.Fill(DT);
                return DT;
            }
        }

        public static DataTable GetTableInfo(string TableName)
        {
            using (SqlConnection MyConn = new SqlConnection(DBConn))
            {
                SqlCommand MyCmd = new SqlCommand(@"select a.name 列名,a.length 长度,case when a.isnullable=1 then '是' when a.isnullable=0 then '否' end 可空,
case when a.IsIdentity=1 then '是' when a.IsIdentity=0 then '否' end 自增,
case when a.PK=1 then '是' when a.PK=0 then '否' end  主键,b.value as 描述,c.name as 类型 from 
(
   select id,colid,name,xtype,length,colstat,autoval,isnullable,COLUMNPROPERTY(a.id,a.name,'IsIdentity') as IsIdentity,
   (SELECT count(*) FROM sysobjects WHERE (name in (SELECT name FROM sysindexes WHERE (id = a.id) AND
   (indid in (SELECT indid FROM sysindexkeys WHERE (id = a.id) AND (colid in (SELECT colid FROM syscolumns WHERE (id = a.id) AND (name = a.name))))))) AND (xtype = 'PK')) as PK
   from syscolumns as a where name<>'rowguid' and id in(select id from sysobjects where xtype='U' and name=@TabName)
) as a 
left outer join sys.extended_properties as b on (a.id=b.major_id and a.colid=b.minor_id)
left outer join systypes as c on (a.xtype=c.xtype and c.xtype=c.xusertype)",MyConn);
                MyCmd.Parameters.Add("@TabName", SqlDbType.NVarChar).Value = TableName;
                SqlDataAdapter MyAdp = new SqlDataAdapter(MyCmd);
                DataTable DT = new DataTable();
                MyAdp.Fill(DT);
                return DT;
            }
        }
    }
}