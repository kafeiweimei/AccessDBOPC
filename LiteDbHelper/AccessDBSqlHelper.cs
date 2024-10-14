/***
*	Title："轻量数据库" 项目
*		主题：Access数据库的sql帮助类
*	Description：
*		功能：sql的常用功能
*	Date：2021
*	Version：0.1版本
*	Author：Coffee
*	Modify Recoder：
*/

using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Text;

namespace LiteDBHelper
{
    public class AccessDBSqlHelper
    {
        #region   基础参数
        //数据库连接字符串
        private string _ConnStr;

        #endregion

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="connectionStr">数据库连接字符串</param>
        public AccessDBSqlHelper(string connectionStr)
        {
            _ConnStr = connectionStr;
        }


        #region 公用方法

        public int GetMaxID(string tableName, string fieldName)
        {
            string strsql = "select max(" + fieldName + ")+1 from " + tableName;
            object obj = GetSingle(strsql);
            if (obj == null)
            {
                return 1;
            }
            else
            {
                return int.Parse(obj.ToString());
            }
        }

        public bool Exists(string strSql)
        {
            object obj = GetSingle(strSql);
            int cmdresult;
            if ((Object.Equals(obj, null)) || (Object.Equals(obj, System.DBNull.Value)))
            {
                cmdresult = 0;
            }
            else
            {
                cmdresult = int.Parse(obj.ToString());
            }
            if (cmdresult == 0)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        public bool Exists(string strSql, params OleDbParameter[] cmdParms)
        {
            object obj = GetSingle(strSql, cmdParms);
            int cmdresult;
            if ((Object.Equals(obj, null)) || (Object.Equals(obj, System.DBNull.Value)))
            {
                cmdresult = 0;
            }
            else
            {
                cmdresult = int.Parse(obj.ToString());
            }
            if (cmdresult == 0)
            {
                return false;
            }
            else
            {
                return true;
            }
        }
        #endregion


        #region 执行简单SQL语句

        /// <summary>
        /// 执行SQL语句，返回影响的记录数
        /// </summary>
        /// <param name="sql">SQL语句</param>
        /// <returns>影响的记录数</returns>
        public int ExecuteSql(string sql)
        {
            using (OleDbConnection connection = new OleDbConnection(_ConnStr))
            {
                using (OleDbCommand cmd = new OleDbCommand(sql, connection))
                {
                    try
                    {
                        connection.Open();
                        int rows = cmd.ExecuteNonQuery();
                        return rows;
                    }
                    catch (OleDbException ex)
                    {
                        connection.Close();
                        throw new Exception(ex.Message);
                    }
                }
            }
        }

        /// <summary>
        /// 执行SQL语句(设置命令的执行等待时间)
        /// </summary>
        /// <param name="sql"></param>
        /// <param name="Times"></param>
        /// <returns></returns>
        public int ExecuteSqlByWaitTime(string sql, int Times)
        {
            using (OleDbConnection connection = new OleDbConnection(_ConnStr))
            {
                using (OleDbCommand cmd = new OleDbCommand(sql, connection))
                {
                    try
                    {
                        connection.Open();
                        cmd.CommandTimeout = Times;
                        int rows = cmd.ExecuteNonQuery();
                        return rows;
                    }
                    catch (OleDbException ex)
                    {
                        connection.Close();
                        throw new Exception(ex.Message);
                    }
                }
            }
        }


        /// <summary>
        /// 执行多条SQL语句(通过事务方式)
        /// </summary>
        /// <param name="sqlList">sql语句列表</param>
        public void ExecuteSqlByTransaction(List<string> sqlList)
        {
            using (OleDbConnection conn = new OleDbConnection(_ConnStr))
            {
                conn.Open();
                OleDbCommand cmd = new OleDbCommand();
                cmd.Connection = conn;
                OleDbTransaction ta = conn.BeginTransaction();
                cmd.Transaction = ta;
                try
                {
                    for (int i = 0; i < sqlList.Count; i++)
                    {
                        string strSql = sqlList[i].ToString();
                        if (strSql.Trim().Length > 1)
                        {
                            cmd.CommandText = strSql;
                            cmd.ExecuteNonQuery();
                        }
                    }
                    ta.Commit();
                }
                catch (OleDbException ex)
                {
                    ta.Rollback();
                    throw new Exception(ex.Message);
                }
                
            }
        }


        /// <summary>
        /// 向数据库里插入图像格式的字段(和上面情况类似的另一种实例)
        /// </summary>
        /// <param name="strSQL">SQL语句</param>
        /// <param name="fs">图像字节,数据库的字段类型为image的情况</param>
        /// <returns>影响的记录数</returns>
        public int ExecuteSqlInsertImage(string sql, byte[] fs)
        {
            using (OleDbConnection connection = new OleDbConnection(_ConnStr))
            {
                OleDbCommand cmd = new OleDbCommand(sql, connection);
                OleDbParameter myParameter = new OleDbParameter("@fs", System.Data.SqlDbType.Image);
                myParameter.Value = fs;
                cmd.Parameters.Add(myParameter);
                try
                {
                    connection.Open();
                    int rows = cmd.ExecuteNonQuery();
                    return rows;
                }
                catch (OleDbException ex)
                {
                    throw new Exception(ex.Message);
                }
                finally
                {
                    cmd.Dispose();
                    connection.Close();
                }
            }
        }

        /// <summary>
        /// 执行一条计算查询结果语句，返回查询结果（object）。
        /// </summary>
        /// <param name="sql">计算查询结果语句</param>
        /// <returns>查询结果（object）</returns>
        public object ExecuteScalar(string sql)
        {
            using (OleDbConnection connection = new OleDbConnection(_ConnStr))
            {
                using (OleDbCommand cmd = new OleDbCommand(sql, connection))
                {
                    try
                    {
                        connection.Open();
                        object obj = cmd.ExecuteScalar();
                        if ((Object.Equals(obj, null)) || (Object.Equals(obj, System.DBNull.Value)))
                        {
                            return null;
                        }
                        else
                        {
                            return obj;
                        }
                    }
                    catch (OleDbException ex)
                    {
                        connection.Close();
                        throw new Exception(ex.Message);
                    }
                }
            }
        }

        /// <summary>
        /// 执行查询语句，返回SqlDataReader(使用该方法必须手动关闭SqlDataReader和连接)
        /// </summary>
        /// <param name="strSQL">查询语句</param>
        /// <returns>SqlDataReader</returns>
        public OleDbDataReader ExecuteReader(string sql)
        {
            OleDbConnection connection = new OleDbConnection(_ConnStr);
            OleDbCommand cmd = new OleDbCommand(sql, connection);
            try
            {
                connection.Open();
                OleDbDataReader myReader = cmd.ExecuteReader();
                return myReader;
            }
            catch (OleDbException ex)
            {
                throw new Exception(ex.Message);
            }
            //finally //不能在此关闭，否则，返回的对象将无法使用
            //{
            // cmd.Dispose();
            // connection.Close();
            //}
        }


        /// <summary>
        /// 执行一个查询,并返回查询结果
        /// </summary>
        /// <param name="sql">要执行的SQL语句</param>
        /// <returns></returns>
        public DataTable ExecuteDataTable(string sql)
        {
            DataTable data = new DataTable();//实例化DataTable，用于装载查询结果集
            using (OleDbConnection connection = new OleDbConnection(_ConnStr))
            {
                using (OleDbCommand command = new OleDbCommand(sql, connection))
                {
                    //通过包含查询SQL的SqlCommand实例来实例化SqlDataAdapter
                    OleDbDataAdapter adapter = new OleDbDataAdapter(command);

                    adapter.Fill(data);//填充DataTable
                }
            }
            return data;
        }

        /// <summary>
        /// 执行查询语句，返回DataSet
        /// </summary>
        /// <param name="sql">查询语句</param>
        /// <returns>DataSet</returns>
        public DataSet Query(string sql)
        {
            using (OleDbConnection connection = new OleDbConnection(_ConnStr))
            {
                DataSet ds = new DataSet();
                try
                {
                    connection.Open();
                    OleDbDataAdapter command = new OleDbDataAdapter(sql, connection);
                    command.Fill(ds, "ds");
                }
                catch (OleDbException ex)
                {
                    throw new Exception(ex.Message);
                }
                return ds;
            }
        }

        /// <summary>
        /// 执行查询语句，返回DataSet,设置命令的执行等待时间
        /// </summary>
        /// <param name="SQLString"></param>
        /// <param name="Times"></param>
        /// <returns></returns>
        public DataSet Query(string SQLString, int Times)
        {
            using (OleDbConnection connection = new OleDbConnection(_ConnStr))
            {
                DataSet ds = new DataSet();
                try
                {
                    connection.Open();
                    OleDbDataAdapter command = new OleDbDataAdapter(SQLString, connection);
                    command.SelectCommand.CommandTimeout = Times;
                    command.Fill(ds, "ds");
                }
                catch (System.Data.OleDb.OleDbException ex)
                {
                    throw new Exception(ex.Message);
                }
                return ds;
            }
        }

        #endregion


        #region 执行带参数的SQL语句

        /// <summary>
        /// 执行SQL语句，返回影响的记录数
        /// </summary>
        /// <param name="sql">SQL语句</param>
        /// <returns>影响的记录数</returns>
        public int ExecuteSql(string sql, params OleDbParameter[] cmdParms)
        {
            using (OleDbConnection connection = new OleDbConnection(_ConnStr))
            {
                using (OleDbCommand cmd = new OleDbCommand())
                {
                    try
                    {
                        PrepareCommand(cmd, connection, null, sql, cmdParms);
                        int rows = cmd.ExecuteNonQuery();
                        cmd.Parameters.Clear();
                        return rows;
                    }
                    catch (System.Data.OleDb.OleDbException E)
                    {
                        throw new Exception(E.Message);
                    }
                }
            }
        }

        /// <summary>
        /// 执行多条SQL语句，实现数据库事务。
        /// </summary>
        /// <param name="SQLStringList">SQL语句的哈希表（key为sql语句，value是该语句的OleDbParameter[]）</param>
        public void ExecuteSqlTran(Hashtable sqlList)
        {
            using (OleDbConnection conn = new OleDbConnection(_ConnStr))
            {
                conn.Open();
                using (OleDbTransaction trans = conn.BeginTransaction())
                {
                    OleDbCommand cmd = new OleDbCommand();
                    try
                    {
                        //循环
                        foreach (DictionaryEntry myDE in sqlList)
                        {
                            string cmdText = myDE.Key.ToString();
                            OleDbParameter[] cmdParms = (OleDbParameter[])myDE.Value;
                            PrepareCommand(cmd, conn, trans, cmdText, cmdParms);
                            int val = cmd.ExecuteNonQuery();
                            cmd.Parameters.Clear();
                            trans.Commit();
                        }
                    }
                    catch
                    {
                        trans.Rollback();
                        throw;
                    }
                }
            }
        }

        /// <summary>
        /// 执行一条计算查询结果语句，返回查询结果（object）。
        /// </summary>
        /// <param name="sql">计算查询结果语句</param>
        /// <returns>查询结果（object）</returns>
        public object GetSingle(string sql, params OleDbParameter[] cmdParms)
        {
            using (OleDbConnection connection = new OleDbConnection(_ConnStr))
            {
                using (OleDbCommand cmd = new OleDbCommand())
                {
                    try
                    {
                        PrepareCommand(cmd, connection, null, sql, cmdParms);
                        object obj = cmd.ExecuteScalar();
                        cmd.Parameters.Clear();
                        if ((Object.Equals(obj, null)) || (Object.Equals(obj, System.DBNull.Value)))
                        {
                            return null;
                        }
                        else
                        {
                            return obj;
                        }
                    }
                    catch (System.Data.OleDb.OleDbException ex)
                    {
                        throw new Exception(ex.Message);
                    }
                }
            }
        }


        /// <summary>
        /// 执行查询语句，返回SqlDataReader (使用该方法切记要手工关闭SqlDataReader和连接)
        /// </summary>
        /// <param name="sql">查询语句</param>
        /// <returns>SqlDataReader</returns>
        public OleDbDataReader ExecuteReader(string sql, params OleDbParameter[] cmdParms)
        {
            OleDbConnection connection = new OleDbConnection(_ConnStr);
            OleDbCommand cmd = new OleDbCommand();
            try
            {
                PrepareCommand(cmd, connection, null, sql, cmdParms);
                OleDbDataReader myReader = cmd.ExecuteReader();
                cmd.Parameters.Clear();
                return myReader;
            }
            catch (System.Data.OleDb.OleDbException e)
            {
                throw new Exception(e.Message);
            }
            //finally //不能在此关闭，否则，返回的对象将无法使用
            //{
            // cmd.Dispose();
            // connection.Close();
            //}
        }

        /// <summary>
        /// 执行一个查询,并返回查询结果
        /// </summary>
        /// <param name="sql">要执行的SQL语句</param>
        /// <param name="commandType">要执行的查询语句的类型，如存储过程或者SQL文本命令</param>
        /// <param name="parameters">Transact-SQL 语句或存储过程的参数数组</param>
        /// <returns></returns>
        public DataTable ExecuteDataTable(string sql, CommandType commandType, OleDbParameter[] parameters)
        {
            DataTable data = new DataTable();//实例化DataTable，用于装载查询结果集
            using (OleDbConnection connection = new OleDbConnection(_ConnStr))
            {
                using (OleDbCommand command = new OleDbCommand(sql, connection))
                {
                    command.CommandType = commandType;//设置command的CommandType为指定的CommandType
                    //如果同时传入了参数，则添加这些参数
                    if (parameters != null)
                    {
                        foreach (OleDbParameter parameter in parameters)
                        {
                            command.Parameters.Add(parameter);
                        }
                    }
                    //通过包含查询SQL的SqlCommand实例来实例化SqlDataAdapter
                    OleDbDataAdapter adapter = new OleDbDataAdapter(command);

                    adapter.Fill(data);//填充DataTable
                }
            }
            return data;
        }

        /// <summary>
        /// 执行查询语句，返回DataSet
        /// </summary>
        /// <param name="sql">查询语句</param>
        /// <returns>DataSet</returns>
        public DataSet Query(string sql, params OleDbParameter[] cmdParms)
        {
            using (OleDbConnection connection = new OleDbConnection(_ConnStr))
            {
                OleDbCommand cmd = new OleDbCommand();
                PrepareCommand(cmd, connection, null, sql, cmdParms);
                using (OleDbDataAdapter da = new OleDbDataAdapter(cmd))
                {
                    DataSet ds = new DataSet();
                    try
                    {
                        da.Fill(ds, "ds");
                        cmd.Parameters.Clear();
                    }
                    catch (System.Data.OleDb.OleDbException ex)
                    {
                        throw new Exception(ex.Message);
                    }
                    return ds;
                }
            }
        }

        /// <summary>
        /// 准备命令
        /// </summary>
        /// <param name="cmd"></param>
        /// <param name="conn"></param>
        /// <param name="trans"></param>
        /// <param name="cmdText"></param>
        /// <param name="cmdParms"></param>
        private void PrepareCommand(OleDbCommand cmd, OleDbConnection conn, OleDbTransaction trans, string cmdText, OleDbParameter[] cmdParms)
        {
            if (conn.State != ConnectionState.Open)
                conn.Open();
            cmd.Connection = conn;
            cmd.CommandText = cmdText;
            if (trans != null)
                cmd.Transaction = trans;
            cmd.CommandType = CommandType.Text;//cmdType;
            if (cmdParms != null)
            {
                foreach (OleDbParameter parameter in cmdParms)
                {
                    if ((parameter.Direction == ParameterDirection.InputOutput || parameter.Direction == ParameterDirection.Input) &&
                      (parameter.Value == null))
                    {
                        parameter.Value = DBNull.Value;
                    }
                    cmd.Parameters.Add(parameter);
                }
            }
        }

        #endregion


        #region 获取根据指定字段排序并分页查询。

        /// <summary>
        ///获取到分页内容
        /// </summary>
        /// <param name="pageIndex">当前页码</param>
        /// <param name="pageSize">分页容量</param>
        /// <param name="primaryKey">主键</param>
        /// <param name="needShowFields">显示的字段</param>
        /// <param name="tableName">表名称</param>
        /// <param name="whereString">查询条件，若有条件限制则必须以where 开头</param>
        /// <param name="orderString">排序规则</param>
        /// <param name="pageCount">传出参数：总页数统计</param>
        /// <param name="recordCount">传出参数：总记录统计</param>
        /// <returns>装载记录的DataTable</returns>
        public DataTable GetPageContent(int pageIndex, int pageSize, string primaryKey, string needShowFields, string tableName, string whereString, string orderString, out int pageCount, out int recordCount)
        {
            if (pageIndex < 1) pageIndex = 1;
            if (pageSize < 1) pageSize = 10;
            if (string.IsNullOrEmpty(needShowFields)) needShowFields = "*";
            if (string.IsNullOrEmpty(orderString)) orderString = primaryKey + " asc ";

            using (OleDbConnection m_Conn = new OleDbConnection(_ConnStr))
            {
                m_Conn.Open();
                string myVw = string.Format("{0} tempVw ", tableName);
                OleDbCommand cmdCount = new OleDbCommand(string.Format(" select count(*) as recordCount from {0} {1}", myVw, whereString), m_Conn);

                recordCount = Convert.ToInt32(cmdCount.ExecuteScalar());

                if ((recordCount % pageSize) > 0)
                    pageCount = recordCount / pageSize + 1;
                else
                    pageCount = recordCount / pageSize;
                OleDbCommand cmdRecord;
                if (pageIndex == 1)//第一页
                {
                    cmdRecord = new OleDbCommand(string.Format("select top {0} {1} from {2} {3} order by {4} ", pageSize, needShowFields, myVw, whereString, orderString), m_Conn);
                }
                else if (pageIndex > pageCount)//超出总页数
                {
                    cmdRecord = new OleDbCommand(string.Format("select top {0} {1} from {2} {3} order by {4} ", pageSize, needShowFields, myVw, "where 1=2", orderString), m_Conn);
                }
                else
                {
                    int pageLowerBound = pageSize * pageIndex;
                    int pageUpperBound = pageLowerBound - pageSize;
                    string recordIDs = recordID(string.Format("select top {0} {1} from {2} {3} order by {4} ", pageLowerBound, primaryKey, myVw, whereString, orderString), pageUpperBound);
                    cmdRecord = new OleDbCommand(string.Format("select {0} from {1} where {2} in ({3}) order by {4} ", needShowFields, myVw, primaryKey, recordIDs, orderString), m_Conn);

                }
                OleDbDataAdapter dataAdapter = new OleDbDataAdapter(cmdRecord);
                DataTable dt = new DataTable();
                dataAdapter.Fill(dt);
                m_Conn.Close();
                m_Conn.Dispose();
                return dt;
            }
        }

        /// <summary>
        /// 分页使用
        /// </summary>
        /// <param name="query"></param>
        /// <param name="passCount"></param>
        /// <returns></returns>
        private string recordID(string query, int passCount)
        {
            using (OleDbConnection m_Conn = new OleDbConnection(_ConnStr))
            {
                m_Conn.Open();
                OleDbCommand cmd = new OleDbCommand(query, m_Conn);
                string result = string.Empty;
                using (OleDbDataReader dr = cmd.ExecuteReader())
                {
                    while (dr.Read())
                    {
                        if (passCount < 1)
                        {
                            result += "," + dr.GetInt32(0);
                        }
                        passCount--;
                    }
                }
                m_Conn.Close();
                m_Conn.Dispose();
                return result.Substring(1);
            }
        }



        /// <summary>
        /// 获取到分页查询的sql
        /// </summary>
        /// <param name="primaryKey">主键（不能为空）</param>
        /// <param name="queryFields">提取字段（不能为空）</param>
        /// <param name="tableName">表（理论上允许多表）</param>
        /// <param name="condition">条件（可以空）</param>
        /// <param name="orderBy">排序，格式：字段名+""+ASC（可以空）</param>
        /// <param name="pageSize">分页数（不能为空）</param>
        /// <param name="pageIndex">当前页，起始为：1（不能为空）</param>
        /// <returns></returns>
        public static string GetPageListSql(string primaryKey, string queryFields, string tableName, string condition, string orderBy, int pageSize, int pageIndex)
        {
            string strTmp = ""; //---strTmp用于返回的SQL语句
            string SqlSelect = "", SqlPrimaryKeySelect = "", strOrderBy = "", strWhere = " where 1=1 ", strTop = "";
            //0：分页数量
            //1:提取字段
            //2:表
            //3:条件
            //4:主键不存在的记录
            //5:排序
            SqlSelect = " select top {0} {1} from {2} {3} {4} {5}";
            //0:主键
            //1:TOP数量,为分页数*(排序号-1)
            //2:表
            //3:条件
            //4:排序
            SqlPrimaryKeySelect = " and {0} not in (select {1} {0} from {2} {3} {4}) ";
            if (orderBy != "")
                strOrderBy = " order by " + orderBy;
            if (condition != "")
                strWhere += " and " + condition;
            int pageindexsize = (pageIndex - 1) * pageSize;
            if (pageindexsize > 0)
            {
                strTop = " top " + pageindexsize.ToString();
                SqlPrimaryKeySelect = String.Format(SqlPrimaryKeySelect, primaryKey, strTop, tableName, strWhere, strOrderBy);
                strTmp = String.Format(SqlSelect, pageSize.ToString(), queryFields, tableName, strWhere, SqlPrimaryKeySelect, strOrderBy);
            }
            else
            {
                strTmp = String.Format(SqlSelect, pageSize.ToString(), queryFields, tableName, strWhere, "", strOrderBy);
            }
            return strTmp;
        }


        /// <summary>
        /// 获取根据指定字段排序并分页查询。DataSet
        /// </summary>
        /// <param name="pageSize">每页要显示的记录的数目</param>
        /// <param name="pageIndex">要显示的页的索引</param>
        /// <param name="tableName">要查询的数据表</param>
        /// <param name="queryFields">要查询的字段,如果是全部字段请填写：*</param>
        /// <param name="primaryKey">主键字段，类似排序用到</param>
        /// <param name="orderBy">是否为升序排列：0为升序，1为降序</param>
        /// <param name="condition">查询的筛选条件</param>
        /// <returns>返回排序并分页查询的DataSet</returns>
        public DataSet GetPagingList(string primaryKey, string queryFields, string tableName, string condition, string orderBy, int pageSize, int pageIndex)
        {
            string sql = GetPageListSql(primaryKey, queryFields, tableName, condition, orderBy, pageSize, pageIndex);
            return Query(sql);
        }

        /// <summary>
        /// 获取根据指定字段排序并分页查询
        /// </summary>
        /// <param name="pageSize">每页要显示的记录的数目</param>
        /// <param name="pageIndex">要显示的页的索引</param>
        /// <param name="tableName">要查询的数据表</param>
        /// <param name="queryFields">要查询的字段,如果是全部字段请填写：*</param>
        /// <param name="primaryKey">主键字段，类似排序用到</param>
        /// <param name="orderBy">是否为升序排列：0为升序，1为降序</param>
        /// <param name="condition">查询的筛选条件</param>
        /// <returns>返回排序并分页查询的DataSet</returns>
        public DataTable GetPagingList2(string primaryKey, string queryFields, string tableName, string condition, string orderBy, int pageSize, int pageIndex)
        {
            string sql = GetPageListSql(primaryKey, queryFields, tableName, condition, orderBy, pageSize, pageIndex);
            return ExecuteDataTable(sql);
        }

        public string GetPagingListSQL(string primaryKey, string queryFields, string tableName, string condition, string orderBy, int pageSize, int pageIndex)
        {
            string sql = GetPageListSql(primaryKey, queryFields, tableName, condition, orderBy, pageSize, pageIndex);
            return sql;
        }


        #endregion


        #region   数据库的常用操作(获取所有表、表包含的所有列信息)

        /// <summary>
        /// 返回当前连接的数据库中的所有数据表信息
        /// </summary>
        /// <returns></returns>
        public DataTable GetAllTableInfo()
        {
            DataTable data = null;
            using (OleDbConnection connection = new OleDbConnection(_ConnStr))
            {
                connection.Open();//打开数据库连接
                data = connection.GetSchema("Tables");
            }
            return data;
        }

        /// <summary>
        /// 返回当前连接的数据库中所有由用户创建的所有数据表信息
        /// </summary>
        /// <returns></returns>
        public DataTable GetUserCreateAllTableInfo()
        {
            DataTable data = null;

            using (OleDbConnection connection = new OleDbConnection(_ConnStr))
            {
                connection.Open();//打开数据库连接
                data = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
            }

            return data;
        }

        /// <summary>
        /// 获取数据库中包含的所有表名称
        /// </summary>
        /// <param name="getAllTable">获取到的所有表</param>
        /// <returns></returns>
        public List<string> GetAllTableName(DataTable getAllTableInfo)
        {
            if (getAllTableInfo == null || getAllTableInfo.Rows.Count < 1) return null;

            List<string> tableNameList = new List<string>();
            for (int i = 0; i < getAllTableInfo.Rows.Count; i++)
            {
                string tmpTableName = getAllTableInfo.Rows[i]["TABLE_NAME"].ToString();
                tableNameList.Add(tmpTableName);
            }

            return tableNameList;
        }

        /// <summary>
        /// 获取到指定表的所有列信息
        /// </summary>
        /// <param name="tableName">表名称</param>
        /// <returns></returns>
        public DataTable GetAllColumnInfoOfTable(string tableName)
        {
            if (string.IsNullOrEmpty(tableName)) return null;

            DataTable data = null;

            using (OleDbConnection connection = new OleDbConnection(_ConnStr))
            {
                connection.Open();//打开数据库连接
                data = connection.GetSchema("columns", new string[] { null, null, tableName });
            }

            return data;
        }

        /// <summary>
        /// 获取表中包含的所有列名称
        /// </summary>
        /// <param name="getAllTable">获取到的所有表</param>
        /// <returns></returns>
        public List<string> GetAllColumnName(DataTable getAllColumnInfo)
        {
            if (getAllColumnInfo == null || getAllColumnInfo.Rows.Count < 1) return null;

            List<string> tableNameList = new List<string>();
            for (int i = 0; i < getAllColumnInfo.Rows.Count; i++)
            {
                string tmpTableName = getAllColumnInfo.Rows[i]["COLUMN_NAME"].ToString();
                tableNameList.Add(tmpTableName);
            }

            return tableNameList;
        }

        #endregion


    }//Class_end
}
