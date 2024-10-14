/***
*	Title："轻量数据库" 项目
*		主题：Access数据库的帮助类
*	Description：
*		功能：
*		    ①构造函数时可以创建Access指定的连接字符串
*		    ②创建Access的mdb类型数据库
*		    ③创建Access数据库中的表
*		    ④给Access数据库的表添加指定类型的字段
*	Date：2022
*	Version：0.1版本
*	Author：Coffee
*	Modify Recoder：
*/

using LiteDBHelper.Model;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data.OleDb;
using System.IO;
using System.Text;

namespace LiteDBHelper
{
    public class AccessDBHelper
    {

        #region   基础参数
        //数据库连接字符串
        private string _ConnStr;

        //获取到数据库连接字符串
        public string ConnStr { get { return _ConnStr; } }

        //SqlHelper实例
        private AccessDBSqlHelper _SqlHelper;

        //获取到SqlHelper实例
        public AccessDBSqlHelper SqlHelper { get { return _SqlHelper; } }


        #endregion 


        #region   构造函数
        /// <summary>
        /// mdb文件的连接字符串构造函数
        /// </summary>
        /// <param name="connnection"></param>
        public AccessDBHelper(string connnection)
        {
            if (string.IsNullOrEmpty(connnection)) return;

            _ConnStr = connnection;

            InstanceSqlHelper(_ConnStr);
        }

        /// <summary>
        /// mdb文件无用户名和密码构造函数
        /// </summary>
        /// <param name="mdbFilePathAndName">mdb文件的路径和名称（比如：@"D:\\HalmEL\\2022-4-11.mdb"）</param>
        public AccessDBHelper(string mdbFilePathAndName,AccessDBType accessDBType)
        {
            if (string.IsNullOrEmpty(mdbFilePathAndName)) return;

            string strDriver = GetDirverOfAccessDBType(accessDBType);
            _ConnStr = $"{strDriver};Data source={mdbFilePathAndName};";

            InstanceSqlHelper(_ConnStr);
        }

        /// <summary>
        /// mdb文件有密码构造函数
        /// </summary>
        /// <param name="mdbFilePathAndName">mdb文件的路径和名称（比如：@"D:\\HalmEL\\2022-4-11.mdb"）</param>
        /// <param name="password">数据库密码</param>
        public AccessDBHelper(string mdbFilePathAndName, string password,AccessDBType accessDBType)
        {
            if (string.IsNullOrEmpty(mdbFilePathAndName) || string.IsNullOrEmpty(password)) return;

            string strDriver = GetDirverOfAccessDBType(accessDBType);
            _ConnStr = $"{strDriver};Data source={mdbFilePathAndName};Jet OleDb:DataBase Password={password};";

            InstanceSqlHelper(_ConnStr);
        }

        /// <summary>
        /// mdb文件有用户和密码构造函数
        /// </summary>
        /// <param name="mdbFilePathAndName">mdb文件的路径和名称（比如：@"D:\\HalmEL\\2022-4-11.mdb"）</param>
        /// <param name="userName">用户名称</param>
        /// <param name="password">用户密码</param>
        public AccessDBHelper(string mdbFilePathAndName, string userName, string password, AccessDBType accessDBType)
        {
            if (string.IsNullOrEmpty(mdbFilePathAndName) || string.IsNullOrEmpty(userName) || string.IsNullOrEmpty(password)) return;

            string strDriver = GetDirverOfAccessDBType(accessDBType);
            _ConnStr = $"{strDriver};Data Source={mdbFilePathAndName};User ID={userName};Password=;Jet OLEDB:Database Password={password}";

            InstanceSqlHelper(_ConnStr);
        }

        #endregion


        #region   创建Access数据库、表及其字段

        /// <summary>
        /// 创建Mdb数据库
        /// </summary>
        /// <param name="mdbFilePathAndName">mdb文件的路径和名称（比如：@"D:\\HalmEL\\2022-4-11.mdb"）</param>
        /// <returns>返回创建结果</returns>
        public ResultInfo CreateMdbDataBase(string mdbFilePathAndName)
        {
            ResultInfo resultInfo = new ResultInfo();

            if (File.Exists(mdbFilePathAndName))
            {
                resultInfo.ResultStatus = ResultStatus.Success;
                resultInfo.Message = $"{mdbFilePathAndName} 文件已经存在！";
            }
            try
            {
                //如果目录不存在，则创建目录
                string folder = Path.GetDirectoryName(mdbFilePathAndName);
                if (!Directory.Exists(folder))
                {
                    Directory.CreateDirectory(folder);
                }
                //创建Catalog目录类
                ADOX.CatalogClass catalog = new ADOX.CatalogClass();

                //根据联结字符串使用Jet数据库引擎创建数据库
                catalog.Create(_ConnStr);
                catalog = null;

                resultInfo.ResultStatus = ResultStatus.Success;
                resultInfo.Message = $"{mdbFilePathAndName} 文件创建成功！";
            }
            catch (Exception ex)
            {
                resultInfo.ResultStatus = ResultStatus.Error;
                resultInfo.Message = $"{ex.Message}";
            }

            return resultInfo;
        }

        /// <summary>
        /// 创建mdb表(字段都是短文本类型)
        /// </summary>
        /// <param name="mdbFilePathAndName">mdb文件的路径和名称（比如：@"D:\\HalmEL\\2022-4-11.mdb"）</param>
        /// <param name="tableName">表名称</param>
        /// <param name="fieldNameList">表字段名称列表</param>
        /// <returns></returns>
        public ResultInfo CreateMdbTable(string mdbFilePathAndName, string tableName, List<string> fieldNameList)
        {
            ResultInfo resultInfo = new ResultInfo();

            if (string.IsNullOrEmpty(mdbFilePathAndName) || string.IsNullOrEmpty(tableName)
                || fieldNameList == null || fieldNameList.Count < 1)
            {
                resultInfo.SetContent(ResultStatus.Error, "内容为空,请检查！", null);

                return resultInfo;
            }


            ADOX.CatalogClass catalog = new ADOX.CatalogClass();
            ADODB.Connection connection = new ADODB.Connection();

            try
            {
                //打开数据库连接
                connection.Open(_ConnStr, null, null, -1);
                catalog.ActiveConnection = connection;

                //新建一个表
                ADOX.TableClass table = new ADOX.TableClass();
                table.ParentCatalog = catalog;
                table.Name = tableName;

                int fieldCount = fieldNameList.Count;
                for (int i = 0; i < fieldCount; i++)
                {
                    //增加一个文本字段
                    string fieldName = fieldNameList[i].ToString();
                    ADOX.ColumnClass column = new ADOX.ColumnClass();
                    column.ParentCatalog = catalog;
                    column.Name = fieldName;
                    column.Properties["Jet OLEDB:Allow Zero Length"].Value = false;
                    //table.Columns.Append(column, ADOX.DataTypeEnum.adLongVarChar, 100);
                    table.Columns.Append(fieldName, ADOX.DataTypeEnum.adVarWChar, 106);

                }

                //将创建的表加入数据库
                catalog.Tables.Append(table);
                table = null;
                catalog = null;

                resultInfo.SetContent(ResultStatus.Success, $"创建：{tableName} 表成功", null);
            }
            catch (Exception ex)
            {
                resultInfo.SetContent(ResultStatus.Error, $"{ex.Message}", null);
            }
            finally
            {
                //关闭连接
                connection.Close();
            }

            return resultInfo;
        }

        /// <summary>
        /// 创建mdb表（可自定义字段类型）
        /// </summary>
        /// <param name="tableName">表名称</param>
        /// <param name="fieldList">字段列表</param>
        /// <returns></returns>
        public ResultInfo CreateMdbTable(string tableName, List<FieldInfo> fieldList)
        {
            ResultInfo resultInfo = new ResultInfo();

            if (string.IsNullOrEmpty(tableName) || fieldList == null || fieldList.Count < 1)
            {
                resultInfo.SetContent(ResultStatus.Error, "内容为空,请检查！", null);
                return resultInfo;
            }

            ADOX.CatalogClass catalog = new ADOX.CatalogClass();
            ADODB.Connection connection = new ADODB.Connection();

            try
            {
                //新建目录且打开数据库连接
                connection.Open(_ConnStr, null, null, -1);
                catalog.ActiveConnection = connection;

                //新建一个表
                ADOX.TableClass curTable = NewTable(catalog, tableName);

                //给表添加字段
                if (curTable != null && fieldList != null && fieldList.Count >= 1)
                {
                    int fieldCount = fieldList.Count;

                    FieldInfo fieldInfo = new FieldInfo();
                    for (int i = 0; i < fieldCount; i++)
                    {
                        fieldInfo = fieldList[i];
                        AddField(catalog, curTable, fieldList[i]);
                    }
                }

                //将创建的表加入数据库
                catalog.Tables.Append(curTable);
                curTable = null;
                catalog = null;


                resultInfo.SetContent(ResultStatus.Success, $"创建：{tableName} 表成功", null);
            }
            catch (Exception ex)
            {
                resultInfo.SetContent(ResultStatus.Error, $"{ex.Message}", null);
            }
            finally
            {
                //关闭连接
                connection.Close();
            }

            return resultInfo;
        }

        #endregion


        #region   执行sql语句

        #endregion 


        #region   私有方法

        /// <summary>
        /// 根据Access类型返回对应的驱动
        /// </summary>
        /// <param name="accessDBType">Access数据库类型</param>
        /// <returns></returns>
        private string GetDirverOfAccessDBType(AccessDBType accessDBType)
        {
            string connStr = $"Microsoft.ACE.OLEDB.12.0";

            switch (accessDBType)
            {
                case AccessDBType.Is2007AndLater:
                    connStr = $"Provider=Microsoft.ACE.OLEDB.12.0";
                    break;
                case AccessDBType.Is2003AndBefore:
                    connStr = $"Provider=Microsoft.Jet.OLEDB.4.0";
                    break;
                default:
                    break;
            }

            return connStr;
        }

        /// <summary>
        /// 实例化SqlHelper
        /// </summary>
        /// <param name="dbConnection">数据库连接字符串</param>
        private void InstanceSqlHelper(string dbConnection)
        {
            if (string.IsNullOrEmpty(dbConnection)) return;

            if (_SqlHelper == null)
            {
                _SqlHelper = new AccessDBSqlHelper(dbConnection);
            }
        }


        /// <summary>
        /// 新建表
        /// </summary>
        /// <param name="parentCatalog">父目录</param>
        /// <param name="tableName">表名称</param>
        /// <returns>返回新建的表</returns>
        private ADOX.TableClass NewTable(ADOX.CatalogClass parentCatalog, string tableName)
        {
            if (parentCatalog == null || string.IsNullOrEmpty(tableName)) return null;

            ADOX.TableClass table = new ADOX.TableClass();
            table.ParentCatalog = parentCatalog;
            table.Name = tableName;

            return table;
        }


        /// <summary>
        /// 添加字段
        /// </summary>
        /// <param name="parentCatalog">父目录</param>
        /// <param name="table">字段加到的表</param>
        /// <param name="fieldInfo">字段信息</param>
        /// <returns>返回添加字段结果</returns>
        private bool AddField(ADOX.CatalogClass parentCatalog, ADOX.TableClass table,
            FieldInfo fieldInfo)
        {
            bool result = false;

            if (parentCatalog == null || table == null || fieldInfo == null) return result;

            //增加指定类型的字段到表中
            ADOX.ColumnClass column = new ADOX.ColumnClass();
            column.ParentCatalog = parentCatalog;
            column.Name = fieldInfo.Name;
            column.Type = fieldInfo.DataType;
            column.DefinedSize = fieldInfo.Length;
            column.Properties["Jet OLEDB:Allow Zero Length"].Value = false;
            if (fieldInfo.IsAutoIncrement)
            {
                column.Properties["AutoIncrement"].Value = true;
            }
            if (fieldInfo.IsPrimaryKey)
            {
                table.Keys.Append("PrimaryKey", ADOX.KeyTypeEnum.adKeyPrimary, fieldInfo.Name, "", "");
            }

            table.Columns.Append(column, column.Type, column.DefinedSize);
            //table.Columns.Append(fieldInfo.Name, fieldInfo.DataType,fieldInfo.Length);

            result = true;

            return result;
        }

        #endregion 


    }//Class_end


    /// <summary>
    /// Access数据库类型
    /// </summary>
    public enum AccessDBType
    {
        //2007及其更高的版本
        Is2007AndLater,
        //2003等之前的版本
        Is2003AndBefore,

    }


}
