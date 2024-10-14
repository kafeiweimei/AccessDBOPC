/***
*	Title："轻量数据库" 项目
*		主题：获取
*	Description：
*		功能：XXX
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

namespace LiteDbHelper
{
    public class MdbHelper
    {
        //数据库连接资源
        private OleDbConnection _oleDbConnection;

        //数据库连接资源
        public OleDbConnection MdbConntion { get { return _oleDbConnection; } }


        //数据库连接字符串
        private string _ConnStr;

        //获取到数据库连接字符串
        public string ConnStr { get { return _ConnStr; } }

        /// <summary>
        /// mdb文件无用户名和密码构造函数
        /// </summary>
        /// <param name="mdbFilePathAndName">mdb文件的路径和名称（比如：@"D:\\HalmEL\\2022-4-11.mdb"）</param>
        public MdbHelper(string mdbFilePathAndName,AccessDBType accessDBType=AccessDBType.Is2007AndLater)
        {
            if (string.IsNullOrEmpty(mdbFilePathAndName)) return;

            string strDriver = GetDirverOfAccessDBType(accessDBType);
            _ConnStr = $"{strDriver}Data source={mdbFilePathAndName};";
        }

        /// <summary>
        /// mdb文件有密码构造函数
        /// </summary>
        /// <param name="mdbFilePathAndName">mdb文件的路径和名称（比如：@"D:\\HalmEL\\2022-4-11.mdb"）</param>
        /// <param name="password">数据库密码</param>
        public MdbHelper(string mdbFilePathAndName, string password,AccessDBType accessDBType=AccessDBType.Is2007AndLater)
        {
            if (string.IsNullOrEmpty(mdbFilePathAndName) || string.IsNullOrEmpty(password)) return;

            string strDriver = GetDirverOfAccessDBType(accessDBType);
            _ConnStr = $"{strDriver}Data source={mdbFilePathAndName};Jet OleDb:DataBase Password={password};";
        }

        /// <summary>
        /// mdb文件有用户和密码构造函数
        /// </summary>
        /// <param name="mdbFilePathAndName">mdb文件的路径和名称（比如：@"D:\\HalmEL\\2022-4-11.mdb"）</param>
        /// <param name="userName">用户名称</param>
        /// <param name="password">用户密码</param>
        public MdbHelper(string mdbFilePathAndName, string userName, string password, AccessDBType accessDBType = AccessDBType.Is2007AndLater)
        {
            if (string.IsNullOrEmpty(mdbFilePathAndName) || string.IsNullOrEmpty(userName) || string.IsNullOrEmpty(password)) return;

            string strDriver = GetDirverOfAccessDBType(accessDBType);
            _ConnStr = $"{strDriver}Data Source={mdbFilePathAndName};User ID={userName};Password=;Jet OLEDB:Database Password={password}";

          
        }

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

        //创建mdb表
        public ResultInfo CreateMdbTable(string mdbFilePathAndName,string tableName, ArrayList arrayList)
        {
            ResultInfo resultInfo = new ResultInfo();


            return resultInfo;
        }


        #region   私有方法

        /// <summary>
        /// 根据Access类型返回对应的驱动
        /// </summary>
        /// <param name="accessDBType">Access数据库类型</param>
        /// <returns></returns>
        private string GetDirverOfAccessDBType(AccessDBType accessDBType)
        {
            string connStr = $"Microsoft.ACE.OLEDB.12.0;";

            switch (accessDBType)
            {
                case AccessDBType.Is2007AndLater:
                    connStr = $"Provider=Microsoft.ACE.OLEDB.12.0;";
                    break;
                case AccessDBType.Is2003AndBefore:
                    connStr = $"Provider=Microsoft.Jet.OLEDB.4.0;";
                    break;
                default:
                    break;
            }

            return connStr;
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
