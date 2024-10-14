/***
*	Title："轻量数据库" 项目
*		主题：这个关于Access数据库的常用操作演示
*	Description：
*		功能：XXX
*	Date：2022
*	Version：0.1版本
*	Author：Coffee
*	Modify Recoder：
*/

using LiteDBHelper;
using LiteDBHelper.Model;
using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace Test_LiteDbOpc
{
    internal class AccessDBOpc
    {

        public static void Demo()
        {
            Console.WriteLine("Hello World!");

            #region   创建Access数据库
            Console.WriteLine($"---------------开始创建Access数据库---------------");

            string mdbFilePathAndName = $"E:\\AccessDB\\{DateTime.Now.ToString("yyyy-MM-dd_HHmmss")}.mdb";

            //创建无账号密码的mdb文件
            AccessDBHelper accessDBHelper1 = new AccessDBHelper(mdbFilePathAndName, AccessDBType.Is2007AndLater);
            ResultInfo result1 = accessDBHelper1.CreateMdbDataBase(mdbFilePathAndName);
            Console.WriteLine($"创建无账号密码的Access数据库情况：{result1.ResultPrint()}");

            ////创建带密码的mdb文件
            //MdbHelper mdbHelper2 = new MdbHelper(mdbFilePathAndName,"123456");
            //ResultInfo result2 = mdbHelper2.CreateMdbDataBase(mdbFilePathAndName);
            //Console.WriteLine($"创建带密码的AccessDB文件情况：{result2.ResultPrint()}");

            ////创建带用户密码的mdb文件
            //MdbHelper mdbHelper3 = new MdbHelper(mdbFilePathAndName,"admin" ,"123456");
            //ResultInfo result3 = mdbHelper3.CreateMdbDataBase(mdbFilePathAndName);
            //Console.WriteLine($"创建带用户密码的AccessDB文件情况：{result3.ResultPrint()}");

            #endregion


            #region   创建Access数据库的表
            Console.WriteLine($"\n---------------开始创建Access数据库的表---------------");

            //创建表方式1
            List<string> fieldNameList = new List<string>() { "Id", "Name", "Sex", "Age", "TelNumber", "Address" };
            ResultInfo result2 = accessDBHelper1.CreateMdbTable(mdbFilePathAndName, "UserInfo", fieldNameList);
            Console.WriteLine($"创建表情况：{result2.ResultPrint()}");

            //创建表方式2
            List<FieldInfo> fieldNameList2 = new List<FieldInfo>()
            {
                new FieldInfo("Id",  ADOX.DataTypeEnum.adInteger,8,true,true),
                new FieldInfo("ImageType",ADOX.DataTypeEnum.adVarWChar,16),
                new FieldInfo("AddDate",ADOX.DataTypeEnum.adDate,7),
                new FieldInfo("ImagePath",ADOX.DataTypeEnum.adVarWChar,255),
                new FieldInfo("IsDisable",ADOX.DataTypeEnum.adBoolean,5),
                new FieldInfo("Position",ADOX.DataTypeEnum.adDouble,20)

            };
            ResultInfo result2_2 = accessDBHelper1.CreateMdbTable("ImageInfo", fieldNameList2);

            Console.WriteLine($"创建Access数据库的表情况：{result2_2.ResultPrint()}");

            #endregion 

            #region   插入数据
            Console.WriteLine($"\n---------------开始执行插入数据---------------");

            //插入Access数据库中ImageInfo表数据
            string insertSql = null;
            List<string> sqlList = new List<string>();
            for (int i = 0; i < 10; i++)
            {

                insertSql = $"Insert Into ImageInfo ([ImageType],[AddDate],[ImagePath],[IsDisable],[Position]) " +
                   $"Values ('jpg','{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")}','E:\\MyProject\\Image\\{i}',{false},'{i}.25879456123');";

                sqlList.Add(insertSql);


                ////执行单条sql语句
                //int affectRow = accessDBHelper1.SqlHelper.ExecuteSql(insertSql);

                //Console.WriteLine($"{i} 插入数据行数：{affectRow}");
            }

            //执行多条sql语句（通过事务方式）
            accessDBHelper1.SqlHelper.ExecuteSqlByTransaction(sqlList);

            Console.WriteLine($"------------插入 {sqlList.Count} 条数据完成----------");
            #endregion

            #region   查询数据
            Console.WriteLine($"\n---------------开始查询所有数据---------------");

            string querySql = $"Select Id,ImageType,AddDate,ImagePath,IsDisable,[Position] From ImageInfo ";
            DataTable queryData = accessDBHelper1.SqlHelper.ExecuteDataTable(querySql);

            PrintDatatable(queryData);

            Console.WriteLine($"---------------查询数据完成---------------");

            #endregion

            #region   更新数据
            Console.WriteLine($"\n---------------开始更新数据---------------");

            string updateSql = $"update ImageInfo set ImageType='png',AddDate='2022-04-17 10:36:36'," +
                $"ImagePath='D:\\TestProject\\image\\11',IsDisable=true,[Position]='36.258941' where Id=1  ";
            int updateData = accessDBHelper1.SqlHelper.ExecuteSql(updateSql);

            Console.WriteLine($"---------------更新 {updateData} 条数据完成---------------");

            Console.WriteLine($"\n---------------更新数据后_开始查询所有数据---------------");

            string querySql2 = $"Select Id,ImageType,AddDate,ImagePath,IsDisable,[Position] From ImageInfo ";
            DataTable queryData2 = accessDBHelper1.SqlHelper.ExecuteDataTable(querySql2);

            PrintDatatable(queryData2);

            Console.WriteLine($"---------------更新数据后_查询数据完成---------------");

            #endregion

            #region   删除数据

            Console.WriteLine($"\n---------------开始删除数据---------------");

            string deleteSql = $"delete * from  ImageInfo where Id=10  ";
            int deleteData = accessDBHelper1.SqlHelper.ExecuteSql(deleteSql);

            Console.WriteLine($"---------------删除 {updateData} 条数据完成---------------");

            Console.WriteLine($"\n--------------删除数据后_开始查询所有数据---------------");

            string querySql3 = $"Select Id,ImageType,AddDate,ImagePath,IsDisable,[Position] From ImageInfo ";
            DataTable queryData3 = accessDBHelper1.SqlHelper.ExecuteDataTable(querySql3);

            PrintDatatable(queryData3);

            Console.WriteLine($"---------------删除数据后_查询数据完成---------------");

            #endregion

            #region   分页查询
            Console.WriteLine($"\n---------------开始分页查询数据---------------");

            int pageIndex = 2, pageSize = 5;

            DataTable pageDt = accessDBHelper1.SqlHelper.GetPageContent(pageIndex, pageSize, "Id", "Id,ImageType,AddDate,ImagePath,IsDisable,[Position]", "ImageInfo", " where ImageType='jpg'", " Id ASC ", out int pageCount, out int total);

            PrintDatatable(pageDt);

            Console.WriteLine($"---------------分页查询完成___当前查询第 {pageIndex} 页 ，每页 {pageSize} 条，共 {pageCount} 页、{total} 条数据---------------");
            #endregion

            #region   获取数据库中包含的所有表

            Console.WriteLine($"\n---------------开始获取数据库中用户创建的所有表名称及其包含的列名称---------------");

            DataTable getAllTableInfo = accessDBHelper1.SqlHelper.GetUserCreateAllTableInfo();

            List<string> getAllTableName = accessDBHelper1.SqlHelper.GetAllTableName(getAllTableInfo);
            if (getAllTableName != null && getAllTableName.Count > 0)
            {
                string tableName = null;
                string columnName = null;
                for (int i = 0; i < getAllTableName.Count; i++)
                {
                    tableName = getAllTableName[i];
                    Console.WriteLine($"\n表的名称为：{tableName}");

                    DataTable getAllColumnInfo = accessDBHelper1.SqlHelper.GetAllColumnInfoOfTable(tableName);
                    List<string> getAllColumnName = accessDBHelper1.SqlHelper.GetAllColumnName(getAllColumnInfo);
                    if (getAllColumnName != null && getAllColumnName.Count > 0)
                    {
                        Console.WriteLine($"{tableName}表包含的所有列名称为：");
                        for (int j = 0; j < getAllColumnName.Count; j++)
                        {
                            columnName = getAllColumnName[j];
                            Console.WriteLine($"列的名称为：{columnName}");
                        }
                    }

                }



            }

            Console.WriteLine($"--------------获取数据库中用户创建的所有表名称及其包含的列名称完成---------------");


            #endregion

            Console.ReadLine();
        }

        //输出DataTable内容
        private static void PrintDatatable(DataTable dt)
        {
            Console.WriteLine($"Id\tImageType\tAddDate\t\tImagePath\t\tIsDisable\tPosition\t");
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    object str = dt.Rows[i][j];

                    Console.Write($"{str}\t");
                }
                Console.WriteLine();
            }
        }



    }//Class_end

}
