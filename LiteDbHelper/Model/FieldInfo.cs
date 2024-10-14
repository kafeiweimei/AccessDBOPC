/***
*	Title："轻量数据库" 项目
*		主题：字段信息
*	Description：
*		功能：
*	Date：2021
*	Version：0.1版本
*	Author：Coffee
*	Modify Recoder：
*/

using System;
using System.Collections.Generic;
using System.Text;

namespace LiteDBHelper.Model
{
    public class FieldInfo
    {
        /// <summary>
        /// 字段名称
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 字段数据类型
        /// </summary>
        public ADOX.DataTypeEnum DataType { get; set; }

        /// <summary>
        /// 字段长度
        /// </summary>
        public int Length { get; set; }

        /// <summary>
        /// 是否为自增字段（默认不知自增字段）
        /// </summary>
        public bool IsAutoIncrement { get; set; } = false;

        /// <summary>
        /// 是否为主键(默认不是主键)
        /// </summary>
        public bool IsPrimaryKey { get; set; } = false;


        /// <summary>
        /// 构造函数
        /// </summary>
        public FieldInfo()
        {

        }


        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="name">字段名称</param>
        /// <param name="dataType">字段数据类型</param>
        /// <param name="length">字段长度</param>
        /// <param name="isAutoIncrement">是否为自增字段（默认不是）</param>
        /// <param name="isPrimaryKey">是否为主键（默认不是）</param>
        public FieldInfo(string name, ADOX.DataTypeEnum dataType,int length,bool isAutoIncrement=false,bool isPrimaryKey=false)
        {
            this.Name = name;
            this.DataType = dataType;
            this.Length = length;
            this.IsAutoIncrement = isAutoIncrement;
            this.IsPrimaryKey = isPrimaryKey;

        }

    }//Class_end
}
