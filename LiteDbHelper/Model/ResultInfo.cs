/***
*	Title："轻量数据库" 项目
*		主题：结果信息
*	Description：
*		功能：XXX
*	Date：2022
*	Version：0.1版本
*	Author：Coffee
*	Modify Recoder：
*/

using System;
using System.Collections.Generic;
using System.Text;

namespace LiteDBHelper.Model
{
    public class ResultInfo
    {
        /// <summary>
        /// 结果状态
        /// </summary>
        public ResultStatus ResultStatus { get; set; }

        /// <summary>
        /// 消息（正常、错误、异常消息）
        /// </summary>
        public string Message { get; set; }

        /// <summary>
        /// 结果
        /// </summary>
        public string Result { get; set; }

        /// <summary>
        /// 构造函数
        /// </summary>
        public ResultInfo()
        {

        }

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="resultStatus">结果状态</param>
        /// <param name="message">消息（正常、错误、异常消息）</param>
        /// <param name="result">结果</param>
        public ResultInfo(ResultStatus resultStatus,string message,string result)
        {
            this.ResultStatus = resultStatus;
            this.Message = message;
            this.Result = result;
        }

        /// <summary>
        /// 设置内容
        /// </summary>
        /// <param name="resultStatus">结果状态</param>
        /// <param name="message">消息（正常、错误、异常消息）</param>
        /// <param name="result">结果</param>
        public void SetContent(ResultStatus resultStatus, string message, string result)
        {
            this.ResultStatus = resultStatus;
            this.Message = message;
            this.Result = result;
        }

        /// <summary>
        /// 结果打印
        /// </summary>
        /// <returns></returns>
        public string ResultPrint()
        {
            string str = $"状态：{ResultStatus} 消息：{Message}";
            if (!string.IsNullOrEmpty(Result))
            {
                str += $" 结果：{Result}";
            }
            return str;
        }


    }//Class_end

}
