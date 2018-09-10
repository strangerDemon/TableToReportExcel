using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;

/// <summary>
/// ErrorUtil 的摘要说明
/// </summary>
public class logUtil
{
    #region 异常信息写入错误日志日志文件
    /// <summary>
    /// 异常信息写入日志文件
    /// </summary>
    /// <Coder>董彦雷：2016-6-12</Coder>
    /// <Modifier></Modifier>
    /// <param name="ex">异常变量</param>
    /// <returns>成功写入返回true，否则返回false</returns>
    public static bool RecordExceptionToFile(Exception ex)
    {
        try
        {
            if (ex.GetType().ToString() == "System.Threading.ThreadAbortException")
            {
                return false;
            }

            //取得当前需要写入的日志文件名称及路径
            string strFullPath = "";
            string logPath = "";


            logPath = HttpContext.Current.Server.MapPath(AppConfig.strLogPath);
            strFullPath = HttpContext.Current.Server.MapPath(AppConfig.strLogPath + @"\" + AppConfig.strErrorFileName);

            //取得异常信息的内容
            string logErrorInfo = GetLogInfo(ex);


            //执行写入
            //检查 Log 文件所存放的目录是否存在,如果不存在，建立该文件夹
            if (!Directory.Exists(logPath))
            {
                Directory.CreateDirectory(logPath);
            }

            //判断当前的日志文件是否创建，如果未创建，执行创建并加入异常内容；
            //如果已经创建则直接追加填写
            if (!File.Exists(strFullPath))
            {
                using (StreamWriter sw = File.CreateText(strFullPath))
                {
                    sw.Write(logErrorInfo);
                    sw.Flush();
                }
            }
            else
            {
                using (StreamWriter sw = File.AppendText(strFullPath))
                {
                    sw.Write(logErrorInfo);
                    sw.Flush();
                }
            }

            return true;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// 异常信息写入日志文件
    /// </summary>
    /// <param name="ex">异常信息</param>
    /// <param name="jo">相关内容</param>
    /// <param name="apiName">apiName</param>
    public static void RecordErrorToFile(Exception ex, JObject jo, string apiName)
    {
        try
        {
            errorModel objError = new errorModel();
            objError.apiName = apiName;
            objError.ex = ex;
            objError.para = jo;

            JObject joError = JObject.Parse(JsonConvert.SerializeObject(objError));

            //取得当前需要写入的日志文件名称及路径
            string strFullPath = "";
            string logPath = "";


            logPath = HttpContext.Current.Server.MapPath(AppConfig.strLogPath);
            strFullPath = HttpContext.Current.Server.MapPath(AppConfig.strLogPath + @"\" + AppConfig.strErrorFileName);

            //取得异常信息的内容
            string strTime = "\r\n------BEGIN----------------------------" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "------------------------------\r\n";
            string logErrorInfo = strTime + joError.ToString();
            logErrorInfo += ("\r\n------END-----------------------------------------------------------------------------\r\n");

            //执行写入
            //检查 Log 文件所存放的目录是否存在,如果不存在，建立该文件夹
            if (!Directory.Exists(logPath))
            {
                Directory.CreateDirectory(logPath);
            }

            //判断当前的日志文件是否创建，如果未创建，执行创建并加入异常内容；
            //如果已经创建则直接追加填写
            if (!File.Exists(strFullPath))
            {
                using (StreamWriter sw = File.CreateText(strFullPath))
                {
                    sw.Write(logErrorInfo);
                    sw.Flush();
                }
            }
            else
            {
                using (StreamWriter sw = File.AppendText(strFullPath))
                {
                    sw.Write(logErrorInfo);
                    sw.Flush();
                }
            }
        }
        catch
        {
            return;
        }
    }

    /// <summary>
    /// 异常信息写入日志文件
    /// </summary>
    /// <param name="ex">异常信息</param>
    /// <param name="jo">相关内容</param>
    /// <param name="ctx">请求上下文</param> 
    public static void RecordErrorToFile(Exception ex, string info, HttpContext ctx)
    {
        try
        {
            errorModel objError = new errorModel();
            objError.apiName = ctx.Request.Url.AbsoluteUri;

            objError.ex = ex;
            objError.para = info;

            JObject joError = JObject.Parse(JsonConvert.SerializeObject(objError));

            //取得当前需要写入的日志文件名称及路径
            string strFullPath = "";
            string logPath = "";

            
             logPath = HttpContext.Current.Server.MapPath(AppConfig.strLogPath);
             strFullPath = HttpContext.Current.Server.MapPath(AppConfig.strLogPath + @"\" + AppConfig.strErrorFileName);
            

            //取得异常信息的内容
            string strTime = "\r\n------BEGIN----------------------------" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "------------------------------\r\n";
            string logErrorInfo = strTime + joError.ToString();
            logErrorInfo += ("\r\n------END-----------------------------------------------------------------------------\r\n");

            //执行写入
            //检查 Log 文件所存放的目录是否存在,如果不存在，建立该文件夹
            if (!Directory.Exists(logPath))
            {
                Directory.CreateDirectory(logPath);
            }

            //判断当前的日志文件是否创建，如果未创建，执行创建并加入异常内容；
            //如果已经创建则直接追加填写
            if (!File.Exists(strFullPath))
            {
                using (StreamWriter sw = File.CreateText(strFullPath))
                {
                    sw.Write(logErrorInfo);
                    sw.Flush();
                }
            }
            else
            {
                using (StreamWriter sw = File.AppendText(strFullPath))
                {
                    sw.Write(logErrorInfo);
                    sw.Flush();
                }
            }
        }
        catch
        {
            return;
        }
    }
    #endregion

    #region 组织异常信息字符串
    /// <summary>
    /// 组织异常信息字符串
    /// </summary>
    /// <Coder>董彦雷：2016-6-12</Coder>
    /// <Modifier></Modifier>
    /// <param name="ex">异常变量</param>
    /// <returns>异常信息字符串</returns>
    private static string GetLogInfo(Exception ex)
    {
        try
        {
            string strNow = DateTime.Now.ToString("yyyyMMdd HH:mm:ss");
            StringBuilder sbLog = new StringBuilder();

            sbLog.Append("\r\n--------------------------------------------\r\n");
            sbLog.Append(strNow);
            sbLog.Append("\r\n\tSource:");
            sbLog.Append(ex.Source);
            sbLog.Append("\r\n\tMessage:");
            sbLog.Append(ex.Message);
            sbLog.Append("\r\n\tStackTrace:");
            sbLog.Append(ex.StackTrace);

            if (ex.InnerException != null)
            {
                sbLog.Append("\r\n\tInnerException:");
                sbLog.Append(ex.InnerException.StackTrace);
            }

            return sbLog.ToString();
        }
        catch
        {
            return string.Empty;
        }
    }
    #endregion

    #region 将文本信息写入日志
    /// <summary>
    /// 将文本信息写入日志
    /// </summary>
    /// <param name="info">文本信息</param>
    public static void RecordInfoToFile(string info)
    {
        try
        {
            //取得当前需要写入的日志文件名称及路径
            string strFullPath = "";
            string logPath = "";
            logPath = AppConfig.strLogPath + @"\OnlineUsersLog\";
            strFullPath = logPath + AppConfig.strErrorFileName;

            //取得异常信息的内容
            string strTime = "\r\n------BEGIN----------------------------" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "------------------------------\r\n";
            info = strTime + info;
            info += ("\r\n------END-----------------------------------------------------------------------------\r\n");


            //检查 Log 文件所存放的目录是否存在,如果不存在，建立该文件夹
            if (!Directory.Exists(logPath))
            {
                Directory.CreateDirectory(logPath);
            }

            if (!File.Exists(strFullPath))
            {
                using (StreamWriter sw = File.CreateText(strFullPath))
                {
                    sw.Write(info);
                    sw.Flush();
                }
            }
            else
            {
                using (StreamWriter sw = File.AppendText(strFullPath))
                {
                    sw.Write(info);
                    sw.Flush();
                }
            }
        }
        catch
        {
            return;
        }
    }
    #endregion

    #region 错误日志对象
    public class errorModel
    {
        /// <summary>
        /// 接口名称
        /// </summary>
        public string apiName
        {
            get;
            set;
        }

        /// <summary>
        /// 请求参数
        /// </summary>
        public object para
        {
            get;
            set;
        }

        /// <summary>
        /// 异常信息
        /// </summary>
        public Exception ex
        {
            get;
            set;
        }
    }
    #endregion

    #region 接口使用日志
    /// <summary>
    /// 接口使用日志-记录至文本文件
    /// </summary>
    /// <param name="jo">参数</param>
    /// <param name="result">执行结果信息</param>
    /// <param name="tokenInfo">token信息</param>
    /// <param name="ctx">上下文</param>
    public static void RecordInfoToFile(JObject jo, Object result, object tokenInfo, HttpContext ctx)
    {
        try
        {
            jo["DToken"] = JObject.FromObject(tokenInfo);
            jo["Result"] = JObject.FromObject(result);
            JObject joFile = new JObject();
            string[] urlSegments = HttpContext.Current.Request.Url.Segments;
            string interfaceName = "\"API Name\":\"" + urlSegments[urlSegments.Length - 1] + "\"";


            //取得当前需要写入的日志文件名称及路径
            string strFullPath = "";
            string logPath = "";

            if (ctx.Request.Url.AbsolutePath.Contains("xmtdtpc.asmx"))
            {
                logPath = AppConfig.strLogPath + @"\WEB";
                strFullPath = AppConfig.strLogPath + @"\WEB\" + AppConfig.strInfoFileName;
            }
            else
            {
                logPath = AppConfig.strLogPath + @"\APP";
                strFullPath = AppConfig.strLogPath + @"\APP\" + AppConfig.strInfoFileName;
            }

            //取得异常信息的内容
            string strTime = "\r\n------BEGIN----------------------------" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "------------------------------\r\n";
            string logInfo = strTime + interfaceName + "\r\n" + jo.ToString();
            logInfo += ("\r\n------END-----------------------------------------------------------------------------\r\n");

            //执行写入
            //检查 Log 文件所存放的目录是否存在,如果不存在，建立该文件夹
            if (!Directory.Exists(logPath))
            {
                Directory.CreateDirectory(logPath);
            }

            //判断当前的日志文件是否创建，如果未创建，执行创建并加入异常内容；
            //如果已经创建则直接追加填写
            if (!File.Exists(strFullPath))
            {
                using (StreamWriter sw = File.CreateText(strFullPath))
                {
                    sw.Write(logInfo);
                    sw.Flush();
                }
            }
            else
            {
                using (StreamWriter sw = File.AppendText(strFullPath))
                {
                    sw.Write(logInfo);
                    sw.Flush();
                }
            }
        }
        catch
        {
            return;
        }
    }

    /// <summary>
    /// 接口使用日志-存入数据库
    /// </summary>
    /// <param name="jo"></param>
    /// <param name="result"></param>
    /// <param name="tokenInfo"></param>
    /// <param name="ctx"></param>
    public static void RecordInfoToDb(JObject jo, Object result, object tokenInfo, HttpContext ctx)
    {
        try//// todo apiUser 
        {
            /*jo["DToken"] = JObject.FromObject(tokenInfo);
            jo["Result"] = JObject.FromObject(result);
            AppInterfaceLogModel model = new AppInterfaceLogModel();
            string[] urlSegments = HttpContext.Current.Request.Url.Segments;
            model.InterfaceName = urlSegments[urlSegments.Length - 1];
            model.Parameter = JObject.Parse(JsonConvert.SerializeObject(jo)).ToString();
            model.UseDateTime = DateTime.Now;

            if (ctx.Request.Url.AbsolutePath.Contains("xmtdtpc.asmx"))
            {
                model.PhoneUUID = "";
                model.DataSource = "web";
            }
            else
            {
                TokenModel tokenModel = tokenInfo as TokenModel;
                model.PhoneUUID = tokenModel.phoneUUID;
                model.DataSource = tokenModel.platform;
            }

            AppInterfaceLogBLL bll = new AppInterfaceLogBLL();
            bll.Add(model);*/
        }
        catch
        {
            return;
        }
    }
    #endregion
}