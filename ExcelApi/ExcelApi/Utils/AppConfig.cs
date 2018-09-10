using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;

/// <summary>
/// AppConfig 的摘要说明
/// </summary>
public class AppConfig
{
    /// <summary>
    /// 日志路径
    /// </summary>
    public static readonly string strLogPath = ConfigurationManager.AppSettings["LogPath"] == null ? "" : ConfigurationManager.AppSettings["LogPath"].ToString();

    /// <summary>
    /// 日志记录模式
    /// </summary>
    public static readonly string strLogLogMode = ConfigurationManager.AppSettings["LogMode"] == null ? "0" : ConfigurationManager.AppSettings["LogMode"].ToString();

    /// <summary>
    /// 错误日志文件名
    /// </summary>
    public static string strErrorFileName = "API_ERROR_" + DateTime.Today.ToString("yyyyMMdd") + ".log";

    /// <summary>
    /// 接口使用日志文件名
    /// </summary>
    public static string strInfoFileName = "API_INFO_" + DateTime.Today.ToString("yyyyMMdd") + ".log";

    private static string _encoding = ConfigurationManager.AppSettings["characterEncoding"] == null ? "UTF-8" : ConfigurationManager.AppSettings["characterEncoding"].ToString();
    
    /// <summary>
    /// 字符编码
    /// </summary>
    public static readonly System.Text.Encoding Encoding = System.Text.Encoding.GetEncoding(_encoding);

    /// <summary>
    /// 系统根路径
    /// </summary>
    public static string strAppPath = "~";

    /// <summary>
    /// 资源路径
    /// </summary>
    public static string dataPath = ConfigurationManager.AppSettings["dataPath"] == null ? "" : ConfigurationManager.AppSettings["dataPath"].ToString();
}