using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Script.Serialization;

/// <summary>
/// BusQueryUtil 的摘要说明
/// </summary>
public class BizApiUtil
{
    #region 配置文件
    public static JObject BaseConfigPara
    {
        get
        {
            string data = System.IO.File.ReadAllText(AppConfig.strAppPath + "DataConfig/TMapConfig/baseConfig.json");
            return JObject.Parse(data);
        }
    }

    /// <summary>
    /// 获取BaseConfigPara
    /// </summary>
    /// <returns></returns>
    public static JObject getBaseConfigPara()
    {
        return BaseConfigPara;
    }

    #endregion

    #region 公共函数
    public static JObject ApiRequest(string method,string para)
    {
        try
        {
            List<HttpParameter> paras = new List<HttpParameter>();
            paras.Add(new HttpParameter("para", para));
            string res = HttpProxysUtil.HttpPost(BaseConfigPara["apiBaseUrl"] + method, paras, "");
            JObject result = JObject.Parse(res);
            return result;
        }
        catch (Exception ex)
        {
            throw ex;
        }
    }
    #endregion

}