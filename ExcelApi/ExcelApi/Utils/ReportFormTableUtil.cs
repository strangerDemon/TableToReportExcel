using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.IO;
/// <summary>
/// SysConfigUtil 的摘要说明
/// </summary>
public class ReportFormTableUtil
{
    #region title para
    public static JObject AllTablePara
    {
        get
        {
            string realPath = HttpContext.Current.Server.MapPath(AppConfig.strAppPath + "/DataConfigs/ReportFormTableTitle.json");
            return JObject.Parse(File.ReadAllText(realPath));
        }
    }

    public static JObject getAllTablePara()
    {
        return AllTablePara;
    }

    public static JObject getTableByCode(string code)
    {
      return JObject.Parse(AllTablePara[code].ToString());      
    }
    #endregion

    #region data 
    public static JObject AllTableData
    {
        get
        {
            string data = File.ReadAllText(HttpContext.Current.Server.MapPath(AppConfig.strAppPath + "/Data/testTableData.json"));
            return JObject.Parse(data);
        }
    }

    public static JObject getAllTableData()
    {
        return AllTablePara;
    }

    public static string getDataByCode(string code)
    {
        return AllTableData[code].ToString();
    }
    #endregion
}