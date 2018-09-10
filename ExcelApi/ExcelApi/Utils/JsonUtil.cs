using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Script.Services;
using System.Web.Script.Serialization;
using System.Data;

/// <summary>
/// JsonUtil 的摘要说明
/// </summary>
public class JsonUtil
{
	public JsonUtil()
	{
		//
		// TODO: 在此处添加构造函数逻辑
		//
	}

    /// <summary>
    /// 合并两个数组
    /// </summary>
    /// <param name="ja1"></param>
    /// <param name="ja2"></param>
    /// <returns></returns>
    public static JArray mergeJArray(JArray ja1, JArray ja2)
    {
        try
        {
            JArray jaResult = new JArray();

            if (ja1 != null)
            {
                foreach (var item in ja1)
                {
                    jaResult.Add(item);
                }
            }

            if (ja2 != null)
            {
                foreach (var item in ja2)
                {
                    jaResult.Add(item);
                }
            }

            return jaResult;
        }
        catch (Exception)
        {
            return null;
        }
    }

    /// <summary>
    /// 将字符串转换为JSON对象
    /// </summary>
    /// <param name="jsonText"></param>
    public static JObject strToJson(string jsonText)
    {
        try
        {
            JObject jo = JObject.Parse(jsonText);
            return jo;
        }
        catch
        {
            return new JObject();
        }
    }

    /// <summary>
    /// 从字符串中获取token字符串
    /// </summary>
    /// <param name="jsonText"></param>
    public static string getTokenFromStr(string jsonText)
    {
        JObject jo = JObject.Parse(jsonText);
        return jo["Token"].ToString();
    }

    /// <summary>
    /// 从JSON对象中获取token
    /// </summary>
    /// <param name="jo"></param>
    /// <returns></returns>
    public static string getTokenFromJson(JObject jo)
    {
        return jo["Token"].ToString();
    }

	/// <summary>
	///	判断输入的数据格式是否为Json	
	/// </summary>
	/// <param name="json"></param>
	/// <returns></returns>
	public static bool isJson(string json)
	{
		try
		{
            JObject jo = JObject.Parse(json);
			return true;
		}
		catch
		{
			return false;
		}
	}

    /// <summary>
    ///	将对象转换为json
    /// </summary>
    /// <param name="json"></param>
    /// <returns></returns>
    public static string obj2Json(object obj)
    {
        try
        {
            string strResult = JsonConvert.SerializeObject(obj);
            return strResult;
        }
        catch
        {
            return null;
        }
    }

    public static object ToJson(string Json)
    {
        return Json == null ? null : JsonConvert.DeserializeObject(Json);
    }
    public static string ToJson(object obj)
    {
        var timeConverter = new IsoDateTimeConverter { DateTimeFormat = "yyyy-MM-dd HH:mm:ss" };
        return JsonConvert.SerializeObject(obj, timeConverter);
    }
    public static string ToJson(object obj, string datetimeformats)
    {
        var timeConverter = new IsoDateTimeConverter { DateTimeFormat = datetimeformats };
        return JsonConvert.SerializeObject(obj, timeConverter);
    }
    public static T ToObject<T>(string Json)
    {
        return Json == null ? default(T) : JsonConvert.DeserializeObject<T>(Json);
    }
    public static List<T> ToList<T>(string Json)
    {
        return Json == null ? null : JsonConvert.DeserializeObject<List<T>>(Json);
    }
    public static DataTable ToTable(string Json)
    {
        return Json == null ? null : JsonConvert.DeserializeObject<DataTable>(Json);
    }
    public static JObject ToJObject(string Json)
    {
        return Json == null ? JObject.Parse("{}") : JObject.Parse(Json.Replace("&nbsp;", ""));
    }
}