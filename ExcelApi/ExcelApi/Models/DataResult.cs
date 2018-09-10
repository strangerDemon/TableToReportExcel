using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Xml.Serialization;

/// <summary>
/// DataResult 的摘要说明
/// </summary>
public class DataResult
{
    /// <summary>
    /// 执行结果代码 
    /// 0 失败
    /// 1 成功
    /// </summary>
    public int RespCode
    {
        get;
        set;
    }

    /// <summary>
    /// 执行结果说明
    /// </summary>
    public string RespDesc
    {
        get;
        set;
    }

    /// <summary>
    /// 执行结果
    /// </summary>
    public object Results
    {
        get;
        set;
    }
}