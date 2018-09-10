using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Web;

/// <summary>
/// proxysUtil 的摘要说明
/// </summary>
public class HttpProxysUtil
{
    #region http post请求通用类
    /// <summary>
    /// 根据参数执行HTTP请求
    /// </summary>
    /// <param name="url"></param>
    /// <param name="parameters"></param>
    /// <returns></returns>
    public static string HttpPost(string url, IEnumerable<HttpParameter> parameters, string token)
    {
        string paramContent = parameters != null
            ? string.Join("&", parameters.Select(p => p.ToString()))
            : "";

        return HttpPost(url, paramContent, token);
    }

    /// <summary>
    /// POST请求
    /// </summary>
    /// <param name="url">请求url</param>
    /// <param name="postContent">post content</param>
    /// <returns></returns>
    public static string HttpPost(string url, string postContent, string token)
    {
        try
        {
            System.Net.ServicePointManager.DefaultConnectionLimit = 200;
            System.GC.Collect();

            HttpWebRequest request = (HttpWebRequest)HttpWebRequest.Create(url); //HttpWebRequest.Create(url);
            request.KeepAlive = false;
            
            request.ProtocolVersion = HttpVersion.Version11;
            request.UserAgent = "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/49.0.2623.221 Safari/537.36 SE 2.X MetaSr 1.0";
            request.Accept = "*/*";

            byte[] content = Encoding.UTF8.GetBytes(postContent);
            request.ContentLength = content.Length;
            request.ContentType = "application/x-www-form-urlencoded";
            request.Method = WebRequestMethods.Http.Post;
            request.Timeout = 5 * 60 * 1000;
            
            if (token != null && token != "")
            {
                string agstoken = "agstoken=" + token;
                request.Headers.Add("Cookie", agstoken);
            }

            using (Stream requestStream = request.GetRequestStream())
            {
                requestStream.Write(content, 0, content.Length);
                requestStream.Close();
                WebResponse response = request.GetResponse();
                using (Stream responseStream = response.GetResponseStream())
                {
                    using (StreamReader reader = new StreamReader(responseStream))
                    {
                        string strResult = reader.ReadToEnd();
                        reader.Close();
                        return strResult;
                    }
                }
            }
        }
        catch(Exception ex)
        {
            return "";
        }
    }
    #endregion

    #region http下载文件
    /// <summary>
    /// http下载文件
    /// </summary>
    /// <param name="url">下载文件地址</param>
    /// <param name="path">文件存放地址，包含文件名</param>
    /// <returns></returns>
    public static bool HttpDownload(string url, string path)
    {
        string tempPath = System.IO.Path.GetDirectoryName(path) + @"\temp";
        System.IO.Directory.CreateDirectory(tempPath);  //创建临时文件目录
        string tempFile = tempPath + @"\" + System.IO.Path.GetFileName(path) + ".temp"; //临时文件
        if (System.IO.File.Exists(tempFile))
        {
            System.IO.File.Delete(tempFile);    //存在则删除
        }
        try
        {
            FileStream fs = new FileStream(tempFile, FileMode.Append, FileAccess.Write, FileShare.ReadWrite);
            // 设置参数
            HttpWebRequest request = WebRequest.Create(url) as HttpWebRequest;
            //发送请求并获取相应回应数据
            HttpWebResponse response = request.GetResponse() as HttpWebResponse;
            //直到request.GetResponse()程序才开始向目标网页发送Post请求
            Stream responseStream = response.GetResponseStream();

            byte[] bArr = new byte[1024];
            int size = responseStream.Read(bArr, 0, (int)bArr.Length);
            while (size > 0)
            {
                fs.Write(bArr, 0, size);
                size = responseStream.Read(bArr, 0, (int)bArr.Length);
            }

            fs.Close();
            responseStream.Close();
            System.IO.File.Move(tempFile, path);
            return true;
        }
        catch
        {
            return false;
        }
    }
    #endregion
}