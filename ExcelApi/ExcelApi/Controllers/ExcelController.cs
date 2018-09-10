using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using System.Data;
using System.Web.Script.Serialization;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ExcelApi.Controllers
{
    public class ExcelController : ApiController
    {
        JObject joPara = null;
        public DataResult oDataResult = new DataResult();

        #region 导出表单 空间分析/分析报表
        /// <summary>
        /// 
        /// </summary>
        /// <param name="tableCode">table code</param>
        /// <param name="param">参数,请求参数、数据源、处理chartData数据</param>
        [HttpGet]
        public void exportTableToExcel(string tableCode, string param)
        {
            try
            {
                JObject table = ReportFormTableUtil.getTableByCode(tableCode);//表单
                JArray tableTitle = JArray.Parse(table["titleList"].ToString());//table 表头
                string method = table["method"].ToString();//请求数据的函数
                int tableDeep = table.Property("tableDeep") == null ? 1 : int.Parse(table["tableDeep"].ToString());//复杂表单模板,复杂表头、复杂表单、复杂表头+表单 表头深度
                bool isSheet = table.Property("isSheet") == null ? false : bool.Parse(table["isSheet"].ToString());//是否是多个sheet
                string type = table["type"].ToString();//请求类型
                string fileName = table["label"].ToString();//不需要带后缀名
                
                string dataString = methodForData(method, type, param, ref tableTitle, ref fileName);
                JObject json = new JObject();
                string fileUrl = AppConfig.dataPath + @"\exportTableToExcel\";

                List<DataTable> dataList = new List<DataTable>();
                List<string> columnJsonList = new List<string>();
                List<string> codeList = new List<string>();
                if (!isSheet)//单个sheet
                {
                    JArray dataJa = JArray.Parse(dataString);
                    string columnJson = "[";
                    string codes = "";

                    DataTable data = new DataTable();
                    int left = 0, width = 0;//单元格开始位置和宽度
                    string columnCode = "";//单元格code和类型
                    for (int i = 0, count = tableTitle.Count; i < count; i++)
                    {
                        JObject column = JObject.Parse(tableTitle[i].ToString());
                        columnJson += getFotmatColumnJson(column, left, 1, tableDeep, ref width, ref columnCode, ref data);
                        codes += columnCode;
                        left += width;
                        width = 0;
                        columnCode = "";
                    }
                    columnJson = columnJson.Substring(0, columnJson.Length - 1);
                    codes = codes.Substring(0, codes.Length - 1);
                    columnJson += "]";
                    DataRow row;
                    string[] codeArr = codes.Split(',');
                    foreach (JObject jo in dataJa)
                    {
                        row = data.NewRow();
                        foreach (string code in codeArr)//遍历元素属性
                        {
                            try
                            {
                                row[code] = jo[code];
                            }
                            catch//赋值失败一定是数字的类型赋值空或字符 数字类型
                            {
                                row[code] = 0;
                            }
                        }
                        data.Rows.Add(row);
                    }
                    dataList.Add(data);
                    columnJsonList.Add(columnJson);
                }
                else//多个sheet
                {
                    JObject dataJo = JObject.Parse(dataString);
                    for (int i = 0, count = tableTitle.Count; i < count; i++)
                    {
                        string columnJson = "[";
                        string codes = "";
                        JObject column = JObject.Parse(tableTitle[i].ToString());
                        JArray list = JArray.Parse(column["list"].ToString());
                        DataTable data = new DataTable();
                        int left = 0, width = 1;//单元格开始位置和宽度
                        string columnCode = "";
                        for (int subIndex = 0, listCount = list.Count; subIndex < listCount; subIndex++)
                        {
                            JObject subColumn = JObject.Parse(list[subIndex].ToString());
                            columnJson += getFotmatColumnJson(subColumn, left, 1, 1, ref width, ref columnCode, ref data);
                            codes += columnCode;
                            left += width;
                            width = 0;
                            columnCode = "";
                        }
                        columnJson = columnJson.Substring(0, columnJson.Length - 1);
                        codes = codes.Substring(0, codes.Length - 1);
                        columnJson += "]";
                        DataRow row;
                        string[] codeArr = codes.Split(',');
                        JArray dataJa = JArray.Parse(dataJo[column["code"].ToString()].ToString());
                        foreach (JObject jo in dataJa)//多层结构数据
                        {
                            row = data.NewRow();
                            foreach (string code in codeArr)//遍历元素属性
                            {
                                try
                                {
                                    row[code] = jo[code];
                                }
                                catch//赋值失败一定是数字的类型赋值空或字符 数字类型
                                {
                                    row[code] = 0;
                                }
                            }
                            data.Rows.Add(row);
                        }
                        dataList.Add(data);
                        columnJsonList.Add(columnJson);
                    }
                }
                ExcelUtil.ExecuteExportExcel(columnJsonList.ToArray(), dataList.ToArray(), fileName);
            }
            catch (Exception ex)
            {
                logUtil.RecordErrorToFile(ex, joPara, "exportTableToExcel");
                ReturnData(0, ex.Message, "");
            }
            finally
            {
            }
        }

        /// <summary>
        ///  根据method 和param 获取对应的数据
        /// </summary>
        /// <param name="method"></param>
        /// <param name="type"></param>
        /// <param name="param"></param>
        /// <param name="tableTitle"></param>
        /// <param name="fileName"></param>
        /// <returns></returns>
        private string methodForData(string method, string type, string param, ref JArray tableTitle, ref string fileName)
        {
            JObject data = new JObject();
            string dataArrStr = "[]";
            switch (type)
            {
                case "reportApi"://通过 repost api 获取
                    JObject para = new JObject();
                    JObject pageInfo = new JObject();
                    pageInfo["rows"] = -1;
                    pageInfo["page"] = 0;
                    para["queryJson"] = param;
                    para["pagination"] = pageInfo;
                    data = BizApiUtil.ApiRequest(method, para.ToString());
                    JObject results = JObject.Parse(data["Results"].ToString());
                    if (results.Property("data") != null)
                    {
                        dataArrStr = results["data"].ToString();
                    }
                    else if (results.Property("value") != null)
                    {
                        dataArrStr = results["value"].ToString();
                    }
                    else
                    {
                        dataArrStr = data["Results"].ToString();
                    }
                    break;
                case "json"://json 本地获取
                    dataArrStr = ReportFormTableUtil.getDataByCode(method);
                    break;
                case "httpPost"://通过 http post 获取数据
                    break;
                case "exe"://通过exe 获取数据
                    break;
                case "self"://数据为参数
                    dataArrStr = param;
                    break;
                case "dataFromParam"://表头，文件名来自参数或者函数，不能共用部分
                    break;
                default:
                    break;

            }
            return dataArrStr;
        }

        /// <summary>
        /// 深度迭代
        /// 根据JObject 获取表单格式 递归实现，因为column内坑了含有多层级结构,
        /// 先子后父，才能获取父的宽度
        /// </summary>
        /// <param name="column"></param>
        /// <param name="left"></param>
        /// <param name="deep"></param>
        /// <param name="tableDeep"></param>
        /// <param name="width"></param>
        /// <param name="code">返回code 和它是否是数字类型</param>
        /// <param name="data"></param>
        /// <returns></returns>
        private string getFotmatColumnJson(JObject column, int left, int deep, int tableDeep, ref int width, ref string code, ref DataTable data)
        {
            try
            {
                if (deep > tableDeep || column == null)
                {
                    return "";
                }
                string align = column.Property("align") == null ? "center" : column["align"].ToString();
                string merge = column.Property("merge") == null ? "" : column["merge"].ToString();//合并方式，这个是数据的单元格合并方式row,col
                string columnJson = "";
                if (column.Property("children") != null)
                {
                    JArray children = JArray.Parse(column["children"].ToString());
                    int totalWidth = 0;
                    int parentLeft = left;
                    foreach (JObject child in children)
                    {
                        columnJson += getFotmatColumnJson(child, left, deep + 1, tableDeep, ref width, ref code, ref data);
                        left += width;
                        totalWidth += width;
                    }
                    width = totalWidth;
                    string position = "";
                    if (tableDeep > 1)
                    {
                        position = "\"isCellRangeAddress\":true," +
                                   "\"left\":" + parentLeft + "," +
                                   "\"right\":" + (parentLeft + totalWidth - 1) + "," +
                                   "\"top\":" + deep + "," +
                                   "\"bottom\":" + deep + ",";
                    }
                    else
                    {
                        position = "\"isCellRangeAddress\":false,";
                    }
                    columnJson = "{" +
                                   "\"label\":\"" + column["label"].ToString() + "\"," +
                                   "\"name\":\"" + column["code"].ToString() + "\"," +
                                   "\"index\":\"" + column["code"].ToString() + "\"," +
                                   "\"width\":" + column["width"].ToString() + "," +
                                   "\"align\":\"" + align + "\"," +
                                   "\"merge\":\"" + merge + "\"," +
                                   "\"sortable\":true," +
                                   "\"title\":true," +
                                   "\"lso\":\"\"," +
                                   "\"hidden\":false," +
                                   "\"widthOrg\":150," +
                                   position +
                                   "\"resizable\":true}," + columnJson;
                }
                else
                {
                    width = 1;
                    string position = "";
                    if (tableDeep > 1)
                    {
                        position = "\"isCellRangeAddress\":true," +
                                   "\"left\":" + left + "," +
                                   "\"right\":" + (left + width - 1) + "," +
                                   "\"top\":" + deep + "," +
                                   "\"bottom\":" + tableDeep + ",";
                    }
                    else
                    {
                        position = "\"isCellRangeAddress\":false,";
                    }
                    columnJson += "{" +
                                   "\"label\":\"" + column["label"].ToString() + "\"," +
                                   "\"name\":\"" + column["code"].ToString() + "\"," +
                                   "\"index\":\"" + column["code"].ToString() + "\"," +
                                   "\"width\":" + column["width"].ToString() + "," +
                                   "\"align\":\"" + align + "\"," +
                                   "\"merge\":\"" + merge + "\"," +
                                   "\"sortable\":true," +
                                   "\"title\":true," +
                                   "\"lso\":\"\"," +
                                   "\"hidden\":false," +
                                   "\"widthOrg\":150," +
                                   position +
                                   "\"resizable\":true},";
                    if (column.Property("isNumber") != null && bool.Parse(column["isNumber"].ToString()))//是否为数字类型
                    {
                        data.Columns.Add(column["code"].ToString(), Type.GetType("System.Double"));//名称
                    }
                    else
                    {
                        data.Columns.Add(column["code"].ToString(), Type.GetType("System.String"));//名称
                    }
                    code += column["code"].ToString() + ",";
                }
                return columnJson;
            }
            catch (Exception ex)
            {
                logUtil.RecordExceptionToFile(ex);
                return "";
            }
        }
        #endregion

        #region 公共函数
        public List<T> JSONStringToList<T>(string strJson)
        {
            JavaScriptSerializer serializer = new JavaScriptSerializer();
            List<T> objList = serializer.Deserialize<List<T>>(strJson);
            return objList;
        }
        
        /// <summary>
        /// 返回数据
        /// </summary>
        /// <param name="code">代码</param>
        /// <param name="desc">描述</param>
        /// <param name="data">数据 "",null,JObject,JArray</param>
        private HttpResponseMessage ReturnData(int code, string desc, object data)
        {
            DataResult objDataResult = new DataResult();
            objDataResult.RespCode = code;
            objDataResult.RespDesc = desc;
            objDataResult.Results = data;
            string json = JsonConvert.SerializeObject(objDataResult);
            return new HttpResponseMessage { Content = new StringContent(json, System.Text.Encoding.UTF8, "application/json") };
        }
        #endregion

        #region mvc api 自带
        // GET api/<controller>
        public IEnumerable<string> Get()
        {
            return new string[] { "value1", "value2" };
        }

        // GET api/<controller>/5
        public string Get(int id)
        {
            return "value";
        }

        // POST api/<controller>
        public void Post([FromBody]string value)
        {
            joPara = JObject.Parse(value);
        }

        // PUT api/<controller>/5
        public void Put(int id, [FromBody]string value)
        {
            joPara = JObject.Parse(value);
        }

        // DELETE api/<controller>/5
        public void Delete(int id)
        {
        }
        #endregion
    }
}