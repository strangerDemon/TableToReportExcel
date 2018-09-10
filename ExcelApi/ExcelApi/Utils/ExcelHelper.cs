using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System.Data;
using System.IO;
using System.Text;
using System.Web;
using System;
using System.Drawing;
using System.Collections.Generic;
using NPOI.XSSF.UserModel;
using NPOI.SS.Util;
/// <summary>
///版 本 V1.0
///Copyright (c) 2010-2015 厦门众图地理信息有限公司　　　　　　　　　　　　　　　　　　　　　　　　　　
///创建人：茄子
///日 期：2015/11/25
///描 述：NPOI Excel DataTable操作类
public class ExcelHelper
{
    #region Excel导出方法 ExcelDownload
    /// <summary>
    /// Excel导出下载
    /// </summary>
    /// <param name="dtSource">DataTable数据源</param>
    /// <param name="excelConfig">导出设置包含文件名、标题、列设置</param>
    public static void ExcelDownload(DataTable[] dtSource, ExcelConfig[] excelConfig)
    {
        HttpContext curContext = HttpContext.Current;
        // 设置编码和附件格式
        curContext.Response.ContentType = "application/ms-excel";
        curContext.Response.ContentEncoding = Encoding.UTF8;
        curContext.Response.Charset = "";
        curContext.Response.AppendHeader("Content-Disposition",
            "attachment;filename=" + HttpUtility.UrlEncode(excelConfig[0].FileName, Encoding.UTF8));
        //调用导出具体方法Export()
        var data = ExportMemoryStream(dtSource, excelConfig).GetBuffer();
        curContext.Response.BinaryWrite(data);
        curContext.Response.End();
    }
    /// <summary>
    /// Excel导出下载
    /// </summary>
    /// <param name="list">数据源</param>
    /// <param name="templdateName">模板文件名</param>
    /// <param name="newFileName">文件名</param>
    public static void ExcelDownload(List<TemplateMode> list, string templdateName, string newFileName)
    {
        HttpResponse response = System.Web.HttpContext.Current.Response;
        response.Clear();
        response.Charset = "UTF-8";
        response.ContentType = "application/vnd-excel";//"application/vnd.ms-excel";
        System.Web.HttpContext.Current.Response.AddHeader("Content-Disposition", string.Format("attachment; filename=" + newFileName));
        System.Web.HttpContext.Current.Response.BinaryWrite(ExportListByTempale(list, templdateName).ToArray());
    }
    #endregion

    #region DataTable导出到Excel文件excelConfig中FileName设置为全路径
    /// <summary>
    /// DataTable导出到Excel文件 Export()
    /// </summary>
    /// <param name="dtSource">DataTable数据源</param>
    /// <param name="excelConfig">导出设置包含文件名、标题、列设置</param>
    public static void ExcelExport(DataTable[] dtSource, ExcelConfig[] excelConfig)
    {
        using (MemoryStream ms = ExportMemoryStream(dtSource, excelConfig))
        {
            using (FileStream fs = new FileStream(excelConfig[0].FileName, FileMode.Create, FileAccess.Write))
            {
                byte[] data = ms.ToArray();
                fs.Write(data, 0, data.Length);
                fs.Flush();
            }
        }
    }
    #endregion

    #region DataTable导出到Excel的MemoryStream
    /// <summary>
    /// DataTable导出到Excel的MemoryStream Export()
    /// </summary>
    /// <param name="dtSource">DataTable数据源</param>
    /// <param name="excelConfig">导出设置包含文件名、标题、列设置</param>
    public static MemoryStream ExportMemoryStream(DataTable[] dtSource, ExcelConfig[] excelConfig)
    {
        HSSFWorkbook workbook = new HSSFWorkbook();
        MemoryStream ms = new MemoryStream();
        for (int dtIndex = 0, length = dtSource.Length; dtIndex < length; dtIndex++)
        {
            // int colint = 0;
            DataTable data = dtSource[dtIndex].Copy();
            var tatalCount = dtSource[dtIndex].Columns.Count;
            // var index = 0;
            /* for (int i = 0; i < dtSource[dtIndex].Columns.Count;)
             {
                 index++;
                 DataColumn column = dtSource[dtIndex].Columns[i];
                 if (excelConfig[dtIndex].ColumnEntity[colint].Column != column.ColumnName)
                 {
                     dtSource[dtIndex].Columns.Remove(column.ColumnName);
                 }
                 else
                 {
                     i++;
                     colint++;
                     if (colint == excelConfig[dtIndex].ColumnEntity.Count)
                     {
                         for (var j = index; j < tatalCount; j++)
                         {
                             DataColumn column1 = data.Columns[j];
                             dtSource[dtIndex].Columns.Remove(column1.ColumnName);
                         }
                         break;
                     }
                 }
             }*/
            ISheet sheet = workbook.CreateSheet((dtIndex + 1).ToString());

            #region 右击文件 属性信息
            {
                DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
                dsi.Company = "NPOI";
                workbook.DocumentSummaryInformation = dsi;

                SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
                si.Author = "administrator"; //填加xls文件作者信息
                si.ApplicationName = "ZoonTop"; //填加xls文件创建程序信息
                si.LastAuthor = "administrator"; //填加xls文件最后保存者信息
                si.Comments = "ZoonTop自动生成excel"; //填加xls文件作者信息
                si.Title = ""; //填加xls文件标题信息
                si.Subject = "";//填加文件主题信息
                si.CreateDateTime = System.DateTime.Now;
                workbook.SummaryInformation = si;
            }
            #endregion

            #region 设置标题样式
            ICellStyle headStyle = workbook.CreateCellStyle();
            int[] arrColWidth = new int[dtSource[dtIndex].Columns.Count];
            string[] arrColName = new string[dtSource[dtIndex].Columns.Count];//列名
            ColumnEntity[] columnModel = new ColumnEntity[dtSource[dtIndex].Columns.Count];//航头属性

            ICellStyle[] arryColumStyle = new ICellStyle[dtSource[dtIndex].Columns.Count];//样式表
           

            if (excelConfig[dtIndex].Background != new Color())
            {
                if (excelConfig[dtIndex].Background != new Color())
                {
                    headStyle.FillPattern = FillPattern.SolidForeground;
                    headStyle.FillForegroundColor = GetXLColour(workbook, excelConfig[dtIndex].Background);
                }
            }
            //title文字
            IFont font = workbook.CreateFont();
            font.FontHeightInPoints = excelConfig[dtIndex].TitlePoint;
            if (excelConfig[dtIndex].ForeColor != new Color())
            {
                font.Color = GetXLColour(workbook, excelConfig[dtIndex].ForeColor);
            }
            font.Boldweight = 700;
            font.FontHeightInPoints = 20;
            headStyle.SetFont(font);
            headStyle.ShrinkToFit = true;
            //垂直居中,水平居中
            headStyle.VerticalAlignment = VerticalAlignment.Center;
            headStyle.Alignment = HorizontalAlignment.Center;
            //边框样式
            headStyle.BorderLeft = BorderStyle.Thin;
            headStyle.BorderRight = BorderStyle.Thin;
            headStyle.BorderTop = BorderStyle.Thin;
            headStyle.BorderBottom = BorderStyle.Thin;
            #endregion

            #region 列头及样式
            ICellStyle cHeadStyle = workbook.CreateCellStyle();
            cHeadStyle.Alignment = HorizontalAlignment.Center; // ------------------
            IFont cfont = workbook.CreateFont();
            cfont.FontHeightInPoints = 15;// excelConfig.HeadPoint;
            cHeadStyle.SetFont(cfont);
            //边框样式
            cHeadStyle.BorderLeft = BorderStyle.Thin;
            cHeadStyle.BorderRight = BorderStyle.Thin;
            cHeadStyle.BorderTop = BorderStyle.Thin;
            cHeadStyle.BorderBottom = BorderStyle.Thin;
            //垂直居中
            cHeadStyle.VerticalAlignment = VerticalAlignment.Center;
            #endregion

            #region 设置内容单元格样式
            foreach (DataColumn item in dtSource[dtIndex].Columns)
            {
                ICellStyle columnStyle = workbook.CreateCellStyle();
                columnStyle.Alignment = HorizontalAlignment.Center;
                arrColWidth[item.Ordinal] = Encoding.GetEncoding(936).GetBytes(item.ColumnName.ToString()).Length;
                arrColName[item.Ordinal] = item.ColumnName.ToString();
                if (excelConfig[dtIndex].ColumnEntity != null)
                {
                    ColumnEntity columnentity = excelConfig[dtIndex].ColumnEntity.Find(t => t.Column == item.ColumnName);
                    if (columnentity != null)
                    {
                        arrColName[item.Ordinal] = columnentity.ExcelColumn;
                        columnModel[item.Ordinal] = columnentity;
                        if (columnentity.Width != 0)
                        {
                            arrColWidth[item.Ordinal] = columnentity.Width;
                        }
                        if (columnentity.Background != new Color())
                        {
                            if (columnentity.Background != new Color())
                            {
                                columnStyle.FillPattern = FillPattern.SolidForeground;
                                columnStyle.FillForegroundColor = GetXLColour(workbook, columnentity.Background);
                            }
                        }
                        if (columnentity.Font != null || columnentity.Point != 0 /*|| columnentity.ForeColor != new Color()*/)
                        {
                            IFont columnFont = workbook.CreateFont();
                            columnFont.FontHeightInPoints = 10;
                            if (columnentity.Font != null)
                            {
                                columnFont.FontName = columnentity.Font;
                            }
                            if (columnentity.Point != 0)
                            {
                                columnFont.FontHeightInPoints = columnentity.Point;
                            }
                            if (columnentity.ForeColor != new Color())
                            {
                                columnFont.Color = GetXLColour(workbook, columnentity.ForeColor);
                            }
                            columnStyle.SetFont(font);
                        }
                        columnStyle.Alignment = getAlignment(columnentity.Alignment);
                    }
                }
                //边框样式
                columnStyle.BorderLeft = BorderStyle.Thin;
                columnStyle.BorderRight = BorderStyle.Thin;
                columnStyle.BorderTop = BorderStyle.Thin;
                columnStyle.BorderBottom = BorderStyle.Thin;
                //单元格文字
                IFont columnsFont = workbook.CreateFont();
                columnsFont.FontHeightInPoints = 12;
                columnStyle.SetFont(columnsFont);
                //columnStyle.ShrinkToFit = true;
                columnStyle.WrapText = true;//自动换行
                //垂直居中
                columnStyle.VerticalAlignment = VerticalAlignment.Center;

                arryColumStyle[item.Ordinal] = columnStyle;
            }
            if (excelConfig[dtIndex].IsAllSizeColumn)
            {
                #region 根据列中最长列的长度取得列宽
                for (int i = 0; i < dtSource[dtIndex].Rows.Count; i++)
                {
                    for (int j = 0; j < dtSource[dtIndex].Columns.Count; j++)
                    {
                        if (arrColWidth[j] != 0)
                        {
                            int intTemp = Encoding.GetEncoding(936).GetBytes(dtSource[dtIndex].Rows[i][j].ToString()).Length;
                            if (intTemp > arrColWidth[j])
                            {
                                arrColWidth[j] = intTemp;
                            }
                        }

                    }
                }
                #endregion
            }
            #endregion

            #region 填充数据           
            ICellStyle dateStyle = workbook.CreateCellStyle();
            IDataFormat format = workbook.CreateDataFormat();
            dateStyle.DataFormat = format.GetFormat("yyyy-mm-dd");

            string xh = dtSource[dtIndex].Columns[0].ToString();//第一个参数，用来合并单元格
            int rowIndex = 0;
            int createIndex = 0;//已经创建的行数
            int dataIndex = 1;
            int[] mergeRows = new int[dtSource[dtIndex].Columns.Count];//需要合并的行数

            IRow[] headerRows;
            IRow headerRow;
            bool IsCellRangeAddress;
            int titleRow = 0;
            double[] totalCount = new double[dtSource[dtIndex].Columns.Count];
            if (dtSource[dtIndex].Rows.Count == 0)//当列表头大于65536行时会报错，但是....
            {
                #region 新建表，填充表头，填充列头，样式
                #region 表头及样式
                if (excelConfig[dtIndex].Title != null)
                {
                    headerRow = sheet.CreateRow(0);
                    if (excelConfig[dtIndex].TitleHeight != 0)
                    {
                        headerRow.Height = (short)(excelConfig[dtIndex].TitleHeight * 20);
                    }
                    headerRow.HeightInPoints = 25;
                    headerRow.Height = 45 * 20;
                    headerRow.CreateCell(0).SetCellValue(excelConfig[dtIndex].Title);
                    headerRow.GetCell(0).CellStyle = headStyle;
                    //标题的宽高
                    sheet.AddMergedRegion(new CellRangeAddress(0, 0, 0, dtSource[dtIndex].Columns.Count - 1)); // ------------------                            
                }
                #endregion

                #region 列头及样式
                IsCellRangeAddress = false;
                rowIndex = 1;
                foreach (ColumnEntity columnEntity in excelConfig[dtIndex].ColumnEntity)
                {
                    if (columnEntity.IsCellRangeAddress)
                    {
                        rowIndex = rowIndex > columnEntity.Bottom ? rowIndex : columnEntity.Bottom;
                        IsCellRangeAddress = true;
                    }
                }
                headerRows = new IRow[rowIndex];
                for (int hi = 0; hi < rowIndex; hi++)
                {
                    headerRow = sheet.CreateRow(hi + 1);
                    headerRow.Height = 30 * 20;
                    headerRows[hi] = headerRow;
                }

                #region 如果设置了列标题就按列标题定义列头，没定义直接按字段名输出
                if (!IsCellRangeAddress)//简单表头
                {
                    foreach (DataColumn column in dtSource[dtIndex].Columns)
                    {
                        headerRows[0].CreateCell(column.Ordinal).SetCellValue(arrColName[column.Ordinal]);
                        headerRows[0].GetCell(column.Ordinal).CellStyle = cHeadStyle;
                        sheet.SetColumnWidth(column.Ordinal, arrColWidth[column.Ordinal] * 48);
                    }
                }
                else//复杂表头
                {
                    int[] indexs = new int[rowIndex];
                    int max = 0;
                    foreach (ColumnEntity columnEntity in excelConfig[dtIndex].ColumnEntity)
                    {
                        int headerIndex = columnEntity.Top - 1;
                        int index = indexs[headerIndex];
                        for (int i = 0; i < columnEntity.Left - index; i++)//填充合并项的空缺
                        {
                            headerRows[headerIndex].CreateCell(indexs[headerIndex]).SetCellValue("");
                            headerRows[headerIndex].GetCell(indexs[headerIndex]).CellStyle = cHeadStyle;
                            indexs[headerIndex]++;
                        }
                        headerRows[headerIndex].CreateCell(indexs[headerIndex]).SetCellValue(columnEntity.ExcelColumn);
                        headerRows[headerIndex].GetCell(indexs[headerIndex]).CellStyle = cHeadStyle;
                        sheet.SetColumnWidth(indexs[headerIndex], columnEntity.Width * 32);
                        if (columnEntity.IsCellRangeAddress)
                        {
                            //设置一个合并单元格区域，使用上下左右定义CellRangeAddress区域
                            //CellRangeAddress四个参数为：起始行，结束行，起始列，结束列
                            sheet.AddMergedRegion(new CellRangeAddress(
                                columnEntity.Top,
                                columnEntity.Bottom,
                                columnEntity.Left,
                                columnEntity.Right
                            ));
                        }
                        indexs[headerIndex]++;
                        max = max > indexs[headerIndex] ? max : indexs[headerIndex];
                    }
                    //填充表头的空缺
                    for (int headerDefectCell = 0; headerDefectCell < indexs.Length; headerDefectCell++)
                    {
                        for (int headerDefectCellNum = indexs[headerDefectCell]; headerDefectCellNum < max; headerDefectCellNum++)
                        {
                            headerRows[headerDefectCell].CreateCell(headerDefectCellNum).SetCellValue("");
                            headerRows[headerDefectCell].GetCell(headerDefectCellNum).CellStyle = cHeadStyle;
                        }
                    }
                    #endregion
                }
                rowIndex += 1;
                createIndex = rowIndex;
                #endregion

                #endregion

                #region 合计
                IRow totalRow = sheet.CreateRow(rowIndex);
                totalRow.Height = 30 * 20;
                foreach (DataColumn column in dtSource[dtIndex].Columns)
                {
                    if (column.Ordinal == 0)
                    {
                        totalRow.CreateCell(column.Ordinal).SetCellValue("合计");
                    }
                    else
                    {
                        if (column.DataType.FullName.Equals("System.Double"))
                        {
                            totalRow.CreateCell(column.Ordinal).SetCellValue(0);
                        }
                        else
                        {
                            totalRow.CreateCell(column.Ordinal).SetCellValue("");
                        }
                    }
                    totalRow.GetCell(column.Ordinal).CellStyle = arryColumStyle[column.Ordinal];
                    sheet.SetColumnWidth(column.Ordinal, arrColWidth[column.Ordinal] * 48);
                }
                #endregion
            }
            foreach (DataRow row in dtSource[dtIndex].Rows)
            {
                #region 新建表，填充表头，填充列头，样式 如果没有数据表头为空
                if (rowIndex == 65535 || rowIndex == 0)
                {
                    if (rowIndex != 0)//数据过多创建新的sheet页
                    {
                        sheet = workbook.CreateSheet();
                    }

                    #region 表头及样式
                    if (excelConfig[dtIndex].Title != null)
                    {
                        headerRow = sheet.CreateRow(0);
                        if (excelConfig[dtIndex].TitleHeight != 0)
                        {
                            headerRow.Height = (short)(excelConfig[dtIndex].TitleHeight * 20);
                        }
                        headerRow.HeightInPoints = 25;
                        headerRow.Height = 45 * 20;
                        headerRow.CreateCell(0).SetCellValue(excelConfig[dtIndex].Title);
                        headerRow.GetCell(0).CellStyle = headStyle;
                        //标题的宽高
                        sheet.AddMergedRegion(new CellRangeAddress(0, 0, 0, dtSource[dtIndex].Columns.Count - 1)); // ------------------                            
                    }
                    #endregion

                    #region 列头及样式
                    IsCellRangeAddress = false;
                    rowIndex = 1;
                    foreach (ColumnEntity columnEntity in excelConfig[dtIndex].ColumnEntity)
                    {
                        if (columnEntity.IsCellRangeAddress)
                        {
                            rowIndex = rowIndex > columnEntity.Bottom ? rowIndex : columnEntity.Bottom;
                            IsCellRangeAddress = true;
                        }
                    }
                    headerRows = new IRow[rowIndex];
                    for (int hi = 0; hi < rowIndex; hi++)
                    {
                        headerRow = sheet.CreateRow(hi + 1);
                        headerRow.Height = 30 * 20;
                        headerRows[hi] = headerRow;
                    }

                    #region 如果设置了列标题就按列标题定义列头，没定义直接按字段名输出
                    if (!IsCellRangeAddress)//简单表头
                    {
                        foreach (DataColumn column in dtSource[dtIndex].Columns)
                        {
                            headerRows[0].CreateCell(column.Ordinal).SetCellValue(arrColName[column.Ordinal]);
                            headerRows[0].GetCell(column.Ordinal).CellStyle = cHeadStyle;
                            //设置列宽 // 第二个参数的单位是1/256个字符宽度，但与前端不一致，故改为48
                            //sheet.SetColumnWidth(column.Ordinal, (arrColWidth[column.Ordinal] + 1) * 256);
                            sheet.SetColumnWidth(column.Ordinal, arrColWidth[column.Ordinal] * 48);
                        }
                    }
                    else//复杂表头
                    {
                        int[] indexs = new int[rowIndex];
                        int max = 0;
                        foreach (ColumnEntity columnEntity in excelConfig[dtIndex].ColumnEntity)
                        {
                            int headerIndex = columnEntity.Top - 1;
                            int index = indexs[headerIndex];
                            for (int i = 0; i < columnEntity.Left - index; i++)//填充合并项的空缺
                            {
                                headerRows[headerIndex].CreateCell(indexs[headerIndex]).SetCellValue("");
                                headerRows[headerIndex].GetCell(indexs[headerIndex]).CellStyle = cHeadStyle;
                                indexs[headerIndex]++;
                            }
                            headerRows[headerIndex].CreateCell(indexs[headerIndex]).SetCellValue(columnEntity.ExcelColumn);
                            headerRows[headerIndex].GetCell(indexs[headerIndex]).CellStyle = cHeadStyle;
                            sheet.SetColumnWidth(indexs[headerIndex], columnEntity.Width * 32);
                            if (columnEntity.IsCellRangeAddress)
                            {
                                //设置一个合并单元格区域，使用上下左右定义CellRangeAddress区域
                                //CellRangeAddress四个参数为：起始行，结束行，起始列，结束列
                                sheet.AddMergedRegion(new CellRangeAddress(
                                    columnEntity.Top,
                                    columnEntity.Bottom,
                                    columnEntity.Left,
                                    columnEntity.Right
                                ));
                            }
                            indexs[headerIndex]++;
                            max = max > indexs[headerIndex] ? max : indexs[headerIndex];
                        }
                        //填充表头的空缺
                        for (int headerDefectCell = 0; headerDefectCell < indexs.Length; headerDefectCell++)
                        {
                            for (int headerDefectCellNum = indexs[headerDefectCell]; headerDefectCellNum < max; headerDefectCellNum++)
                            {
                                headerRows[headerDefectCell].CreateCell(headerDefectCellNum).SetCellValue("");
                                headerRows[headerDefectCell].GetCell(headerDefectCellNum).CellStyle = cHeadStyle;
                            }
                        }                     
                        #endregion
                    }
                    rowIndex += 1;
                    createIndex = rowIndex;
                    titleRow = rowIndex;
                    #endregion
                }
                #endregion

                #region 填充内容

                if (createIndex <= rowIndex)
                {
                    IRow dataRow = sheet.CreateRow(rowIndex);
                    dataRow.Height = 30 * 20;
                    createIndex++;
                }
                foreach (DataColumn column in dtSource[dtIndex].Columns)
                {
                    if (columnModel[column.Ordinal].Merge == "row")//合并行，行和列不可能同时出现在一个单元格
                    {
                        ICell newCell = sheet.GetRow(rowIndex).CreateCell(column.Ordinal);
                        newCell.CellStyle = arryColumStyle[column.Ordinal];
                        string drValue = row[column].ToString();
                        if (mergeRows[column.Ordinal] == 0)//计算要合并几行
                        {
                            //不计算父子关系，同级关系，只要和下一个相同，就合并 to edit
                            //改成只有序号一样的合并 to edit 
                            //增加父级序号也要一致(父级序号是子级的子集：父级：B，子级BB)
                            for (int subDataIndex = dataIndex, total = dtSource[dtIndex].Rows.Count; subDataIndex < total; subDataIndex++)
                            {
                                /*if (drValue.Equals(dtSource[dtIndex].Rows[subDataIndex][column].ToString()))
                                {
                                    mergeRows[column.Ordinal]++;
                                }*/
                                string[] arr = column.ColumnName.ToString().Split('_');
                                string mark = "";

                                if (arr.Length > 1)
                                {
                                    int markCount = 0;
                                    int markLength = arr[0].Length;
                                    bool isAllEqually = true;
                                    foreach (DataColumn parentColumn in dtSource[dtIndex].Columns)
                                    {
                                        if (markCount <= markLength)
                                        {
                                            mark = markCount > 0 ? arr[0].Substring(0, markCount) + "_XH" : "XH";
                                            if (parentColumn.ColumnName.Equals(mark))
                                            {
                                                if (!row[parentColumn.ColumnName].ToString().Equals(dtSource[dtIndex].Rows[subDataIndex][parentColumn.ColumnName].ToString()))
                                                {
                                                    isAllEqually = false;
                                                }
                                                markCount++;
                                            }
                                        }
                                        else
                                        {
                                            break;
                                        }
                                    }
                                    if (isAllEqually) mergeRows[column.Ordinal]++;
                                }
                                else
                                {
                                    mark = "XH";
                                    //当前序号一致 最外层
                                    if (row[mark].ToString().Equals(dtSource[dtIndex].Rows[subDataIndex][mark].ToString()))
                                    {
                                        mergeRows[column.Ordinal]++;
                                    }
                                }
                            }
                            for (int i = 1; i <= mergeRows[column.Ordinal]; i++)//预创建行
                            {
                                if (sheet.GetRow(rowIndex + i) != null)
                                {
                                    IRow dataRow = sheet.CreateRow(rowIndex + i);
                                    dataRow.Height = 30 * 20;
                                    createIndex++;
                                }
                            }
                            sheet.AddMergedRegion(new CellRangeAddress(
                                        rowIndex,
                                        rowIndex + mergeRows[column.Ordinal],
                                        column.Ordinal,
                                        column.Ordinal
                                    ));
                        }
                        else
                        {
                            drValue = "";
                            mergeRows[column.Ordinal]--;
                        }
                        SetCell(newCell, dateStyle, column.DataType, drValue);
                    }
                    else if (columnModel[column.Ordinal].Merge == "col")//合并列
                    {

                    }
                    else//无合并单元格
                    {
                        ICell newCell = sheet.GetRow(rowIndex).CreateCell(column.Ordinal);
                        newCell.CellStyle = arryColumStyle[column.Ordinal];
                        string drValue = row[column].ToString();
                        SetCell(newCell, dateStyle, column.DataType, drValue);
                    }
                    if (column.DataType.FullName.Equals("System.Double"))
                    {
                        double value;
                        try
                        {
                            value = double.Parse(row[column].ToString());
                            totalCount[column.Ordinal] += value;
                        }
                        catch
                        {
                            totalCount[column.Ordinal] += 0;
                        }
                    }
                }
                #endregion

                #region 合计
                if (rowIndex == 65534 || dtSource[dtIndex].Rows.Count + titleRow == rowIndex + 1)//最后一行 或者所有数据填充完
                {
                    IRow totalRow = sheet.CreateRow(rowIndex + 1);
                    totalRow.Height = 30 * 20;
                    foreach (DataColumn column in dtSource[dtIndex].Columns)
                    {
                        if (column.Ordinal == 0)
                        {
                            totalRow.CreateCell(column.Ordinal).SetCellValue("合计");
                        }
                        else
                        {
                            if (column.DataType.FullName.Equals("System.Double"))
                            {
                                if (column.ColumnName.IndexOf("_") > 0)//子表 包含小计
                                {
                                    totalRow.CreateCell(column.Ordinal).SetCellValue(totalCount[column.Ordinal] / 2);
                                }
                                else
                                {
                                    totalRow.CreateCell(column.Ordinal).SetCellValue(totalCount[column.Ordinal]);
                                }
                            }
                            else
                            {
                                totalRow.CreateCell(column.Ordinal).SetCellValue("");
                            }
                        }
                        totalRow.GetCell(column.Ordinal).CellStyle = arryColumStyle[column.Ordinal];
                        sheet.SetColumnWidth(column.Ordinal, arrColWidth[column.Ordinal] * 48);
                    }
                    rowIndex++;
                }
                #endregion

                rowIndex++;
                dataIndex++;
            }
            #endregion
        }
        workbook.Write(ms);
        ms.Flush();
        ms.Position = 0;
        return ms;
    }
    #endregion

    #region ListExcel导出(加载模板)
    /// <summary>
    /// List根据模板导出ExcelMemoryStream
    /// </summary>
    /// <param name="list"></param>
    /// <param name="templdateName"></param>
    public static MemoryStream ExportListByTempale(List<TemplateMode> list, string templdateName)
    {
        try
        {

            string templatePath = HttpContext.Current.Server.MapPath("/") + "/Resource/ExcelTemplate/";
            string templdateName1 = string.Format("{0}{1}", templatePath, templdateName);

            FileStream fileStream = new FileStream(templdateName1, FileMode.Open, FileAccess.Read);
            ISheet sheet = null;
            if (templdateName.IndexOf(".xlsx") == -1)//2003
            {
                HSSFWorkbook hssfworkbook = new HSSFWorkbook(fileStream);
                sheet = hssfworkbook.GetSheetAt(0);
                SetPurchaseOrder(sheet, list);
                sheet.ForceFormulaRecalculation = true;
                using (MemoryStream ms = new MemoryStream())
                {
                    hssfworkbook.Write(ms);
                    ms.Flush();
                    return ms;
                }
            }
            else//2007
            {
                XSSFWorkbook xssfworkbook = new XSSFWorkbook(fileStream);
                sheet = xssfworkbook.GetSheetAt(0);
                SetPurchaseOrder(sheet, list);
                sheet.ForceFormulaRecalculation = true;
                using (MemoryStream ms = new MemoryStream())
                {
                    xssfworkbook.Write(ms);
                    ms.Flush();
                    return ms;
                }
            }

        }
        catch (Exception)
        {
            throw;
        }
    }
    /// <summary>
    /// 赋值单元格
    /// </summary>
    /// <param name="sheet"></param>
    /// <param name="list"></param>
    private static void SetPurchaseOrder(ISheet sheet, List<TemplateMode> list)
    {
        try
        {
            foreach (var item in list)
            {
                IRow row = null;
                ICell cell = null;
                row = sheet.GetRow(item.row);
                if (row == null)
                {
                    row = sheet.CreateRow(item.row);
                }
                cell = row.GetCell(item.cell);
                if (cell == null)
                {
                    cell = row.CreateCell(item.cell);
                }
                cell.SetCellValue(item.value);
            }
        }
        catch (Exception)
        {
            throw;
        }
    }
    #endregion

    #region 设置表格内容
    private static void SetCell(ICell newCell, ICellStyle dateStyle, Type dataType, string drValue)
    {
        switch (dataType.ToString())
        {
            case "System.String"://字符串类型
                newCell.SetCellValue(drValue);
                break;
            case "System.DateTime"://日期类型
                System.DateTime dateV;
                if (System.DateTime.TryParse(drValue, out dateV))
                {
                    newCell.SetCellValue(dateV);
                }
                else
                {
                    newCell.SetCellValue("");
                }
                newCell.CellStyle = dateStyle;//格式化显示
                break;
            case "System.Boolean"://布尔型
                bool boolV = false;
                bool.TryParse(drValue, out boolV);
                newCell.SetCellValue(boolV);
                break;
            case "System.Int16"://整型
            case "System.Int32":
            case "System.Int64":
            case "System.Byte":
                int intV = 0;
                int.TryParse(drValue, out intV);
                newCell.SetCellValue(intV);
                break;
            case "System.Decimal"://浮点型
            case "System.Double":
                double doubV = 0;
                double.TryParse(drValue, out doubV);
                newCell.SetCellValue(doubV);
                break;
            case "System.DBNull"://空值处理
                newCell.SetCellValue("");
                break;
            default:
                newCell.SetCellValue("");
                break;
        }
    }
    #endregion

    #region 从Excel导入
    /// <summary>
    /// 读取excel ,默认第一行为标头
    /// </summary>
    /// <param name="strFileName">excel文档路径</param>
    /// <returns></returns>
    public static DataTable ExcelImport(string strFileName)
    {
        DataTable dt = new DataTable();

        ISheet sheet = null;
        using (FileStream file = new FileStream(strFileName, FileMode.Open, FileAccess.Read))
        {
            if (strFileName.IndexOf(".xlsx") == -1)//2003
            {
                HSSFWorkbook hssfworkbook = new HSSFWorkbook(file);
                sheet = hssfworkbook.GetSheetAt(0);
            }
            else//2007
            {
                XSSFWorkbook xssfworkbook = new XSSFWorkbook(file);
                sheet = xssfworkbook.GetSheetAt(0);
            }
        }

        System.Collections.IEnumerator rows = sheet.GetRowEnumerator();

        IRow headerRow = sheet.GetRow(0);
        int cellCount = headerRow.LastCellNum;

        for (int j = 0; j < cellCount; j++)
        {
            ICell cell = headerRow.GetCell(j);
            dt.Columns.Add(cell.ToString());
        }

        for (int i = (sheet.FirstRowNum + 1); i <= sheet.LastRowNum; i++)
        {
            IRow row = sheet.GetRow(i);
            DataRow dataRow = dt.NewRow();

            for (int j = row.FirstCellNum; j < cellCount; j++)
            {
                if (row.GetCell(j) != null)
                    dataRow[j] = row.GetCell(j).ToString();

            }
            dt.Rows.Add(dataRow);


        }
        return dt;
    }
    #endregion

    #region RGB颜色转NPOI颜色
     private static short GetXLColour(HSSFWorkbook workbook, Color SystemColour)
     {
         short s = 0;
         HSSFPalette XlPalette = workbook.GetCustomPalette();
         NPOI.HSSF.Util.HSSFColor XlColour = XlPalette.FindColor(SystemColour.R, SystemColour.G, SystemColour.B);
         if (XlColour == null)
         {
             if (NPOI.HSSF.Record.PaletteRecord.STANDARD_PALETTE_SIZE < 255)
             {
                 XlColour = XlPalette.FindSimilarColor(SystemColour.R, SystemColour.G, SystemColour.B);
                 s = XlColour.Indexed;
             }

         }
         else
             s = XlColour.Indexed;
         return s;
     }
    #endregion

    #region 设置列的对齐方式
    /// <summary>
    /// 设置对齐方式
    /// </summary>
    /// <param name="style"></param>
    /// <returns></returns>
    private static HorizontalAlignment getAlignment(string style)
    {
        switch (style)
        {
            case "center":
                return HorizontalAlignment.Center;
            case "left":
                return HorizontalAlignment.Left;
            case "right":
                return HorizontalAlignment.Right;
            case "fill":
                return HorizontalAlignment.Fill;
            case "justify":
                return HorizontalAlignment.Justify;
            case "centerselection":
                return HorizontalAlignment.CenterSelection;
            case "distributed":
                return HorizontalAlignment.Distributed;
        }
        return NPOI.SS.UserModel.HorizontalAlignment.General;


    }

    #endregion
}

