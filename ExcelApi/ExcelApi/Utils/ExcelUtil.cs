using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.IO;
using Aspose.Words.Tables;
using System.Drawing;


/// <summary>
/// 版 本
/// Copyright (c) 2014-2016 厦门众图地理信息有限公司
/// 创建人：冬瓜
/// 日 期：2016.2.03 10:58
/// 描 述：公共控制器
/// </summary>
public class ExcelUtil
{
    #region 导出Excel

    /// <summary>
    /// 执行导出Excel
    /// </summary>
    /// <param name="columnJson">表头</param>
    /// <param name="rowData">数据</param>
    /// <param name="filename">文件名</param>
    /// <returns></returns>
    public static void ExecuteExportExcel(string[] columnJsons, DataTable[] rowData, string filename)
    {
        //设置导出格式
        ExcelConfig[] excelconfigs = new ExcelConfig[columnJsons.Length];
        for (int columnIndex = 0, length = columnJsons.Length; columnIndex < length; columnIndex++)
        {
            ExcelConfig excelconfig = new ExcelConfig();
            excelconfig.Title = filename;
            excelconfig.TitleFont = "微软雅黑";
            excelconfig.TitlePoint = 15;
            excelconfig.Background = Color.FromArgb(235, 238, 245);
            excelconfig.FileName = filename + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
            //excelconfig.IsAllSizeColumn = true;
            excelconfig.ColumnEntity = new List<ColumnEntity>();
            //表头
            List<GridColumnModel> columnData = JsonUtil.ToList<GridColumnModel>(columnJsons[columnIndex]);

            foreach (GridColumnModel gridcolumnmodel in columnData)
            {
                if (gridcolumnmodel.hidden.ToLower() == "false" && gridcolumnmodel.label != null)
                {
                    string align = gridcolumnmodel.align;
                    excelconfig.ColumnEntity.Add(new ColumnEntity()
                    {
                        Column = gridcolumnmodel.name,
                        ExcelColumn = gridcolumnmodel.label,
                        Width = gridcolumnmodel.width,
                        Alignment = gridcolumnmodel.align,
                        IsCellRangeAddress = gridcolumnmodel.isCellRangeAddress,
                        Left = gridcolumnmodel.left,
                        Right = gridcolumnmodel.right,
                        Top = gridcolumnmodel.top,
                        Bottom = gridcolumnmodel.bottom,
                        Merge=gridcolumnmodel.merge
                    });
                }
            }
            excelconfigs[columnIndex] = excelconfig;
        }

        ExcelHelper.ExcelDownload(rowData, excelconfigs);
    }
    #endregion

    public static void MergeCells(Cell startCell, Cell endCell)
    {
        Table parentTable = startCell.ParentRow.ParentTable;

        // Find the row and cell indices for the start and end cell.
        Point startCellPos = new Point(startCell.ParentRow.IndexOf(startCell), parentTable.IndexOf(startCell.ParentRow));
        Point endCellPos = new Point(endCell.ParentRow.IndexOf(endCell), parentTable.IndexOf(endCell.ParentRow));
        // Create the range of cells to be merged based off these indices. Inverse each index if the end cell if before the start cell. 
        Rectangle mergeRange = new Rectangle(Math.Min(startCellPos.X, endCellPos.X), Math.Min(startCellPos.Y, endCellPos.Y),
            Math.Abs(endCellPos.X - startCellPos.X) + 1, Math.Abs(endCellPos.Y - startCellPos.Y) + 1);

        foreach (Row row in parentTable.Rows)
        {
            foreach (Cell cell in row.Cells)
            {
                Point currentPos = new Point(row.IndexOf(cell), parentTable.IndexOf(row));

                // Check if the current cell is inside our merge range then merge it.
                if (mergeRange.Contains(currentPos))
                {
                    if (currentPos.X == mergeRange.X)
                        cell.CellFormat.HorizontalMerge = CellMerge.First;
                    else
                        cell.CellFormat.HorizontalMerge = CellMerge.Previous;

                    if (currentPos.Y == mergeRange.Y)
                        cell.CellFormat.VerticalMerge = CellMerge.First;
                    else
                        cell.CellFormat.VerticalMerge = CellMerge.Previous;
                }
            }
        }
    }
}
