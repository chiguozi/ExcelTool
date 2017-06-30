using System.IO;
using NPOI.SS.UserModel;
using UnityEngine;
using UnityEditor;

public class NPOILoader
{
    [MenuItem("Excel/Test")]
    public static void Test()
    {
        string path = Application.dataPath + "/ExcelTool/Excel/Test1.xls";
        ExcelOriginData data = new ExcelOriginData(path);
        var tableData = new TableOriginData();
        tableData.tabName = "test";
        tableData.LastRowNum = 1;
        tableData.AddRow();
        tableData.AddCell(0, new TableOriginCell() { objValue = "123", valueType = ExcelValueType.String, isDirty = true });
        data.tabDataList.Add(tableData);
        var wiret = new NPOIWriter();
        wiret.Write(data, true);
        Debug.LogError("loaded");
    }

    public ExcelOriginData Load(string fullPath)
    {
        var workBook = LoadWorkBook(fullPath);
        return GetExcelOriginData(workBook, fullPath);
    }

    ExcelOriginData GetExcelOriginData(IWorkbook workBook, string fullPath)
    {
        if (workBook == null)
            return null;
        ExcelOriginData excelData = new ExcelOriginData(fullPath);
        for(int i = 0; i < workBook.NumberOfSheets; i++)
        {
            excelData.AddTableData(GetTabOriginData(workBook.GetSheetAt(i)));
        }
        return excelData;
    }


    IWorkbook LoadWorkBook(string fullPath)
    {
        IWorkbook workBook;
        using (FileStream fs = File.Open(fullPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
        {
            if (fs == null)
            {
                Debug.LogError("找不到文件 " + fullPath);
                return null;
            }
            workBook = WorkbookFactory.Create(fs);
        }
        return workBook;
    }

    TableOriginData GetTabOriginData(ISheet sheet)
    {
        TableOriginData tabData = new TableOriginData();
        tabData.tabName = sheet.SheetName;
        tabData.LastRowNum = sheet.LastRowNum;
        //不使用前面的空行保留
        for(int i = 0; i <= sheet.LastRowNum; i++ )
        {
            var row = sheet.GetRow(i);
            tabData.AddRow();
            if (row == null)
                continue;
            for(int j = 0; j < row.LastCellNum; j++)
            {
                tabData.AddCell(i, GetTableOriginCell(row.GetCell(j)));
            }
        }
        return tabData;
    }

    TableOriginCell GetTableOriginCell(ICell cell)
    {
        if (cell == null)
            return null;
        TableOriginCell cellData = new TableOriginCell();
        cellData.objValue = GetValueType(cell);
        //支持公式
        cellData.strValue = cellData.objValue == null? "" : cellData.objValue.ToString();
        cellData.row = cell.RowIndex;
        cellData.column = cell.ColumnIndex;
        var color = cell.CellStyle.FillForegroundColorColor;
        if (color != null)
            cellData.rgb = color.RGB;
        return cellData;
    }

    object GetValueType(ICell cell)
    {
        var type = CellType.Error;
        if (cell == null)
            return null;
        type = cell.CellType;
        switch (type)
        {
            case CellType.Boolean:
                if (cell.BooleanCellValue == true)
                    return "true";
                else
                    return "false";
            case CellType.Numeric:
                return cell.NumericCellValue;
            case CellType.String:
                string str = cell.StringCellValue;
                if (string.IsNullOrEmpty(str))
                    return null;
                return str.ToString();
            case CellType.Error:
                return cell.ErrorCellValue;
            case CellType.Formula:
                {
                    type = cell.CachedFormulaResultType;
                    switch (type)
                    {
                        case CellType.Boolean:
                            if (cell.BooleanCellValue == true)
                                return "true";
                            else
                                return "false";
                        case CellType.Numeric:
                            return cell.NumericCellValue;
                        case CellType.String:
                            string str1 = cell.StringCellValue;
                            if (string.IsNullOrEmpty(str1))
                                return null;
                            return str1.ToString();
                        case CellType.Error:
                            return cell.ErrorCellValue;
                        case CellType.Unknown:
                        case CellType.Blank:
                        default:
                            return null;
                    }
                }
            case CellType.Unknown:
            case CellType.Blank:
            default:
                return null;
        }
    }
}
