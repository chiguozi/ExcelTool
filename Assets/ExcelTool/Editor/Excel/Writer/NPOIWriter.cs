using System.Collections.Generic;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

public class NPOIWriter
{

    //路径合法性外部判断
    public void Write(ExcelOriginData excelData, bool autoCreate)
    {
        var workBook = LoadWorkBook(excelData.fullPath);
        if (workBook == null && autoCreate)
            workBook = CreateWorkBook(excelData.fullPath);
        if (workBook == null)
            return;
        for(int i = 0; i < excelData.tabDataList.Count; i++)
        {
            WriteTabDataToSheet(workBook, excelData.tabDataList[i]);
        }
        SaveToFile(workBook, excelData.fullPath);
    }

    void SaveToFile(IWorkbook workBook, string fullPath)
    {
        using (FileStream fs = File.Open(fullPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.ReadWrite))
        {
            workBook.Write(fs);
            workBook.Close();
            fs.Close();
        }
    }


    void WriteTabDataToSheet(IWorkbook workBook, TableOriginData tabData)
    {
        var sheet = GetWorkSheet(workBook, tabData.tabName);
        for(int i = 0; i < tabData.LastRowNum; i++)
        {
            var dataList = tabData.GetRow(i);
            WriteDataToExcelRow(sheet, dataList, i);
        }
    }

    void WriteDataToExcelRow(ISheet sheet, List<TableOriginCell> dataList, int rowIndex)
    {
        if (dataList == null || dataList.Count  == 0)
            return;
        var row = sheet.GetRow(rowIndex);
        if (row == null)
            row = sheet.CreateRow(rowIndex);
        for(int i = 0; i < dataList.Count; i++)
        {
            ICell cell = row.GetCell(i);
            if (cell == null)
                cell = row.CreateCell(i);
            SetValueToCell(cell, dataList[i]);
        }
    }

    void SetValueToCell(ICell cell, TableOriginCell cellData)
    {
        if (cellData == null || !cellData.isDirty || cellData.objValue == null)
            return;
        var type = cellData.valueType;
        var data = cellData.objValue;
        switch (type)
        {
            case ExcelValueType.Boolean:
                cell.SetCellValue((bool)data);
                break;
            case ExcelValueType.Numeric:
                cell.SetCellValue((double)data);
                break;
            case ExcelValueType.String:
                cell.SetCellValue((string)data);
                break;
            case ExcelValueType.Formula:
                cell.SetCellFormula(cellData.originString);
                break;
            default:
                break;
        }
    }

    ISheet GetWorkSheet(IWorkbook workBook, string name)
    {
        var sheet = workBook.GetSheet(name);
        if(sheet == null)
        {
            sheet = workBook.CreateSheet(name);
        }
        return sheet;
    }

    IWorkbook LoadWorkBook(string fullPath)
    {
        IWorkbook workBook;
        using (FileStream fs = File.Open(fullPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
        {
            workBook = WorkbookFactory.Create(fs);
        }
        return workBook;
    }

    IWorkbook CreateWorkBook(string fullPath)
    {
        var excelType = ExcelUtil.GetExcelType(fullPath);
        switch (excelType)
        {
            case ExcelType.XLS:
                return new HSSFWorkbook();
            //暂时不支持xlsx
            case ExcelType.XLSX:
                return new XSSFWorkbook();
        }
        return null;
    }


}
