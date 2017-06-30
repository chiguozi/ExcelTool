using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using System.IO;

public class ExcelUtil
{

    public static string GetFileNameWithoutExtension(string fullPath)
    {
        return Path.GetFileNameWithoutExtension(fullPath);
    }

    public static string GetFileExtension(string fullPath)
    {
        return Path.GetExtension(fullPath);
    }

    public static ExcelType GetExcelType(string fullPath)
    {
        var ext = GetFileExtension(fullPath);
        if (ext == "xls")
            return ExcelType.XLS;
        if (ext == "xlsx")
            return ExcelType.XLSX;
        return ExcelType.None;
    }
}
