using System.Collections;
using System.Collections.Generic;
using UnityEngine;

public class ExcelOriginData
{
    public string fullPath;
    public string fileName;
    public List<TableOriginData> tabDataList = new List<TableOriginData>();

    public ExcelOriginData(string fullPath)
    {
        this.fullPath = fullPath;
        fileName = ExcelUtil.GetFileNameWithoutExtension(fullPath);
    }

    public void AddTableData(TableOriginData data)
    {
        tabDataList.Add(data);
    }
}
