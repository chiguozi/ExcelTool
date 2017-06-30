using System.Collections;
using System.Collections.Generic;
using UnityEngine;

//excel原始数据 不做任何删减 空行使用空列表占位
public class TableOriginData
{
    public string tabName;
    public List<List<TableOriginCell>> cellList = new List<List<TableOriginCell>>();
    public int LastRowNum;

    //如果excelrow为空 也需要添加一个空数组来占位
    public void AddRow()
    {
        cellList.Add(new List<TableOriginCell>());
    }

    public void AddCell(int row, TableOriginCell cell)
    {
        if(row >= cellList.Count)
        {
            Debug.LogError("Row Error");
            return;
        }
        cellList[row].Add(cell);
    }

    public List<TableOriginCell> GetRow(int row)
    {
        if (row >= cellList.Count)
            return null;
        return cellList[row];
    }
}
