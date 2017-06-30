using System.Collections;
using System.Collections.Generic;
using UnityEngine;

public class TableOriginCell
{
    public object objValue;
    public string strValue;
    public int row;
    public int column;
    public byte[] rgb;

    //用于Excel Formula
    public string originString;

    public bool isDirty;

    public ExcelValueType valueType;

    public Color bgColor
    {
        get
        {
            if (rgb == null)
                return Color.white;
            return new Color(rgb[0] / 255f, rgb[1] / 255f, rgb[2] / 255f);
        }
    }
}
