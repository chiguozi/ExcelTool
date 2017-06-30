using System.Collections;
using System.Collections.Generic;
using UnityEngine;

public interface IExcelLoader
{
    ExcelOriginData Load(string fullPath);
}
