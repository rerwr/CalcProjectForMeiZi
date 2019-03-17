using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using ExcelDataReader;
//using LinqToExcel;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using UnityEngine;
//using Excel = Microsoft.Office.Interop.Excel;
/// <summary>
/// 行业计算服务
/// </summary>
public interface IndustryService
{
    /// <summary>
    /// 获得制造业和生产服务业第J项指标t年的数值
    /// </summary>
    /// <param name="jtarget"></param>
    /// <param name="year"></param>
    /// <returns></returns>
    float GetServiceValue(int jTarget,int tYear);

    float GetMaxTargetValue(int jTarget, int tYear);

    float GetMinTargetValue(int jTarget, int tYear);
    

}

public interface CalcWeight
{
     float GetWeight(int jTarget, int jYear);
}

public interface GetTotalNum
{
    int GetTotalCount();
}

/// <summary>
/// 目前制造业和服务业计算公式一致
/// </summary>
public class ManufacturingService : IndustryService,GetTotalNum {
    public float GetServiceValue(int jTarget, int tYear)
    {
        throw new System.NotImplementedException();
    }

    public float GetMaxTargetValue(int jTarget, int tYear)
    {
        throw new System.NotImplementedException();
    }

    public float GetMinTargetValue(int jTarget, int tYear)
    {
        throw new System.NotImplementedException();
    }

   

    public int GetTotalCount()
    {
        throw new NotImplementedException();
    }
}

public class StandardDataNormalize:CalcWeight
{
    private IndustryService service;

    public StandardDataNormalize(IndustryService service)
    {
        this.service = service;
    }

    /// <summary>
    /// 获得指标标准化数值
    /// </summary>
    /// <param name="jTarget"></param>
    /// <param name="jYear"></param>
    /// <returns></returns>
    public float GetUajTargetStandardValue(int jTarget,int jYear)
    {
        float maxValue= service.GetMaxTargetValue(jTarget, jYear);

        float minValue = service.GetMinTargetValue(jTarget, jYear);

        
        float seriveTarget = service.GetServiceValue(jTarget, jYear);
        float delta= (maxValue - minValue);
        if (Math.Abs(delta) > 0.00001f)
        {
            float serviceTargetActiveStandardValue = (seriveTarget - minValue) / delta;
            return serviceTargetActiveStandardValue;
//            serviceTargetNegativeStandardValue1 = (maxValue - seriveTarget) / delta;
        }
        else
        {
            throw new Exception("被除数不能为0");
        }
    }
    /// <summary>
    /// 获得制造业和生产性服务业在t年的综合发展水平
    /// </summary>
    /// <param name="t"></param>
    /// <returns></returns>
    public float GetUajDevelopLevelCell(int tTarget,int year)
    {
      
        float value= GetUajTargetStandardValue(tTarget, year);
        return GetWeight(tTarget,year)* value;

    }

    /// <summary>
    /// 第 t 年的总体综合发展水平
    /// </summary>
    /// <param name="jYear"></param>
    /// <returns></returns>
    public float GetUajTotalDevelopLevel(int jYear)
    {
        var getTotalNum= (GetTotalNum)service;
        if (getTotalNum != null)
        {
            return -1;
        }
        float UajTotalDevelopLevel=0;
        for (int i = 0; i < getTotalNum.GetTotalCount(); i++)
        {
            UajTotalDevelopLevel += GetUajDevelopLevelCell(i, jYear);
        }
        if (Math.Abs(UajTotalDevelopLevel) < 0.0001)
        {
            Debug.LogError("计算错误");
        }
        return UajTotalDevelopLevel;
    }


    public float GetWeight(int jTarget,int jYear)
    {
       return GetUajTargetStandardValue(jTarget,jYear);
    }

}

public class GetCouplingFactor
{

    StandardDataNormalize standardForFactory = new StandardDataNormalize(new ManufacturingService());
    StandardDataNormalize standardForServices = new StandardDataNormalize(new ManufacturingService());

    /// <summary>
    /// 计算耦合度
    /// </summary>
    /// <returns></returns>
    public float GetCouplingFactor1(int tYear)
    {
        float serviceValue = standardForServices.GetUajTotalDevelopLevel(tYear);
        float standardValue = standardForFactory.GetUajTotalDevelopLevel(tYear);


        float C = serviceValue * standardValue / (serviceValue + standardValue);
        return C;
    }
    static string path= Application.dataPath+"/zl.xlsx";
    public static void Linq2WorkSheet()
    {
//        var execelfile = new ExcelQueryFactory(path);
//        var tsheet = execelfile.Worksheet(0);
        
       
    }

//    public static  void OpenExcel()
//    {
//        object missing = Missing.Value;
//        Excel.Application excel = new Excel.Application();//启动excel程序
//        try
//        {
//            if (excel == null)
//            {
//                Debug.Log("无法访问Excel程序，请重新安装Microsoft Office Excel。");
//            }
//            else
//            {
//                excel.Visible = false;//设置调用引用的Excel文件是否可见
//                excel.UserControl = true;//设置调用引用的Excel是由用户创建或打开的
//                // 以只读的形式打开EXCEL文件（工作簿）想了解这堆参数请访问https://msdn.microsoft.com/zh-cn/library/office/microsoft.office.interop.excel.workbooks.open.aspx
//                Excel.Workbook wb = excel.Application.Workbooks.Open(path, missing, true, missing, missing, missing,
//                    missing, missing, missing, true, missing, missing, missing, missing, missing);
//                //取得第一个工作表
//                Excel.Worksheet ws = (Excel.Worksheet)wb.Worksheets[1];//索引从1开始 
//                //取得总记录行数(包括标题列)  
//                int rowCount = ws.UsedRange.Cells.Rows.Count; //得到行数 
//                int colCount = ws.UsedRange.Cells.Columns.Count;//得到列数
//                //初始化datagridview1
//              
//                //取得第一行，生成datagridview标题列(下标是从1开始的)
//                StringBuilder sb=new StringBuilder();
//                for (int i = 1; i <= colCount; i++)
//                {
//                    string cellStr = ws.Cells[1, i].ToString().Trim();
//                    sb.Append(cellStr);
//                }
//                 Debug.Log(sb);
//
//            }
//        }
//        catch (Exception ex)
//        {
//            Debug.Log("读取Excel文件失败： " + ex.Message);
//        }
//        finally
//        {
////            CloseExcel(excel);//关闭Excel进程
//        }
//    }

    public static void ReadStream()
    {
        using (var stream = File.Open(path, FileMode.Open, FileAccess.Read))
        {

            // Auto-detect format, supports:
            //  - Binary Excel files (2.0-2003 format; *.xls)
            //  - OpenXml Excel files (2007 format; *.xlsx)
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {

                // Choose one of either 1 or 2:

                // 1. Use the reader methods
                do
                {
                    while (reader.Read())
                    {
                        Debug.Log(reader.Name);
                        var cells= reader.MergeCells;
                        // reader.GetDouble(0);
                    }
                } while (reader.NextResult());

                // 2. Use the AsDataSet extension method
                var result = reader.AsDataSet();

                // The result of each spreadsheet is in result.Tables
            }
        }
    }

    public static void GetData()
    {
        FileInfo newFile = new FileInfo(path);
        using (ExcelPackage xlPackage = new ExcelPackage(newFile)) //如果mynewfile.xlsx存在，就打开它，否则就在该位置上创建
        {
            // get the first worksheet in the workbook
           
            var worksheets = xlPackage.Workbook.Worksheets;
            var sheet= worksheets["Sheet1"];
//            for (int i = 0; i < worksheets.Count; i++)
                foreach (var workSheet in worksheets)
                {
//                    var workSheet = worksheets[i];
                    int maxColumnNum = workSheet.Dimension.End.Column;//最大列
                    int minColumnNum = workSheet.Dimension.Start.Column;//最小列


                    int maxRowNum = workSheet.Dimension.End.Row;//最小行
                    int minRowNum = workSheet.Dimension.Start.Row;//最大行

                    Debug.Log("-------------输出名字-------------->" + workSheet.Name);

                    for (int j = 1; j <=maxRowNum; j++)
                    {
                        StringBuilder row = new StringBuilder();
                        row.Append($"第{j}行");
                        for (int k = 1; k <= maxColumnNum; k++)
                        {
                            var value1 = workSheet.Cells[j, k];
                            var value = value1.Value;
                            string valueStr = string.Format("Cell({0},{1}).Value={2}", j, k,
                                value); //循环取出单元格值
                            //string value_Str = string.Format("Cell({0},{1}).Formula={2}", 6, iCol, worksheets[i].Cells[6, iCol].Formula);//取公式（失败了）
                            row.Append(valueStr);
                            Debug.Log(valueStr);
                        }
                    }
                }
           

        }
    }
}

