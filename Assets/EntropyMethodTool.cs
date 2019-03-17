using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication1
{
    public class EntropyMethodTool
    {
        // 源数据
        private DataTable sourceTable;

        // 目标数据
        private DataTable targetTable;

        // 名称列表
        private Dictionary<string, double> columnNameDictionary;

        // 构造函数
        public EntropyMethodTool(DataTable sourceTable)
        {
            this.sourceTable = sourceTable;
            this.targetTable = new DataTable();
            this.columnNameDictionary = new Dictionary<string, double>();
        }

        // 获取属性列表
        private void GetColumnNameDictionary()
        {
            for (int i = 1; i < sourceTable.Columns.Count; i++)
            {
                columnNameDictionary.Add(sourceTable.Columns[i].ColumnName, 0.0);
                for (int j = 0; j < sourceTable.Rows.Count; j++)
                {
                    double number = Convert.ToDouble(sourceTable.Rows[j][i].ToString());
                    columnNameDictionary[sourceTable.Columns[i].ColumnName] += number;
                }
            }
        }

        // 步骤一：计算权重
        private void CalculateWeight()
        {
            for (int i = 0; i < sourceTable.Columns.Count; i++)
            {
                if (i == 0)
                {
                    targetTable.Columns.Add(sourceTable.Columns[i].ColumnName, typeof(string));
                }
                else
                {
                    targetTable.Columns.Add(sourceTable.Columns[i].ColumnName, typeof(double));
                }
            }

            for (int i = 0; i < sourceTable.Rows.Count; i++)
            {
                var row = targetTable.NewRow();
                for (int j = 1; j < sourceTable.Columns.Count; j++)
                {
                    string columnName = sourceTable.Columns[j].ColumnName;
                    double number = Convert.ToDouble(sourceTable.Rows[i][j].ToString()) / columnNameDictionary[columnName];
                    row[j] = number * Math.Log(number);
                }
                targetTable.Rows.Add(row);
            }

            double K = -1 / Math.Log(sourceTable.Rows.Count);
            for (int i = 1; i < targetTable.Columns.Count; i++)
            {
                string columnName = targetTable.Columns[i].ColumnName;
                columnNameDictionary[columnName] = 0.0;

                for (int j = 0; j < targetTable.Rows.Count; j++)
                {
                    double number = Convert.ToDouble(targetTable.Rows[j][i].ToString());
                    columnNameDictionary[columnName] += number;
                }

                columnNameDictionary[columnName] *= K;
                columnNameDictionary[columnName] = 1 - columnNameDictionary[columnName];
            }

            double sum = 0.0;
            foreach (KeyValuePair<string, double> kvp in columnNameDictionary)
            {
                sum += kvp.Value;
            }

            List<string> keys = new List<string>(columnNameDictionary.Keys);
            foreach (string key in keys)
            {
                double number = Math.Round(columnNameDictionary[key] / sum, 2);
                columnNameDictionary[key] = number;
            }
        }

        // 步骤二：综合打分
        public Dictionary<string, double> CalculateScore()
        {
            GetColumnNameDictionary();
            CalculateWeight();

            Dictionary<string, double> dictionary = new Dictionary<string, double>();
            foreach (DataRow row in sourceTable.Rows)
            {
                dictionary.Add(row[0].ToString(), 0.0);
            }

            List<double> score = new List<double>();
            for (int i = 0; i < sourceTable.Rows.Count; i++)
            {
                double sum = 0.0;
                for (int j = 1; j < sourceTable.Columns.Count; j++)
                {
                    string columnName = sourceTable.Columns[j].ColumnName;
                    sum += Convert.ToDouble(sourceTable.Rows[i][j].ToString()) * columnNameDictionary[columnName];
                }
                score.Add(sum);
            }

            List<string> keys = new List<string>(dictionary.Keys);
            for (int i = 0; i < keys.Count; i++)
            {
                dictionary[keys[i]] = score[i];
            }
            return dictionary;
        }
    }
}