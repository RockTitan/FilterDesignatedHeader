using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;

namespace FilterDesignatedHeader
{
    public static class Extension
    {
        /// <summary>
        /// Convert 2-D object array to List.
        /// </summary>
        /// <param name="cellValues"></param>
        /// <returns></returns>
        public static List<string[]> ToRowList(this object[,] cellValues)
        {
            //取得維度長度
            int arrayLength1 = cellValues.GetLength(0); //Row Count
            int arrayLength2 = cellValues.GetLength(1); //Column Count

            List<string[]> rowDataList = new List<string[]>();
            for (int ii = 0; ii <= (arrayLength1 - 1); ii++)
            {
                string[] rowData = new string[arrayLength2 + 1];
                for (int jj = 0; jj <= (arrayLength2 - 1); jj++)
                {
                    rowData[jj] = cellValues[ii + 1, jj + 1] == null ? string.Empty : cellValues[ii + 1, jj + 1].ToString().Trim();
                }
                rowDataList.Add(rowData);
            }
            return rowDataList;
        }

        /// <summary>
        /// Convert 2-D object array to DataTable. The first row is DataTable column name.
        /// </summary>
        /// <param name="cellValues"></param>
        /// <returns></returns>
        public static DataTable ToDataTable(this object[,] cellValues)
        {
            DataTable dt = new DataTable();

            //取得維度長度
            int arrayLength1 = cellValues.GetLength(0); //Row Count
            int arrayLength2 = cellValues.GetLength(1); //Column Count

            //處理標題列
            for (int i = 0; i < arrayLength2; i++)
            {
                dt.Columns.Add(cellValues[1, i + 1] == null ? string.Empty : cellValues[1, i + 1].ToString().Trim());
            }

            //標題列之後的資料
            for (int i = 1; i <= (arrayLength1 - 1); i++)
            {
                DataRow dr = dt.NewRow();
                for (int j = 0; j <= (arrayLength2 - 1); j++)
                {
                    dr[j] = cellValues[i + 1, j + 1] == null ? string.Empty : cellValues[i + 1, j + 1].ToString().Trim();
                }
                dt.Rows.Add(dr);
            }
            return dt;
        }
    }
}
