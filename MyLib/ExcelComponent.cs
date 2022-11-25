using System;
using System.IO;
using System.Data;
using System.Reflection;
using System.Collections.Generic;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System.Linq;
namespace MyLib
{
    public class ExcelComponent
    {
        private IWorkbook workbook = null;
        public byte[] export(DataTable source, int headerIndex = 0)
        {
            if (this.workbook is null)
            {
                this.workbook = new XSSFWorkbook();
            }
            this.createSheet(source, 0, headerIndex);
            MemoryStream stream = new MemoryStream();
            this.workbook.Write(stream, false);
            stream.Flush();
            byte[] result = stream.ToArray();
            return result;
        }

        public byte[] export(DataSet source)
        {
            if(this.workbook is null)
            {
                this.workbook = new XSSFWorkbook();
            }
            for (int i = 0; i < source.Tables.Count; i++)
            {
                DataTable dt = source.Tables[i];
                this.createSheet(dt, i);
            }
            MemoryStream stream = new MemoryStream();
            this.workbook.Write(stream, false);
            stream.Flush();
            byte[] result = stream.ToArray();
            return result;
        }

        public byte[] export<T>(List<T> items)
        {
            if (this.workbook is null)
            {
                this.workbook = new XSSFWorkbook();
            }
            ISheet sheet = this.workbook.CreateSheet(typeof(T).Name);
            IRow header = sheet.CreateRow(0);
            PropertyInfo[] Props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            int i = 0;
            foreach (PropertyInfo prop in Props)
            {
                ICell cell = header.CreateCell(i++);
                cell.SetCellValue(prop.Name);
            }
            i = 0;
            foreach (var item in items)
            {
                IRow rows = sheet.CreateRow(++i);
                for (int j = 0; j < Props.Length; j++)
                {
                    ICell cell = rows.CreateCell(j);
                    var the_value = Props[j].GetValue(item, null);
                    cell.SetCellValue(the_value is null? "" : the_value.ToString());
                }
            }
            MemoryStream stream = new MemoryStream();
            this.workbook.Write(stream, false);
            stream.Flush();
            byte[] result = stream.ToArray();
            stream.Close();
            return result;
        }

        public DataTable readFileDT(FileStream fs, int headerIndex = 0)
        {
            if (this.workbook is null)
            {
                this.workbook = new XSSFWorkbook(fs);
            }
            DataTable result = readSheet(0, headerIndex);
            fs.Close();
            return result;
        }

        public DataSet readFileDS(FileStream fs)
        {
            this.workbook = new XSSFWorkbook(fs);
            DataSet result = new DataSet();
            for (int i = 0; i < this.workbook.NumberOfSheets; i++)
            {
                DataTable dt = readSheet(i, 0);
                result.Tables.Add(dt);
            }
            fs.Close();
            return result;
        }

        public List<T> readFileDM<T>(FileStream fs, int sheetIndex = 0, int headerIndex = 0) where T : new ()
        {
            if (this.workbook is null)
            {
                this.workbook = new XSSFWorkbook(fs);
            }
            ISheet sheet = this.workbook.GetSheetAt(sheetIndex);
            IRow row = sheet.GetRow(headerIndex);
            List<T> dmResult = new List<T>();
            List<string> columns = new List<string>();
            string tempname = "";

            for (int i = 0; i < row.LastCellNum; i++)
            {
                string cellValue = row.GetCell(i).ToString();
                columns.Append(cellValue);
            }

            for (int i = headerIndex + 1; i <= sheet.LastRowNum; i++)
            {
                T t = new T();
                PropertyInfo[] propertys = t.GetType().GetProperties();
                row = sheet.GetRow(i);
                if (row != null)
                {
                    foreach (PropertyInfo pi in propertys)
                    {
                        tempname = pi.Name;
                        if (columns.Contains(tempname))
                        {
                            if (!pi.CanWrite) continue;
                            object value = row.GetCell(columns.IndexOf(tempname));
                            if (value != null) pi.SetValue(t, value, null);
                        }
                    }
                    dmResult.Add(t);
                }
            }
            return dmResult;
        }

        private void createSheet(DataTable source, int sheetIndex, int headerIndex = 0)
        {
            ISheet sheet = string.IsNullOrEmpty(source.TableName) ? this.workbook.CreateSheet("Sheet" + sheetIndex.ToString()) : this.workbook.CreateSheet(source.TableName);
            IRow header = sheet.CreateRow(headerIndex);
            for (int i = 0; i < source.Columns.Count; i++)
            {
                ICell cell = header.CreateCell(i);
                cell.SetCellValue(source.Columns[i].ColumnName);
            }

            // data
            for (int i = 0; i < source.Rows.Count; i++)
            {
                IRow rows = sheet.CreateRow(headerIndex + 1 + i);
                for (int j = 0; j < source.Columns.Count; j++)
                {
                    ICell cell = rows.CreateCell(j);
                    cell.SetCellValue(source.Rows[i][j].ToString());
                }
            }
        }

        private DataTable readSheet(int sheetIndex, int headerIndex)
        {
            ISheet sheet = this.workbook.GetSheetAt(sheetIndex);
            IRow row = sheet.GetRow(headerIndex);
            DataTable dt = new DataTable();
            dt.TableName = this.workbook.GetSheetName(sheetIndex);

            for (int i = 0; i < row.LastCellNum; i++)
            {
                string cellValue = row.GetCell(i).ToString();
                dt.Columns.Add(cellValue);
            }

            for (int i = headerIndex + 1; i <= sheet.LastRowNum; i++)
            {
                DataRow dr = dt.NewRow();
                row = sheet.GetRow(i);
                if (row != null)
                {
                    for (int j = 0; j < row.LastCellNum; j++)
                    {
                        if (row.GetCell(j).CellType == NPOI.SS.UserModel.CellType.Formula)
                        {
                            dr[j] = row.GetCell(j).NumericCellValue;
                        }
                        else
                        {
                            dr[j] = row.GetCell(j).ToString();
                        }
                    }
                }
                dt.Rows.Add(dr);
            }
            return dt;
        }
    }
}
