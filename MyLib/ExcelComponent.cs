using System;
using System.IO;
using System.Data;
using System.Reflection;
using System.ComponentModel;
using System.Collections.Generic;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using System.Net.Http.Headers;

namespace MyLib
{
    public class ExcelComponent
    {
        private IWorkbook workbook = new XSSFWorkbook();
        private SheetRange sheetRange = new SheetRange();
        public byte[] export(DataTable source)
        {
            this.createSheet(source, this.sheetRange.MinSheetIndex);
            MemoryStream stream = new MemoryStream();
            this.workbook.Write(stream, false);
            stream.Flush();
            byte[] result = stream.ToArray();
            return result;
        }

        public byte[] export(DataSet source)
        {
            for (int i = this.sheetRange.MinSheetIndex; i < (this.sheetRange.MaxSheetIndex ?? source.Tables.Count); i++)
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
            ISheet sheet = this.workbook.CreateSheet(typeof(T).Name);
            PropertyInfo[] Props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            int i = this.sheetRange.MinColIndex;
            if(this.sheetRange.MinRowIndex >= 0)
            {
                IRow header = sheet.CreateRow(this.sheetRange.MinRowIndex);
                foreach (PropertyInfo prop in Props)
                {
                    ICell cell = header.CreateCell(i++);
                    var attr = prop.GetCustomAttribute<DisplayNameAttribute>(false);
                    if (attr == null)
                        cell.SetCellValue(prop.Name);
                    else
                        cell.SetCellValue(attr.DisplayName);
                }
            }
            i = this.sheetRange.MinRowIndex;
            foreach (var item in items)
            {
                IRow rows = sheet.CreateRow(++i);
                for (int j = 0; j <  Props.Length; j++)
                {
                    ICell cell = rows.CreateCell(j + this.sheetRange.MinColIndex);
                    var the_value = Props[j].GetValue(item, null);
                    switch (Props[j].PropertyType.Name)
                    {
                        case "UInt16":
                        case "UInt32":
                        case "UInt64":
                        case "Int16":
                        case "Int32":
                        case "Int64":
                            ICellStyle _intstyle = workbook.CreateCellStyle();
                            _intstyle.DataFormat = workbook.CreateDataFormat().GetFormat("#,##0");
                            cell.CellStyle = _intstyle;
                            cell.SetCellType(CellType.Numeric);
                            cell.SetCellValue(int.Parse(the_value.ToString()));
                            break;
                        case "Boolean":
                            cell.SetCellType(CellType.Boolean);
                            cell.SetCellValue(Convert.ToBoolean(the_value.ToString()));
                            break;
                        case "Float":
                        case "Double":
                        case "Decimal":
                            ICellStyle _doublestyle = workbook.CreateCellStyle();
                            _doublestyle.DataFormat = workbook.CreateDataFormat().GetFormat("#,##0.00");
                            cell.CellStyle = _doublestyle;
                            cell.SetCellType(CellType.Numeric);
                            cell.SetCellValue(Convert.ToDouble(the_value.ToString()));
                            break;
                        default:
                            cell.SetCellType(CellType.String);
                            cell.SetCellValue(the_value.ToString());
                            break;
                    }
                }
            }
            MemoryStream stream = new MemoryStream();
            this.workbook.Write(stream, false);
            stream.Flush();
            byte[] result = stream.ToArray();
            stream.Close();
            return result;
        }

        public DataTable readFileDT(FileStream fs)
        {
            this.workbook = new XSSFWorkbook(fs);
            DataTable result = readSheet(this.sheetRange.MinSheetIndex);
            fs.Close();
            return result;
        }

        public DataSet readFileDS(FileStream fs)
        {
            this.workbook = new XSSFWorkbook(fs);
            DataSet result = new DataSet();
            for (int i = this.sheetRange.MinSheetIndex; i < (this.sheetRange.MaxSheetIndex ?? this.workbook.NumberOfSheets); i++)
            {
                DataTable dt = readSheet(i);
                result.Tables.Add(dt);
            }
            fs.Close();
            return result;
        }

        public List<T> readFileDM<T>(FileStream fs) where T : new ()
        {
            this.workbook = new XSSFWorkbook(fs);
            ISheet sheet = this.workbook.GetSheetAt(this.sheetRange.MinSheetIndex);
            IRow row = sheet.GetRow(this.sheetRange.MinRowIndex);
            List<T> dmResult = new List<T>();
            List<string> columns = new List<string>();

            for (int i = this.sheetRange.MinColIndex; i < (this.sheetRange.MaxColIndex ?? row.LastCellNum) ; i++)
            {
                columns.Add(row.GetCell(i).ToString());
            }

            for (int i = this.sheetRange.MinRowIndex + 1; i <= (this.sheetRange.MaxRowIndex ?? sheet.LastRowNum) ; i++)
            {
                T t = new T();
                PropertyInfo[] propertys = t.GetType().GetProperties();
                row = sheet.GetRow(i);
                if (row != null)
                {
                    foreach (PropertyInfo pi in propertys)
                    {
                        var attr = pi.GetCustomAttribute<DisplayNameAttribute>(false);
                        if (columns.Contains(pi.Name) || columns.Contains(attr.DisplayName))
                        {
                            if (!pi.CanWrite) continue;
                            ICell value = columns.IndexOf(pi.Name) == -1 ? row.GetCell(columns.IndexOf(attr.DisplayName)) : row.GetCell(columns.IndexOf(pi.Name));
                            switch (value.CellType)
                            {
                                case CellType.Blank:
                                    pi.SetValue(t, "", null);
                                    break;
                                case CellType.Numeric when DateUtil.IsCellDateFormatted(value):
                                    pi.SetValue(t, value.DateCellValue, null);
                                    break;
                                case CellType.Numeric:
                                    pi.SetValue(t, value.NumericCellValue, null);
                                    break;
                                case CellType.Boolean:
                                    pi.SetValue(t, value.BooleanCellValue, null);
                                    break;
                                case CellType.Error:
                                    throw new Exception($"Parse cell of index (${value.Address}) fail");
                                default:
                                    pi.SetValue(t, value.StringCellValue, null);
                                    break;
                            }
                        }
                    }
                    dmResult.Add(t);
                }
            }
            return dmResult;
        }

        private void createSheet(DataTable source, int sheetIndex)
        {
            ISheet sheet = string.IsNullOrEmpty(source.TableName) ? this.workbook.CreateSheet("Sheet" + sheetIndex.ToString()) : this.workbook.CreateSheet(source.TableName);
            if(this.sheetRange.MinRowIndex >= 0)
            {
                IRow header = sheet.CreateRow(this.sheetRange.MinRowIndex);
                for (int i = 0; i < source.Columns.Count; i++)
                {
                    ICell cell = header.CreateCell(i + this.sheetRange.MinColIndex);
                    cell.SetCellValue(source.Columns[i].ColumnName);
                }
            }
            for (int i = 0; i < source.Rows.Count; i++)
            {
                IRow rows = sheet.CreateRow(i + this.sheetRange.MinRowIndex + 1);
                for (int j = 0 ; j < source.Columns.Count; j++)
                {
                    ICell cell = rows.CreateCell(j + this.sheetRange.MinColIndex);
                    var tt = source.Columns[j].ColumnName;
                    switch (source.Columns[j].DataType.Name)
                    {
                        case "UInt16":
                        case "UInt32":
                        case "UInt64":
                        case "Int16":
                        case "Int32":
                        case "Int64":
                            ICellStyle _intstyle = workbook.CreateCellStyle();
                            _intstyle.DataFormat = workbook.CreateDataFormat().GetFormat("#,##0");
                            cell.CellStyle = _intstyle; 
                            cell.SetCellType(CellType.Numeric);
                            cell.SetCellValue(int.Parse(source.Rows[i][j].ToString()));
                            break;
                        case "Boolean":
                            cell.SetCellType(CellType.Boolean);
                            cell.SetCellValue(Convert.ToBoolean(source.Rows[i][j].ToString()));
                            break;
                        case "Double":
                        case "Decimal":
                            ICellStyle _doublestyle = workbook.CreateCellStyle();
                            _doublestyle.DataFormat = workbook.CreateDataFormat().GetFormat("#,##0.00");
                            cell.CellStyle = _doublestyle;
                            cell.SetCellType(CellType.Numeric);
                            cell.SetCellValue(Convert.ToDouble(source.Rows[i][j].ToString()));
                            break;
                        default:
                            cell.SetCellType(CellType.String);
                            cell.SetCellValue(source.Rows[i][j].ToString());
                            break;
                    }
                }
            }
        }

        private DataTable readSheet(int sheetIndex)
        {
            ISheet sheet = this.workbook.GetSheetAt(sheetIndex);
            IRow row = sheet.GetRow(this.sheetRange.MinRowIndex);
            DataTable dt = new DataTable();
            dt.TableName = sheet.SheetName;

            for (int i = this.sheetRange.MinColIndex; i < (this.sheetRange.MaxColIndex ?? row.LastCellNum) ; i++)
            {
                string cellValue = row.GetCell(i).ToString();
                dt.Columns.Add(cellValue);
            }
            for (int i = 0; i < (this.sheetRange.MaxRowIndex == null? sheet.LastRowNum - this.sheetRange.MinRowIndex : this.sheetRange.MaxRowIndex - this.sheetRange.MinRowIndex); i++)
            {
                DataRow dr = dt.NewRow();
                row = sheet.GetRow(i + 1 + this.sheetRange.MinRowIndex);
                if (row != null)
                {
                    for (int j = 0; j < (this.sheetRange.MaxColIndex == null ? row.LastCellNum - this.sheetRange.MinColIndex : this.sheetRange.MaxColIndex - this.sheetRange.MinColIndex); j++)
                    {
                        var cell = row.GetCell(j + this.sheetRange.MinColIndex);
                        switch (cell.CellType)
                        {
                            case CellType.Numeric when DateUtil.IsCellDateFormatted(cell):
                                dr[j] = cell.DateCellValue;
                                break;
                            case CellType.Formula:
                                HSSFFormulaEvaluator eva = new HSSFFormulaEvaluator(workbook);
                                dr[j] = eva.Evaluate(cell).StringValue;
                                break;
                            case CellType.Numeric:
                                dr[j] = cell.NumericCellValue;
                                break;
                            case CellType.Boolean:
                                dr[j] = cell.BooleanCellValue;
                                break;
                            case CellType.Error:
                                dr[j] = cell.ErrorCellValue;
                                break;
                            default:
                                dr[j] = cell.StringCellValue;
                                break;
                        }
                    }
                }
                dt.Rows.Add(dr);
            }
            return dt;
        }

        public void SetRange(SheetRange _sr)
        {
            this.sheetRange.MinSheetIndex = _sr.MinSheetIndex;
            this.sheetRange.MaxSheetIndex = _sr.MaxSheetIndex;
            this.sheetRange.MinRowIndex = _sr.MinRowIndex;
            this.sheetRange.MaxRowIndex = _sr.MaxRowIndex;
            this.sheetRange.MinColIndex = _sr.MinColIndex;
            this.sheetRange.MaxColIndex = _sr.MaxColIndex;
        }
    }

    public class SheetRange
    {
        public int MinSheetIndex { get; set; } = 0;
        public int? MaxSheetIndex { get; set; } = null;
        public int MinRowIndex { get; set; } = 0;
        public int? MaxRowIndex { get; set; } = null;
        public int MinColIndex { get; set; } = 0;
        public int? MaxColIndex { get; set; } = null;
    }
}
