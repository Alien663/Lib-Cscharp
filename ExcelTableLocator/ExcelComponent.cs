using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.ComponentModel;
using System.Data;
using System.Reflection;

namespace ExcelConverter
{
    public class ExcelComponent
    {
        private IWorkbook workbook = new XSSFWorkbook();
        private SheetRange sheetRange = new SheetRange();
        private Dictionary<string, string> DataTypeStyle = new Dictionary<string, string>
        {
            { "UInt16", "#,##0" },
            { "UInt32", "#,##0" },
            { "UInt64", "#,##0" },
            { "Int16", "#,##0" },
            { "Int32", "#,##0" },
            { "Int64", "#,##0" },
            { "Float", "#,##0.00" },
            { "Double", "#,##0.00" },
            { "Decimal", "#,##0.00" },
        };

        public byte[] export(DataTable source)
        {
            this.createSheet(source, 0);
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
            if (this.sheetRange.MinRowIndex >= 0)
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
                for (int j = 0; j < Props.Length; j++)
                {
                    ICell cell = rows.CreateCell(j + this.sheetRange.MinColIndex);
                    if (DataTypeStyle.ContainsKey(Props[j].PropertyType.Name))
                    {
                        ICellStyle _datastyle = workbook.CreateCellStyle();
                        _datastyle.DataFormat = workbook.CreateDataFormat()
                                    .GetFormat(DataTypeStyle[Props[j].PropertyType.Name]);
                        cell.CellStyle = _datastyle;
                    }
                    var the_value = Props[j].GetValue(item, null);
                    switch (Props[j].PropertyType.Name)
                    {
                        case "UInt16":
                            cell.SetCellType(CellType.Numeric);
                            cell.SetCellValue(Convert.ToUInt16(the_value.ToString()));
                            break;
                        case "UInt32":
                            cell.SetCellType(CellType.Numeric);
                            cell.SetCellValue(Convert.ToUInt32(the_value.ToString()));
                            break;
                        case "UInt64":
                            cell.SetCellType(CellType.Numeric);
                            cell.SetCellValue(Convert.ToUInt64(the_value.ToString()));
                            break;
                        case "Int16":
                            cell.SetCellType(CellType.Numeric);
                            cell.SetCellValue(Convert.ToInt16(the_value.ToString()));
                            break;
                        case "Int32":
                            cell.SetCellType(CellType.Numeric);
                            cell.SetCellValue(Convert.ToInt32(the_value.ToString()));
                            break;
                        case "Int64":
                            cell.SetCellType(CellType.Numeric);
                            cell.SetCellValue(Convert.ToInt64(the_value.ToString()));
                            break;
                        case "Boolean":
                            cell.SetCellType(CellType.Boolean);
                            cell.SetCellValue(Convert.ToBoolean(the_value.ToString()));
                            break;
                        case "Float":
                        case "Double":
                        case "Decimal":
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

        public List<T> readFileDM<T>(FileStream fs) where T : new()
        {
            this.workbook = new XSSFWorkbook(fs);
            ISheet sheet = this.workbook.GetSheetAt(this.sheetRange.MinSheetIndex);
            IRow row = sheet.GetRow(this.sheetRange.MinRowIndex);
            List<T> dmResult = new List<T>();
            List<string> columns = new List<string>();

            for (int i = this.sheetRange.MinColIndex; i < (this.sheetRange.MaxColIndex ?? row.LastCellNum); i++)
            {
                columns.Add(row.GetCell(i).ToString());
            }

            for (int i = this.sheetRange.MinRowIndex + 1; i <= (this.sheetRange.MaxRowIndex ?? sheet.LastRowNum); i++)
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
                            ICell cell = columns.IndexOf(pi.Name) == -1 ?
                                row.GetCell(this.sheetRange.MinColIndex + columns.IndexOf(attr.DisplayName)) :
                                row.GetCell(this.sheetRange.MinColIndex + columns.IndexOf(pi.Name));

                            string value = "";
                            switch (cell.CellType)
                            {
                                case CellType.String:
                                    value = cell.StringCellValue;
                                    break;
                                case CellType.Numeric:
                                    value = cell.NumericCellValue.ToString();
                                    break;
                                case CellType.Boolean:
                                    value = cell.BooleanCellValue.ToString();
                                    break;
                                case CellType.Formula:
                                    value = cell.CachedFormulaResultType.ToString();
                                    break;
                                case CellType.Unknown:
                                    value = cell.StringCellValue;
                                    break;
                                case CellType.Error:
                                    value = cell.ErrorCellValue.ToString();
                                    break;

                            }

                            switch (pi.PropertyType.Name)
                            {
                                case "UInt16":
                                    pi.SetValue(t, Convert.ToUInt16(value), null);
                                    break;
                                case "UInt32":
                                    pi.SetValue(t, Convert.ToUInt32(value), null);
                                    break;
                                case "UInt64":
                                    pi.SetValue(t, Convert.ToUInt64(value), null);
                                    break;
                                case "Int16":
                                    pi.SetValue(t, Convert.ToInt16(value), null);
                                    break;
                                case "Int32":
                                    pi.SetValue(t, Convert.ToInt32(value), null);
                                    break;
                                case "Single":
                                    pi.SetValue(t, Convert.ToSingle(value), null);
                                    break;
                                case "Double":
                                    pi.SetValue(t, Convert.ToDouble(value), null);
                                    break;
                                case "Decimal":
                                    pi.SetValue(t, Convert.ToDecimal(value), null);
                                    break;
                                case "Int64":
                                case "BigInteger":
                                    pi.SetValue(t, Convert.ToInt64(value), null);
                                    break;
                                case "Boolean":
                                    pi.SetValue(t, Convert.ToBoolean(value), null);
                                    break;
                                case "DateTime":
                                    pi.SetValue(t, Convert.ToDateTime(value), null);
                                    break;
                                case "DateOnly":
                                    pi.SetValue(t, DateOnly.FromDateTime(Convert.ToDateTime(value)), null);
                                    break;
                                case "TimeOnly":
                                    pi.SetValue(t, TimeOnly.FromDateTime(Convert.ToDateTime(value)), null);
                                    break;
                                case "Guid":
                                    pi.SetValue(t, Guid.Parse(value), null);
                                    break;
                                default:
                                    pi.SetValue(t, value, null);
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
            ISheet sheet = string.IsNullOrEmpty(source.TableName) ?
                this.workbook.CreateSheet("Sheet" + sheetIndex.ToString()) :
                this.workbook.CreateSheet(source.TableName);
            if (this.sheetRange.MinRowIndex >= 0)
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
                for (int j = 0; j < source.Columns.Count; j++)
                {
                    ICell cell = rows.CreateCell(j + this.sheetRange.MinColIndex);
                    if (DataTypeStyle.ContainsKey(source.Columns[j].DataType.Name))
                    {
                        ICellStyle _datastyle = workbook.CreateCellStyle();
                        _datastyle.DataFormat = workbook.CreateDataFormat()
                                    .GetFormat(DataTypeStyle[source.Columns[j].DataType.Name]);
                        cell.CellStyle = _datastyle;
                    }
                    switch (source.Columns[j].DataType.Name)
                    {
                        case "UInt16":
                            cell.SetCellType(CellType.Numeric);
                            cell.SetCellValue(Convert.ToUInt16(source.Rows[i][j].ToString()));
                            break;
                        case "UInt32":
                            cell.SetCellType(CellType.Numeric);
                            cell.SetCellValue(Convert.ToUInt32(source.Rows[i][j].ToString()));
                            break;
                        case "UInt64":
                            cell.SetCellType(CellType.Numeric);
                            cell.SetCellValue(Convert.ToUInt64(source.Rows[i][j].ToString()));
                            break;
                        case "Int16":
                            cell.SetCellType(CellType.Numeric);
                            cell.SetCellValue(Convert.ToInt16(source.Rows[i][j].ToString()));
                            break;
                        case "Int32":
                            cell.SetCellType(CellType.Numeric);
                            cell.SetCellValue(Convert.ToInt32(source.Rows[i][j].ToString()));
                            break;
                        case "Int64":
                            cell.SetCellType(CellType.Numeric);
                            cell.SetCellValue(Convert.ToInt64(source.Rows[i][j].ToString()));
                            break;
                        case "Boolean":
                            cell.SetCellType(CellType.Boolean);
                            cell.SetCellValue(Convert.ToBoolean(source.Rows[i][j].ToString()));
                            break;
                        case "Float":
                        case "Double":
                        case "Decimal":
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

            for (int i = this.sheetRange.MinColIndex; i < (this.sheetRange.MaxColIndex ?? row.LastCellNum); i++)
            {
                string cellValue = row.GetCell(i).ToString();
                dt.Columns.Add(cellValue);
            }
            for (int i = 0; i < (this.sheetRange.MaxRowIndex == null ? sheet.LastRowNum - this.sheetRange.MinRowIndex : this.sheetRange.MaxRowIndex - this.sheetRange.MinRowIndex); i++)
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

        public void setDataTypeStyle(Dictionary<string, string> pairs)
        {
            foreach (string pair in pairs.Keys)
            {
                if (DataTypeStyle.ContainsKey(pair))
                {
                    DataTypeStyle[pair] = pairs[pair];
                }
                else
                {
                    DataTypeStyle.Add(pair, pairs[pair]);
                }
            }
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
