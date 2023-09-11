using System.Data;
using System.Reflection;

namespace TableConverter
{
    public static class DataTableExtensions
    {
        public static IList<T> ToList<T>(this DataTable table) where T : new()
        {
            IList<PropertyInfo> properties = typeof(T).GetProperties().ToList();
            IList<T> result = new List<T>();
            foreach (var row in table.Rows)
            {
                var item = CreateItemFromRow<T>((DataRow)row, properties);
                result.Add(item);
            }
            return result;
        }

        public static IList<T> ToList<T>(this DataTable table, Dictionary<string, string> mappings) where T : new()
        {
            IList<PropertyInfo> properties = typeof(T).GetProperties().ToList();
            IList<T> result = new List<T>();
            foreach (var row in table.Rows)
            {
                var item = CreateItemFromRow<T>((DataRow)row, properties, mappings);
                result.Add(item);
            }
            return result;
        }

        private static T CreateItemFromRow<T>(DataRow row, IList<PropertyInfo> properties) where T : new()
        {
            T item = new T();
            foreach (var property in properties)
            {
                if (row.Table.Columns.Contains(property.Name))
                {
                    property.SetValue(item, row[property.Name], null);
                }
            }
            return item;
        }

        private static T CreateItemFromRow<T>(DataRow row, IList<PropertyInfo> properties, Dictionary<string, string> mappings) where T : new()
        {
            T item = new T();
            foreach (var property in properties)
            {
                if (mappings.ContainsKey(property.Name))
                {
                    property.SetValue(item, row[mappings[property.Name]], null);
                }
            }
            return item;
        }
    }
    public class DataModelExtensions
    {
        public DataTable ToDataTable<T>(List<T> items)
        {
            DataTable dataTable = new DataTable(typeof(T).Name);
            //Get all the properties
            PropertyInfo[] Props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            foreach (PropertyInfo prop in Props)
            {
                //Setting column names as Property names
                dataTable.Columns.Add(prop.Name, prop.PropertyType);
            }
            foreach (T item in items)
            {
                var values = new object[Props.Length];
                for (int i = 0; i < Props.Length; i++)
                {
                    //inserting property values to datatable rows
                    values[i] = Props[i].GetValue(item, null);
                }
                dataTable.Rows.Add(values);
            }
            //put a breakpoint here and check datatable
            return dataTable;
        }
    }

    public class Tokenization
    {
        public List<TokenModel> Segment(string context)
        {
            string temp = context;
            List<TokenModel> result = new List<TokenModel>();
            List<PunctuationModel> AllPunctuationMarks = new List<PunctuationModel>
        {
            new PunctuationModel {Marks= ","},
            new PunctuationModel {Marks= "."},
            new PunctuationModel {Marks= "?"},
            new PunctuationModel {Marks= "!"},
            new PunctuationModel {Marks= "，"},
            new PunctuationModel {Marks= "。"},
            new PunctuationModel {Marks= "？"},
            new PunctuationModel {Marks= "！"},
            new PunctuationModel {Marks= ";"},
            new PunctuationModel {Marks= ":"},
            new PunctuationModel {Marks= "："},
            new PunctuationModel {Marks= "；"},
            new PunctuationModel {Marks= "'"},
            new PunctuationModel {Marks= "\""},
            new PunctuationModel {Marks= "("},
            new PunctuationModel {Marks= ")"},
            new PunctuationModel {Marks= "["},
            new PunctuationModel {Marks= "]"},
            new PunctuationModel {Marks= "{"},
            new PunctuationModel {Marks= "}"},
            new PunctuationModel {Marks= "（"},
            new PunctuationModel {Marks= "）"},
            new PunctuationModel {Marks= "［"},
            new PunctuationModel {Marks= "］"},
            new PunctuationModel {Marks= "｛"},
            new PunctuationModel {Marks= "｝"},
            new PunctuationModel {Marks= "「"},
            new PunctuationModel {Marks= "」"},
            new PunctuationModel {Marks= "『"},
            new PunctuationModel {Marks= "』"},
            new PunctuationModel {Marks= "\n"},
        };
            List<PunctuationModel> PunctuationMarks = AllPunctuationMarks.Where(p => context.IndexOf(p.Marks) >= 0).ToList();
            int ID = 1;
            while (temp.Length > 0)
            {
                int min_index = int.MaxValue;
                string mark = "";
                PunctuationMarks.ForEach(item =>
                {
                    int indexof = temp.IndexOf(item.Marks);
                    if (indexof >= 0 && indexof < min_index)
                    {
                        min_index = indexof;
                        mark = item.Marks;
                    }
                });

                if (min_index == int.MaxValue)
                {
                    min_index = temp.Length;
                }
                if (!string.IsNullOrWhiteSpace(temp.Substring(0, min_index)))
                    result.Add(new TokenModel
                    {
                        ID = ID++,
                        Context = temp.Substring(0, min_index),
                        Mark = mark,
                    });
                if (min_index == int.MaxValue) break;
                temp = temp.Substring(min_index + 1);
            }
            return result;
        }

        public List<TokenModel> Tokenize(string context, int window = 6)
        {
            List<TokenModel> result = new List<TokenModel>();
            int ID = 1;
            for (int i = 1; i <= window; i++)
            {
                for (int j = 0; j <= context.Length - i; j++)
                {
                    result.Add(new TokenModel { ID = ID++, Context = context.Substring(j, i) });
                }
            }
            return result;
        }
    }

    public class TokenModel
    {
        public int ID { get; set; }
        public string Context { get; set; }
        public string Mark { get; set; }
    }
    public class PunctuationModel
    {
        public string Marks { get; set; }
        public int index { get; set; } = -999;
    }
}
