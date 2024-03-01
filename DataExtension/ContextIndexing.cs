using Data.Extension;

namespace Data.Extension
{
    public static class ContextIndexing
    {
        public static List<TokenModel> Segment(string context)
        {
            string temp = context;
            List<TokenModel> result = new List<TokenModel>();
            List<string> AllPunctuationMarks = new List<string>
            {
                ",",".","?","!","，","。","？","！",";",":","：",
                "；","'","\"","(",")","[","]","{","}","（","）",
                "［","］","｛","｝","「","」","『","』","\n",
            };
            List<string> PunctuationMarks = AllPunctuationMarks.Where(p => context.IndexOf(p) >= 0).ToList();
            int ID = 1;
            while (temp.Length > 0)
            {
                int min_index = int.MaxValue;
                string mark = "";
                PunctuationMarks.ForEach(item =>
                {
                    int indexof = temp.IndexOf(item);
                    if (indexof >= 0 && indexof < min_index)
                    {
                        min_index = indexof;
                        mark = item;
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

        public static List<TokenModel> Tokenize(string context, int window = 6)
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
}
