namespace Alien.Common.Utility;

public static class ContextIndexing
{
    public static List<TokenModel> Segment(string context, List<string>? customPunctuationMarks = null)
    {
        if (string.IsNullOrEmpty(context))
        {
            throw new ArgumentException("Input context cannot be null or empty.", nameof(context));
        }

        ReadOnlySpan<char> temp = context.AsSpan();
        List<TokenModel> result = new List<TokenModel>();
        List<string> AllPunctuationMarks = customPunctuationMarks?? new List<string>
        {
            ",",".","?","!","，","。","？","！",";",":","：",
            "；","'","\"","(",")","[","]","{","}","（","）",
            "［","］","｛","｝","「","」","『","』","\n",
        };
        List<string> PunctuationMarks = AllPunctuationMarks.Where(p => context.Contains(p)).ToList();
        int ID = 1;
        while (!temp.IsEmpty)
        {
            int min_index = int.MaxValue;
            string mark = "";
            foreach(var item in PunctuationMarks)
            {
                int index = temp.IndexOf(item.AsSpan());
                if(index >=0 && index < min_index)
                {
                    min_index = index;
                    mark = item;
                }
            }

            if (min_index == int.MaxValue)
            {
                min_index = temp.Length;
            }

            if(!temp.Slice(0, min_index).IsWhiteSpace())
            {
                result.Add(new TokenModel
                {
                    ID = ID++,
                    Context = temp.Slice(0, min_index).ToString(),
                    Mark = mark,
                });
            }

            if (min_index == int.MaxValue) break;
            temp = temp.Slice(min_index + 1);
        }
        return result;
    }

    public static List<TokenModel> Tokenize(string context, int window = 6)
    {
        if (string.IsNullOrEmpty(context))
        {
            throw new ArgumentException("Input context cannot be null or empty.", nameof(context));
        }

        ReadOnlySpan<char> temp = context.AsSpan();
        List<TokenModel> result = new List<TokenModel>();
        int ID = 1;

        for (int i = 1; i <= window; i++)
        {
            for (int j = 0; j <= temp.Length - i; j++)
            {
                result.Add(new TokenModel { ID = ID++, Context = temp.Slice(j, i).ToString() });
            }
        }
        return result;
    }
}

public class TokenModel
{
    public int ID { get; set; }
    public string Context { get; set; } = "";
    public string Mark { get; set; } = "";
}