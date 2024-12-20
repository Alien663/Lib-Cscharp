# Data Extension

This library is made to convet DataTable and Class Model.
Thus, I add Segment and Tokenization here.

## DataTable and Class Model Transfer

* DataTable to Class Model auto mapping

```csharp
List<Student> dmData = (List<Student>)dtStudent.ToList<Student>();
```

* DataTable to Calss Model with Mapping(Class Model, DataTable)

```csharp
List<Student> dmData = (List<Student>)dtStudent.ToList<Student>(
    new Dictionary<string, string>
    {
        { "Name", "Name" },
        { "StudentId", "ID" },
        { "Age", "Age" },
    });
```

* Class Model to DataTable

```csharp
DataTable dtData = ClassModelConvert.ToDataTable(dmStudent);
```

## Structure A Context

### Segment

It use punctuation to split an article to segments.
Better to see sample code in unit test.

### Token

Use N-Gram to split tokens out.
Better to see sample code in unit test.

```csharp
string test = @"蘇子與客泛舟遊於赤壁之下";
List<TokenModel> _result = ContextIndexing.Tokenize(test);
```

