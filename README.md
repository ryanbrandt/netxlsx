# NetXLSX

Generalized XLSX utilities for C#

# Docs

## `XLSXReader`

Utility class for all XLSX parsing/reading

```csharp
XLSXReader _reader = new XLSXReader();
```

## GetRecords

Generic method to parse all data in an XLSX file from a given worksheet into a List of type T.

T must be a model for which each property has an `XLSXMapping` attribute and T must have a no-arg constructor.

```csharp
public class MyModel
{
    [XLSXMapping("Some Property Column", typeof(string))]
    public string SomeProperty { get; set; }

    public MyModel() { }
}
...

List<MyModel> records = _reader.GetRecords<MyModel>("<path to my XLSX file>", "<sheet of interest>");
```

## GetWorksheets

Retrieves all worksheet names in a given XLSX file

```csharp
List<string> sheets = _reader.GetWorksheets("<path to my XLSX file>");
```

## GetXLSXHeaders

Retrieves all column headers in a given XLSX file worksheet

```csharp
List<string> headers = _reader.GetXLSXHeaders("<path to my XLSX file>", "<sheet of interest>");
```
