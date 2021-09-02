[![ryanbrandt](https://circleci.com/gh/ryanbrandt/netxlsx.svg?style=svg)](https://app.circleci.com/pipelines/github/ryanbrandt/netxlsx)
<a href="https://codeclimate.com/github/ryanbrandt/netxlsx/maintainability"><img src="https://api.codeclimate.com/v1/badges/7223116e178d83ad84b9/maintainability" /></a>

# NetXLSX

An easier to use .NET XLSX API

# Documentation

## `XLSXReader`

Utility class for XLSX read operations

```csharp
XLSXReader _reader = new XLSXReader();
```

### GetRecords

Generic method to parse all data in an XLSX file from a given worksheet into a List of type T.

Each property within T which we wish to have automapped to the corresponding XLSX column must have an `XLSXMapping` attribute, specifying the associated column name and type

```csharp
public class MyModel
{
    [XLSXMapping("Some Property Column", typeof(string))]
    public string SomeProperty { get; set; }

    public string SomeUnmappedProperty { get; set; }
}
...

List<MyModel> records = _reader.GetRecords<MyModel>("<path to my XLSX file>", "<sheet of interest>");
```

### GetWorksheets

Retrieves all worksheet names in a given XLSX file

```csharp
List<string> sheets = _reader.GetWorksheets("<path to my XLSX file>");
```

## `XLSXWriter`

Utility class for XLSX write operations

### PutRecords

Generic method to write a list of type T to a specified XLSX worksheet.

As with `GetRecords`, each property within T must have an `XLSXMapping` attribute if we wish for it to be written to a column within the spreadsheet.

```csharp
public class MyModel
{
    [XLSXMapping("Some Property Column", typeof(string))]
    public string SomePropertyIWantWritten { get; set; }

    public string SomePropertyIDontWantWritten { get; set; }
}
...

_writer.PutRecords<MyModel>("<path to my XLSX File>", "<sheet of interest>");
```

This method will write to an existing XLSX file, if one is found at the specified path, or create a new one, if one is not found at the specified path.

By default, this method will overwrite a file completely. If you wish to append to an existing file, specify `append = true`

```csharp
_writer.PutRecords<MyModel>("<path to my XLSX File>", "<sheet of interest>", true);
```
