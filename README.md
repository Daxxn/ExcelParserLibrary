# Excel Parser Library

Wrapper library for getting data from Excel files. It's not meant for writing data or handling complex files.

## Features

- Automatic property mapping
- Custom filtering methods

# Usage

Add attributes to any needed properties.

```cs
using ExcelParserLibrary.Models;

/// <summary>
/// Desired model to parse
/// </summary>
public class Model
{
   [ExcelProperty("Prop Name")]
   public string NameThatDoesntMatchFile { get; set; }
   
   public int Number { get; set; }

   [ExcelIgnore]
   public double IgnoredProperty { get; set; }
}
```

### DI Use

Register the parser with the Dependency Injection service.

program.cs
```cs
using ExcelParserLibrary;

var builder = Application.CreateBuilder(args);

// Recommend using a factory for creation.
Services.AddAbstractFactory<IExcelParser, ExcelParser>();
```

Using the library to parse an Excel file.

SomeService.cs
```cs
using ExcelParserLibrary;

public class SomeService
{
   private readonly IAbstractFactory<IExcelParser> _parserFactory;

   public SomeService(IAbstractFactory<IExcelParser> parserFactory) => _parserFactory = parserFactory;

   public Model ParseFile(string path)
   {
      var parser = _parserFactory.Create();
      var result = parser.ParseFile<Model>(path);
      if (result.errors != null)
      {
         // Handle any line errors found during parsing.
      }
      return result.Data;
   }

   public Model ParseFileWithOptions(string path)
   {
      var parser = _parserFactory.Create();
      var options = new ExcelParserOptions()
      {
         IgnoreCase = true,
         IgnorePropertyErrors = true,
      };
      var result = parser.ParseFile<Model>(path, options);
      if (result.errors != null)
      {
         // Handle any line errors found during parsing.
      }
      return result.Data;
   }
}
```

### Normal Use

```cs
using ExcelParserLibrary;

//...

var options = new ExcelParserOptions
{
   IgnoreCase = true,
   ExclusionFunctions = new List<Func<string[], bool>>()
   {
      // Checks for any lines that aren't complete.
      (props) => props.Length <= 4;
   }
}
try
{
   var result = new ExcelParser(options).ParseFile("path/to/file.xlsx");
   var data = result.Data;
   // Do whatever with data...
}
catch (ExcelPropertyException)
{
   // Exception thrown while parsing line.
   // If line errors is not critical, set options.IgnorePropertyErrors to true.
   throw;
}
catch (exception)
{
   // Handle any other exceptions thrown during file parse.
   throw;
}
```