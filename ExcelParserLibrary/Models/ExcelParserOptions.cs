using ExcelParserLibrary.Models.Exceptions;

namespace ExcelParserLibrary.Models;

/// <summary>
/// Options for Excel parsing.
/// </summary>
public class ExcelParserOptions
{
   /// <summary>
   /// Ignores property case.
   /// </summary>
   public bool IgnoreCase { get; set; } = false;
   /// <summary>
   /// Ignore errors while parsing each row.
   /// <list type="table">
   ///   <item>
   ///      <term>True</term>
   ///      <description>Any errors are added to the error output and parsing continues.</description>
   ///   </item>
   ///   <item>
   ///      <term>False</term>
   ///      <description>An error while parsing throws an <see cref="ExcelPropertyException"/> immedaitely.</description>
   ///   </item>
   /// </list>
   /// </summary>
   public bool IgnorePropertyErrors { get; set; } = false;
   /// <summary>
   /// Row index of the property header.
   /// <para/>
   /// Default = 0
   /// </summary>
   public int HeaderRow { get; set; } = 0;
   /// <summary>
   /// Row start index of the data.
   /// <para/>
   /// Default = 1
   /// </summary>
   public int DataStartRow { get; set; } = 1;
   /// <summary>
   /// Exclusion functions used to filter lines that don't contain useful data.
   /// </summary>
   public IEnumerable<Func<string[], bool>>? ExclusionFunctions { get; set; }
}
