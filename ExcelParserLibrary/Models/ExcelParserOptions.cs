using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using ExcelParserLibrary.Models.Exceptions;

namespace ExcelParserLibrary.Models;

/// <summary>
/// Options for Excel parsing.
/// </summary>
public class ExcelParserOptions
{
   /// <summary>
   /// Ignores letter casing.
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
   public bool IgnorePropertyErrors { get; set; } = true;
   /// <summary>
   /// Row index of the property header.
   /// </summary>
   public int HeaderRow { get; set; } = 0;
   /// <summary>
   /// Row start index of the data.
   /// </summary>
   public int DataStartRow { get; set; } = 1;
   /// <summary>
   /// Exclusion functions used to filter lines that don't contain useful data.
   /// </summary>
   public IEnumerable<Func<string[], bool>>? ExclusionFunctions { get; set; }
}
