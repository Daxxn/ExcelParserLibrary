using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelParserLibrary.Models;

/// <summary>
/// Ignore while parsing Excel file.
/// </summary>
[AttributeUsage(AttributeTargets.Property, Inherited = false, AllowMultiple = false)]
public sealed class ExcelIgnoreAttribute : Attribute
{
   /// <summary>
   /// Ignore while parsing Excel file.
   /// </summary>
   public ExcelIgnoreAttribute() { }
}
