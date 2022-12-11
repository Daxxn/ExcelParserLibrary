using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelParserLibrary.Models;

/// <summary>
/// Excel file property name.
/// <para/>
/// Renames the property to match a property in the Excel file.
/// </summary>
[AttributeUsage(AttributeTargets.Property, Inherited = false, AllowMultiple = false)]
public sealed class ExcelPropertyAttribute : Attribute
{
   private readonly string _propString;

   /// <summary>
   /// Excel file property name.
   /// <para/>
   /// Renames the property to match a property in the Excel file.
   /// </summary>
   /// <param name="propertyString">Name of the property</param>
   public ExcelPropertyAttribute(string propertyString)
   {
      _propString = propertyString;
   }

   /// <summary>
   /// Only used for debugging.
   /// </summary>
   public override string ToString() => $"ExcelProperty {PropertyString}";

   /// <summary>
   /// Assigned property name
   /// </summary>
   public string PropertyString => _propString;
}
