using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelParserLibrary.Models
{
   [AttributeUsage(AttributeTargets.Property, Inherited = false, AllowMultiple = false)]
   public sealed class ExcelPropertyAttribute : Attribute
   {
      private readonly string _propString;

      // This is a positional argument
      public ExcelPropertyAttribute(string propertyString)
      {
         _propString = propertyString;
      }

      public bool CompareProperty(string prop, bool ignoreCase) =>
         ignoreCase ? prop.ToLower() == _propString.ToLower() : prop == _propString;

      public override string ToString() => $"ExcelProperty {PropertyString}";

      public string PropertyString => _propString;
   }
}
