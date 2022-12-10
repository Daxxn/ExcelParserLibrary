using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelParserLibrary.Models;

[AttributeUsage(AttributeTargets.Property, Inherited = false, AllowMultiple = false)]
public sealed class ExcelIgnoreAttribute : Attribute
{
   readonly bool _ignoreProp;

   public ExcelIgnoreAttribute(bool ignoreProp = true)
   {
      _ignoreProp = ignoreProp;
   }

    public bool IgnoreProperty => _ignoreProp;
}
