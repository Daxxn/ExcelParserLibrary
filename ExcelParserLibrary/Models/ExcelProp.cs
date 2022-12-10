using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ExcelParserLibrary.Models;

internal class ExcelProp
{
   public PropertyInfo Prop { get; set; }
   public string Name { get; set; }

   public ExcelProp(PropertyInfo prop, string name)
   {
      Prop = prop;
      Name = name;
   }
}
