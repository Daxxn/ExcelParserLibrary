using System.Reflection;

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
