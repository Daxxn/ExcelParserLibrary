using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ExcelParserLibrary.Models;

internal class PropertyMap
{
   #region Local Props
   public ExcelProp this[int index] => Map[index];
   public int MaxSize { get; init; }
   Dictionary<int, ExcelProp> Map { get; set; }
   #endregion

   #region Constructors
   public PropertyMap(int maxSize)
   {
      MaxSize = maxSize;
      Map = new(maxSize);
   }
   #endregion

   #region Methods
   public bool ContainsKey(int index) => Map.ContainsKey(index);
   public void AddProp(int index, PropertyInfo info, string name) => Map[index] = new(info, name);
   public ExcelProp? GetProp(int index) => Map.TryGetValue(index, out var prop) ? prop : null;
   #endregion

   #region Full Props

   #endregion
}
