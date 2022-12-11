using System.Collections;
using System.Reflection;

namespace ExcelParserLibrary.Models;

internal class PropertyMap : IEnumerable<KeyValuePair<int, ExcelProp>>
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
   public IEnumerator<KeyValuePair<int, ExcelProp>> GetEnumerator() => Map.GetEnumerator();
   IEnumerator IEnumerable.GetEnumerator() => Map.GetEnumerator();
   #endregion

   #region Full Props

   #endregion
}
