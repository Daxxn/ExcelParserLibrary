using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelParserLibrary.Models.Exceptions;

public class ExcelPropertyException : Exception
{
   public int Index { get; set; }
   public string ColumnName { get; set; }
   public Type Type { get; set; }

   public ExcelPropertyException(
      int index, string columnName, Type type
      ) : base("Unable to map property.")
   {
      Index = index;
      ColumnName = columnName;
      Type = type;
   }

   public ExcelPropertyException(
      int index, string columnName, Type type,
      string? message
      ) : base(message)
   {
      Index = index;
      ColumnName = columnName;
      Type = type;
   }

   public ExcelPropertyException(
      int index,
      string columnName,
      Type type,
      string? message,
      Exception? innerException
      ) : base(message, innerException)
   {
      Index = index;
      ColumnName = columnName;
      Type = type;
   }
}
