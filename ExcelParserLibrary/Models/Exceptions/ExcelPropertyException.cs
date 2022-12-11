namespace ExcelParserLibrary.Models.Exceptions;

/// <summary>
/// Error while mapping a model property to an excel property.
/// </summary>
public class ExcelPropertyException : Exception
{
   /// <summary>
   /// Mapped Excel property index.
   /// </summary>
   public int Index { get; set; }
   /// <summary>
   /// Name of the Excel property.
   /// </summary>
   public string ColumnName { get; set; }
   /// <summary>
   /// Type of the column property.
   /// </summary>
   public Type Type { get; set; }

   /// <summary>
   /// Error while mapping a model property to an excel property.
   /// </summary>
   public ExcelPropertyException(
      int index, string columnName, Type type
      ) : base("Unable to map property.")
   {
      Index = index;
      ColumnName = columnName;
      Type = type;
   }

   /// <summary>
   /// Error while mapping a model property to an excel property.
   /// </summary>
   public ExcelPropertyException(
      int index, string columnName, Type type,
      string? message
      ) : base(message)
   {
      Index = index;
      ColumnName = columnName;
      Type = type;
   }

   /// <summary>
   /// Error while mapping a model property to an excel property.
   /// </summary>
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
