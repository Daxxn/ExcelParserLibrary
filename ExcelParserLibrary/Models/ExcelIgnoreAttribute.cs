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
