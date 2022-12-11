namespace ExcelParserLibrary.Models;

/// <summary>
/// Results from Excel file parse. Including data and any errors.
/// </summary>
/// <typeparam name="T">The model representing one line in the file.</typeparam>
public class ExcelResult<T>
{
   /// <summary>
   /// All data parsed from the file.
   /// </summary>
   public IEnumerable<T> Data { get; init; }
   /// <summary>
   /// Errors that occured while parsing.
   /// </summary>
   public IEnumerable<Exception>? Errors { get; init; }

   /// <inheritdoc/>
   public ExcelResult(IEnumerable<T> data) => Data = data;
   /// <inheritdoc/>
   public ExcelResult(IEnumerable<T> data, IEnumerable<Exception>? errors)
   {
      Data = data;
      Errors = errors;
   }
}
