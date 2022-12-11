using ExcelParserLibrary.Models;

namespace ExcelParserLibrary;

/// <summary>
/// Excel File Parser
/// </summary>
public interface IExcelParser
{
   /// <summary>
   /// Parse an Excel file and return the desired models.
   /// </summary>
   /// <typeparam name="T">The type of model to create.</typeparam>
   /// <param name="stream">Excel file stream to read.</param>
   /// <returns><see cref="ExcelResult{T}"/> with all models and errors.</returns>
   Task<ExcelResult<T>> ParseFileAsync<T>(Stream stream) where T : class, new();
   /// <summary>
   /// Parse an Excel file and return the desired models.
   /// </summary>
   /// <typeparam name="T">The type of model to create.</typeparam>
   /// <param name="stream">Excel file stream to read.</param>
   /// <param name="options">Extra options to configure parsing.</param>
   /// <returns><see cref="ExcelResult{T}"/> with all models and errors.</returns>
   Task<ExcelResult<T>> ParseFileAsync<T>(Stream stream, ExcelParserOptions options) where T : class, new();
   /// <summary>
   /// Parse an Excel file and return the desired models.
   /// </summary>
   /// <typeparam name="T">The type of model to create.</typeparam>
   /// <param name="filePath">Local path to the file.</param>
   /// <returns><see cref="ExcelResult{T}"/> with all models and errors.</returns>
   Task<ExcelResult<T>> ParseFileAsync<T>(string filePath) where T : class, new();
   /// <summary>
   /// Parse an Excel file and return the desired models.
   /// </summary>
   /// <typeparam name="T">The type of model to create.</typeparam>
   /// <param name="filePath">Local path to the file.</param>
   /// <param name="options">Extra options to configure parsing.</param>
   /// <returns><see cref="ExcelResult{T}"/> with all models and errors.</returns>
   Task<ExcelResult<T>> ParseFileAsync<T>(string filePath, ExcelParserOptions options) where T : class, new();
   /// <summary>
   /// Parse an Excel file and return the desired models.
   /// </summary>
   /// <typeparam name="T">The type of model to create.</typeparam>
   /// <param name="stream">Excel file stream to read.</param>
   /// <returns><see cref="ExcelResult{T}"/> with all models and errors.</returns>
   ExcelResult<T> ParseFile<T>(Stream stream) where T : class, new();
   /// <summary>
   /// Parse an Excel file and return the desired models.
   /// </summary>
   /// <typeparam name="T">The type of model to create.</typeparam>
   /// <param name="stream">Excel file stream to read.</param>
   /// <param name="options">Extra options to configure parsing.</param>
   /// <returns><see cref="ExcelResult{T}"/> with all models and errors.</returns>
   ExcelResult<T> ParseFile<T>(Stream stream, ExcelParserOptions options) where T : class, new();
   /// <summary>
   /// Parse an Excel file and return the desired models.
   /// </summary>
   /// <typeparam name="T">The type of model to create.</typeparam>
   /// <param name="filePath">Local path to the file.</param>
   /// <returns><see cref="ExcelResult{T}"/> with all models and errors.</returns>
   ExcelResult<T> ParseFile<T>(string filePath) where T : class, new();
   /// <summary>
   /// Parse an Excel file and return the desired models.
   /// </summary>
   /// <typeparam name="T">The type of model to create.</typeparam>
   /// <param name="filePath">Local path to the file.</param>
   /// <param name="options">Extra options to configure parsing.</param>
   /// <returns><see cref="ExcelResult{T}"/> with all models and errors.</returns>
   ExcelResult<T> ParseFile<T>(string filePath, ExcelParserOptions options) where T : class, new();
}