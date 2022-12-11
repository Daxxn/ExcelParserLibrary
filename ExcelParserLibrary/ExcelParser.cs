using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Reflection.Metadata;
using System.Text;
using System.Threading.Tasks;

using ExcelParserLibrary.Models;
using ExcelParserLibrary.Models.Exceptions;

using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ExcelParserLibrary;

/// <inheritdoc/>
public class ExcelParser : IExcelParser
{
   #region Local Props
   /// <inheritdoc/>
   private ExcelParserOptions _options = new();
   #endregion

   #region Constructors
   /// <summary>
   /// Default constructor
   /// </summary>
   public ExcelParser() { }
   /// <summary>
   /// Normal constructor.
   /// <para/>
   /// As in, not used with dependency injection.
   /// </summary>
   /// <param name="options">Extra options to use for parsing.</param>
   public ExcelParser(ExcelParserOptions options) => _options = options;
   #endregion

   #region Methods
   #region Extended Methods
   #region Async Methods
   /// <inheritdoc/>
   public async Task<ExcelResult<T>> ParseFileAsync<T>(string filePath, ExcelParserOptions options) where T : class, new()
   {
      _options = options;
      return await Task.Run(() =>
      {
         using FileStream stream = new(filePath, FileMode.Open);
         using IWorkbook package = Path.GetExtension(filePath) == ".xls" ? new HSSFWorkbook(stream) : new XSSFWorkbook(stream);
         return ParseFile<T>(package);
      });
   }
   /// <inheritdoc/>
   public async Task<ExcelResult<T>> ParseFileAsync<T>(string filePath) where T : class, new()
   {
      return await Task.Run(() =>
      {
         using FileStream stream = new(filePath, FileMode.Open);
         using IWorkbook package = Path.GetExtension(filePath) == ".xls" ? new HSSFWorkbook(stream) : new XSSFWorkbook(stream);
         return ParseFile<T>(package);
      });
   }
   /// <inheritdoc/>
   public async Task<ExcelResult<T>> ParseFileAsync<T>(Stream stream, ExcelParserOptions options) where T : class, new()
   {
      _options = options;
      return await Task.Run(() =>
      {
         using HSSFWorkbook package = new(stream);
         return ParseFile<T>(package);
      });
   }
   /// <inheritdoc/>
   public async Task<ExcelResult<T>> ParseFileAsync<T>(Stream stream) where T : class, new()
   {
      return await Task.Run(() =>
      {
         using HSSFWorkbook package = new(stream);
         return ParseFile<T>(package);
      });
   }
   #endregion
   /// <inheritdoc/>
   public ExcelResult<T> ParseFile<T>(string filePath, ExcelParserOptions options) where T : class, new()
   {
      _options = options;
      using FileStream stream = new(filePath, FileMode.Open);
      using IWorkbook package = Path.GetExtension(filePath) == ".xls" ? new HSSFWorkbook(stream) : new XSSFWorkbook(stream);
      return ParseFile<T>(package);
   }
   /// <inheritdoc/>
   public ExcelResult<T> ParseFile<T>(string filePath) where T : class, new()
   {
      using FileStream stream = new(filePath, FileMode.Open);
      using IWorkbook package = Path.GetExtension(filePath) == ".xls" ? new HSSFWorkbook(stream) : new XSSFWorkbook(stream);
      return ParseFile<T>(package);
   }
   /// <inheritdoc/>
   public ExcelResult<T> ParseFile<T>(Stream stream, ExcelParserOptions options) where T : class, new()
   {
      _options = options;
      using HSSFWorkbook package = new(stream);
      return ParseFile<T>(package);
   }
   /// <inheritdoc/>
   public ExcelResult<T> ParseFile<T>(Stream stream) where T : class, new()
   {
      using HSSFWorkbook package = new(stream);
      return ParseFile<T>(package);
   }
   #endregion

   private ExcelResult<T> ParseFile<T>(IWorkbook package) where T : class, new()
   {
      List<T> data = new();
      List<Exception> errors = new();
      var sheet = package.GetSheetAt(0);
      var map = GetPropMap<T>(sheet);

      for (int r = _options.DataStartRow; r < sheet.LastRowNum; r++)
      {
         try
         {
            var rowData = sheet.GetRow(r).Cells;
            T newObject = new();
            foreach (var kv in map)
            {
               if (rowData[kv.Key].CellType != CellType.Unknown)
               {
                  if (_options.ExclusionFunctions != null)
                  {
                     if (_options.ExclusionFunctions.Any(func => func(rowData.Select(cell => cell.StringCellValue).ToArray())))
                        break;
                  }
                  if (!SetProperty(rowData[kv.Key].StringCellValue, kv.Value.Prop, newObject))
                  {
                     throw new ExcelPropertyException(kv.Key, kv.Value.Name, rowData[kv.Key].GetType());
                  }
               }
            }
            data.Add(newObject);
         }
         catch (Exception e)
         {
            if (_options.IgnorePropertyErrors)
            {
               errors.Add(e);
            }
            else
               throw;
         }
      }

      return errors.Count > 0 ? new(data, errors) : new(data);
   }

   #region Util Methods
   /// <summary>
   /// Builds the property map based on the <see cref="ExcelPropertyAttribute"/> assigned in <typeparamref name="T"/>.
   /// </summary>
   /// <typeparam name="T">Model to base parsing on</typeparam>
   /// <param name="sheet"><see cref="ISheet"/> from the Excel file.</param>
   /// <returns>A map of matching properties.</returns>
   private PropertyMap GetPropMap<T>(ISheet sheet) where T : class, new()
   {
      (var names, var maxPropCount) = GetHeaderNames(sheet);
      var map = new PropertyMap(maxPropCount);
      var props = new T().GetType().GetProperties();

      foreach (var prop in props)
      {
         if (prop.GetCustomAttribute<ExcelIgnoreAttribute>() is null && prop.CanWrite)
         {
            string propName;
            var excelProp = prop.GetCustomAttribute<ExcelPropertyAttribute>();
            propName = excelProp is null ? prop.Name : excelProp.PropertyString;

            if (_options.IgnoreCase)
            {
               for (int i = 0; i < names.Length; i++)
               {
                  if (names[i].ToLower() == propName.ToLower())
                  {
                     map.AddProp(i, prop, propName);
                     break;
                  }
               }
            }
            else
            {
               for (int i = 0; i < names.Length; i++)
               {
                  if (names[i] == propName)
                  {
                     map.AddProp(i, prop, propName);
                     break;
                  }
               }
            }
         }
      }
      return map;
   }

   private (string[] names, int maxLength) GetHeaderNames(ISheet sheet)
   {
      List<string> names = new();
      var rowCells = sheet.GetRow(_options.HeaderRow).Cells;
      foreach (var cell in rowCells)
      {
         names.Add(cell.StringCellValue);
      }
      return (names.ToArray(), rowCells.Count);
   }

   private static bool SetProperty(string data, PropertyInfo prop, object obj)
   {
      if (prop.PropertyType == typeof(string))
      {
         prop.SetValue(obj, data);
         return true;
      }
      else if (prop.PropertyType == typeof(int))
      {
         if (int.TryParse(data, out int intNum))
         {
            prop.SetValue(obj, intNum);
            return true;
         }
      }
      else if (prop.PropertyType == typeof(uint))
      {
         if (uint.TryParse(data, out uint uintNum))
         {
            prop.SetValue(obj, uintNum);
            return true;
         }
      }
      else if (prop.PropertyType == typeof(double))
      {
         if (double.TryParse(data, out double doubleNum))
         {
            prop.SetValue(obj, doubleNum);
            return true;
         }
      }
      else if (prop.PropertyType == typeof(decimal))
      {
         if (string.IsNullOrEmpty(data))
         {
            return false;
         }
         if (char.IsSymbol(data[0]))
         {
            if (decimal.TryParse(data.Remove(0, 1), out decimal currencyNum))
            {
               prop.SetValue(obj, currencyNum);
               return true;
            }
         }
         if (decimal.TryParse(data, out decimal decimalNum))
         {
            prop.SetValue(obj, decimalNum);
            return true;
         }
      }
      else if (prop.PropertyType.IsEnum)
      {
         if (int.TryParse(data, out int enumInt))
         {
            var enumCast = Convert.ChangeType(data, prop.PropertyType);
            if (enumCast != null)
            {
               prop.SetValue(obj, enumCast);
               return true;
            }
         }
         if (Enum.TryParse(prop.PropertyType, data, out object? enumObj))
         {
            prop.SetValue(obj, enumObj);
            return true;
         }
      }
      return false;
   }
   #endregion
   #endregion

   #region Full Props

   #endregion
}
