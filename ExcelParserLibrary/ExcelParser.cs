using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Reflection.Metadata;
using System.Text;
using System.Threading.Tasks;

using ExcelParserLibrary.Models;
using ExcelParserLibrary.Models.Exceptions;

using OfficeOpenXml;

namespace ExcelParserLibrary;

public class ExcelParser
{
   #region Local Props
   private ExcelParserOptions _options = new();
   #endregion

   #region Constructors
   public ExcelParser() => ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
   #endregion

   #region Methods

   #region Extended Methods
   public ExcelResult<T> ParseFile<T>(string filePath, ExcelParserOptions options) where T : class, new()
   {
      _options = options;
      using ExcelPackage package = new(filePath);
      return ParseFile<T>(package);
   }
   public ExcelResult<T> ParseFile<T>(string filePath) where T : class, new()
   {
      var fileInfo = new FileInfo(filePath);
      using ExcelPackage package = new(fileInfo);
      return ParseFile<T>(package);
   }
   public ExcelResult<T> ParseFile<T>(Stream stream, ExcelParserOptions options) where T : class, new()
   {
      _options = options;
      using ExcelPackage package = new(stream);
      return ParseFile<T>(package);
   }
   public ExcelResult<T> ParseFile<T>(Stream stream) where T : class, new()
   {
      using ExcelPackage package = new(stream);
      return ParseFile<T>(package);
   }
   #endregion

   private ExcelResult<T> ParseFile<T>(ExcelPackage package) where T : class, new()
   {
      List<T> data = new();
      List<Exception> errors = new();
      var sheet = package.Workbook.Worksheets.First();
      var map = GetPropMap<T>(sheet);

      for (int r = _options.DataStartRow; r < _options.MaxLength; r++)
      {
         try
         {
            T newObject = new();
            for (int c = 0; c < map.MaxSize; c++)
            {
               if (sheet.Cells[r, c].Value is not null)
               {
                  if (map.ContainsKey(r))
                  {
                     if (!SetProperty(sheet.Cells[r, c].Value, map[c].Prop, newObject))
                     {
                        throw new ExcelPropertyException(c, map[c].Name, sheet.Cells[r, c].Value.GetType());
                     }
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
   private PropertyMap GetPropMap<T>(ExcelWorksheet sheet) where T : class, new()
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

            for (int i = 1; i < names.Length; i++)
            {
               if (names[i] == propName)
               {
                  map.AddProp(i, prop, propName);
               }
            }
         }
      }
      return map;
   }

   private (string[] names, int maxLength) GetHeaderNames(ExcelWorksheet sheet)
   {
      List<string> names = new()
      {
         ""
      };
      for (int i = 1; i < sheet.Rows[_options.HeaderRow].Count(); i++)
      {
         if (sheet.Cells[_options.HeaderRow, i].Value is string str)
         {
            names.Add(str);
         }
         else
         {
            return (names.ToArray(), i);
         }
      }
      throw new Exception("Unable to find column names.");
   }

   private bool SetProperty(object data, PropertyInfo prop, object obj)
   {
      if (prop.PropertyType == data.GetType())
      {
         prop.SetValue(obj, data);
         return true;
      }
      return false;
   }
   #endregion
   #endregion

   #region Full Props

   #endregion
}
