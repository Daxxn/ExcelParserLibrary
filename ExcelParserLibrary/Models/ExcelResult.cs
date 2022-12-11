using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelParserLibrary.Models;

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

   public ExcelResult(IEnumerable<T> data) => Data = data;
   public ExcelResult(IEnumerable<T> data, IEnumerable<Exception>? errors)
   {
      Data = data;
      Errors = errors;
   }
}
