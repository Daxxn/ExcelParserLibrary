using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelParserLibrary.Models;

public class ExcelParserOptions
{
   public bool IgnorePropertyErrors { get; set; } = true;
   public int HeaderRow { get; set; } = 1;
   public int DataStartRow { get; set; } = 2;
   public int MaxLength { get; set; } = int.MaxValue;
}
