using ExcelParserLibrary;
using ExcelParserLibrary.Test;

using TestingConsole.Model;

namespace TestingConsole;

internal class Program
{
   static void Main(string[] args)
   {
      Console.WriteLine("Excel Parser Testing Console");

      string path = @"F:\Electrical\PartInvoices\Mouser\260150809.xls";
      Console.WriteLine($"Parsing {Path.GetFileName(path)} File.");

      //NPOIParser.OpenFile(path);

      ExcelParser parser = new();
      var results = parser.ParseFile<MouserPartModel>(path);

      foreach (var item in results.Data)
      {
         Console.Write('\t');
         Console.WriteLine(item);
      }
   }
}