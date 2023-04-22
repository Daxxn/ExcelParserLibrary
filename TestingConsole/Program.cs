using ExcelParserLibrary;

using TestingConsole.Model;

namespace TestingConsole;

internal class Program
{
   public static string? FixedFile { get; set; }
   static void Main(string[] args)
   {
      FixedFile = @"F:\Electrical\PartInvoices\Mouser\260150809.xls";
      Console.WriteLine("Excel Parser Testing Console");

      Console.WriteLine("Enter file path:");
      string? path;
      if (FixedFile is null)
      {
         path = Console.ReadLine();
         if (path == null)
         {
            Console.WriteLine("No path provided...");
            return;
         }
         Console.WriteLine($"Parsing {Path.GetFileName(path)} File.");
      }
      else
      {
         path = FixedFile;
      }

      ExcelParser parser = new();
      var results = parser.ParseFile<MouserPartModel>(path);

      foreach (var item in results.Data)
      {
         Console.Write('\t');
         Console.WriteLine(item);
      }
   }
}