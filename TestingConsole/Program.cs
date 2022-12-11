using ExcelParserLibrary;

using TestingConsole.Model;

namespace TestingConsole;

internal class Program
{
   static void Main(string[] args)
   {
      Console.WriteLine("Excel Parser Testing Console");

      Console.WriteLine("Enter file path:");
      string? path = Console.ReadLine();
      if (path == null)
      {
         Console.WriteLine("No path provided...");
         return;
      }
      Console.WriteLine($"Parsing {Path.GetFileName(path)} File.");

      ExcelParser parser = new();
      var results = parser.ParseFile<MouserPartModel>(path);

      foreach (var item in results.Data)
      {
         Console.Write('\t');
         Console.WriteLine(item);
      }
   }
}