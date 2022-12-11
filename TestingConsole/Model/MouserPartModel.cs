using ExcelParserLibrary.Models;

namespace TestingConsole.Model;

public class MouserPartModel
{
   [ExcelProperty("Mouser #:")]
   public string SupplierPartNumber { get; set; } = null!;
   [ExcelProperty("Mfr. #:")]
   public string PartNumber { get; set; } = null!;
   [ExcelProperty("Desc.:")]
   public string? Description { get; set; }
   [ExcelProperty("Customer #")]
   public string? Reference { get; set; }
   [ExcelProperty("Order Qty.")]
   public uint OrderQuantity { get; set; }
   [ExcelProperty("Price (USD)")]
   public decimal UnitPrice { get; set; }

   [ExcelIgnore]
   public decimal TotalPrice => UnitPrice * OrderQuantity;

   public override string ToString() => $"{OrderQuantity} {PartNumber} {SupplierPartNumber} - {UnitPrice} {TotalPrice}";
}
