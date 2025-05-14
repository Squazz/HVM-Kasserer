using ClosedXML.Excel;

namespace HVM_Kasserer
{
    class ExcelHelperMethods
    {
        public static int GetColumnIndex(IXLWorksheet worksheet, string columnHeader)
        {
            var headerRow = worksheet.Row(1); // Assuming headers are in the first row
            var headerCell = headerRow.CellsUsed()
                .FirstOrDefault(cell => cell.GetValue<string>().Trim().Equals(columnHeader, StringComparison.OrdinalIgnoreCase));

            return headerCell?.Address.ColumnNumber ?? -1;
        }
    }
}
