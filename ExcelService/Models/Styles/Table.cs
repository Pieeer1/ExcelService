using ExcelService.Models.Numerics;

namespace ExcelService.Models.Styles
{
    public class Table
    {
        public Table(UIntVector2 startIndexes, UIntVector2 endIndexes)
        {
            StartIndexes = startIndexes;
            EndIndexes = endIndexes;
        }

        public Table(UIntVector2 startIndexes, UIntVector2 endIndexes, string? tableName = null, Border? border = null, string[]? columnNames = null)
        {
            TableName = tableName;
            StartIndexes = startIndexes;
            EndIndexes = endIndexes;
            Border = border;
            ColumnNames = columnNames;
        }

        public string? TableName { get; set; }
        public UIntVector2 StartIndexes { get; set; }
        public UIntVector2 EndIndexes { get; set; }
        public Border? Border { get; set; }
        public string[]? ColumnNames { get; set; }
    }
}
