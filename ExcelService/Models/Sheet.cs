﻿using ExcelService.Extensions;
using ExcelService.Models.Styles;
using System.Linq.Expressions;
using System.Reflection;

namespace ExcelService.Models
{
    public class Sheet
    {
        public Sheet(Row headerRow, IEnumerable<Row> rows, string name, IEnumerable<Table>? tables = null)
        {
            HeaderRow = headerRow;
            Rows = rows;
            Name = name;
            Tables = tables ?? new List<Table>();
        }
        public Row HeaderRow { get; private set; }
        public IEnumerable<Row> Rows { get; private set; }
        public string Name { get; private set; }
        public IEnumerable<Table>? Tables { get; private set; }


        public static Sheet GetSheetFromDataSet<T>(IEnumerable<T> objects, IEnumerable<IEnumerable<Style>>? styles, string sheetName, IEnumerable<Table>? tables = null)
        {
            List<Row> rows = new List<Row>();
            string firstProperty = string.Empty;
            List<Cell> cells = new List<Cell>();
            foreach (PropertyInfo property in typeof(T).GetProperties().Where(p => p.CanRead))
            {
                if (property.Name == firstProperty)
                {
                    break;
                }
                firstProperty = firstProperty == string.Empty ? property.Name : firstProperty;
                cells.Add(new Cell(property.Name));
            }
            for (int i =0; i < objects.Count(); i++)
            {
                rows.Add(Row.GenerateRowFromObject(objects.ElementAt(i), styles?.ElementAt(i+1)));
            }

            return new Sheet(new Row(cells, styles?.ElementAt(0)), rows, sheetName, tables);
        }
        public Cell this[uint x, uint y]
        {
            get => GetCell(x, y);
            set => SetCell(x, y, value);
        }
        public Cell this[char x, uint y]
        {
            get => GetCell(x, y);
            set => SetCell(x, y, value);
        }

        public Cell GetCell(uint x, uint y)
        {
            return Rows.ElementAt((int)x).Cells.ElementAt((int)y);
        }
        public Cell GetCell(char x, uint y)
        {
            if ((int)x < 65 || (int)x > 90)
            {
                throw new InvalidOperationException("First Argument must be argument of type A-Z");
            }
            return Rows.ElementAt((int)y).Cells.ElementAt((int)(x - 65));
        }

        public void SetCell(uint x, uint y, Cell value)
        {
            Rows.ElementAt((int)x).Cells.ElementAt((int)y).SetCell(value);
        }
        public void SetCell(char x, uint y, Cell value)
        {
            if ((int)x < 65 || (int)x > 90)
            {
                throw new InvalidOperationException("First Argument must be argument of type A-Z");
            }
            Rows.ElementAt((int)y).Cells.ElementAt((int)(x - 65)).SetCell(value);
        }
        public void StyleRowWhere<T>(Expression<Func<T, bool>> expression, Style style)
        {
            foreach (Row row in Rows)
            {
                if (expression.Compile().Invoke(row.ObjectReference))
                {
                    foreach (Cell cell in row.Cells)
                    {
                        cell.SetStyle(style);
                    }
                }
            }
        }
        public void StyleCellWhere<T>(Expression<Func<T, bool>> expression, Style style)
        {
            foreach (Row row in Rows)
            {
                if (expression.Compile().Invoke(row.ObjectReference))
                {
                    foreach (Cell cell in row.Cells)
                    {
                        string? resolvedArgString = expression.ResolveArgs();
                        if (resolvedArgString is not null && typeof(T).GetProperty(resolvedArgString)?.GetValue(row.ObjectReference) == cell.Data)
                        {
                            cell.SetStyle(style);
                        }
                    }
                }
            }
        }
        public void AddTable(Table table) => Tables = Tables?.Concat(new List<Table>() { table });
    }
}
