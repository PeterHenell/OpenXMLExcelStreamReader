using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace ExcelReaderTest
{
    public class SheetOperator : IDisposable
    {
        private WorkbookPart workbookPart;
        private OpenXmlReader reader;
        private Sheet currentSheet;
        private DataTable schema;


        public SheetOperator(WorkbookPart workbookPart, OpenXmlReader reader, Sheet currentSheet)
        {
            this.workbookPart = workbookPart;
            this.reader = reader;
            this.currentSheet = currentSheet;
        }

        public void ForEachRow(Action<DataTable, DataRow> forEachRow)
        {
            while (reader.Read())
            {
                if (reader.ElementType == typeof(Row))
                {
                    if (schema == null)
                    {
                        schema = CreateSchema();
                    }
                    else
                    {
                        DataRow dataRow = ProcessOneRow();
                        forEachRow(schema, dataRow);
                    }
                }
            }
        }

        private DataRow ProcessOneRow()
        {
            var dataRow = schema.NewRow();
            var colCount = 0;

            reader.ReadFirstChild();
            do
            {
                if (reader.ElementType == typeof(Cell))
                {
                    Cell c = (Cell)reader.LoadCurrentElement();
                    string cellValue;
                    if (c.DataType != null && c.DataType == CellValues.SharedString)
                    {
                        SharedStringItem ssi = workbookPart.SharedStringTablePart.
                            SharedStringTable.Elements<SharedStringItem>().ElementAt
                            (int.Parse(c.CellValue.InnerText));
                        cellValue = ssi.Text.Text;
                    }
                    else
                    {
                        cellValue = c.CellValue.InnerText;
                    }

                    dataRow[colCount++] = cellValue;
                }

            } while (reader.ReadNextSibling());

            return dataRow;
        }

        /// <summary>
        /// Creates schema from first row(or any row, this method does not care)
        /// </summary>
        /// <returns></returns>
        private DataTable CreateSchema()
        {
            var schema = new DataTable(currentSheet.Name);
            var colCount = 0;
            reader.ReadFirstChild();
            do
            {
                if (reader.ElementType == typeof(Cell))
                {
                    Cell c = (Cell)reader.LoadCurrentElement();
                    string cellValue;
                    if (c.DataType != null && c.DataType == CellValues.SharedString)
                    {
                        SharedStringItem ssi = workbookPart.SharedStringTablePart.
                            SharedStringTable.Elements<SharedStringItem>().ElementAt
                            (int.Parse(c.CellValue.InnerText));
                        cellValue = ssi.Text.Text;
                    }
                    else
                    {
                        //if (c.CellValue == null)
                        //{
                        //    cellValue = "Column" + colCount;
                        //}
                        //else
                        //{
                            cellValue = c.CellValue.InnerText;
                        //}
                    }
                    colCount++;
                    schema.Columns.Add(cellValue, typeof(string));
                }

            } while (reader.ReadNextSibling());
            return schema;
        }

        //private void ProcessRowsInPart(WorkbookPart workbookPart, OpenXmlReader reader)
        //{
        //    if (reader.ElementType == typeof(Row))
        //    {
        //        reader.ReadFirstChild();
        //        do
        //        {
        //            if (reader.ElementType == typeof(Cell))
        //            {
        //                Cell c = (Cell)reader.LoadCurrentElement();
        //                string cellValue;
        //                if (c.DataType != null && c.DataType == CellValues.SharedString)
        //                {
        //                    SharedStringItem ssi = workbookPart.SharedStringTablePart.
        //                        SharedStringTable.Elements<SharedStringItem>().ElementAt
        //                        (int.Parse(c.CellValue.InnerText));
        //                    cellValue = ssi.Text.Text;
        //                    Console.Write(cellValue + ":");
        //                }
        //                else
        //                {
        //                    cellValue = c.CellValue.InnerText;
        //                    Console.Write(cellValue + ":");
        //                }
        //            }

        //        } while (reader.ReadNextSibling());

        //        Console.WriteLine();
        //        Console.ReadKey();
        //    }
        //}

        public void Dispose()
        {
            if (schema != null)
            {
                schema.Dispose();
            }
        }
    }
}
