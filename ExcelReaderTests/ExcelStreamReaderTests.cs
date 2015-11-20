using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReaderTest
{
    [TestFixture]
    public class ExcelStreamReaderTests
    {
        string fileName = @"C:\Users\pehe\Downloads\REPORT 73.xlsx";

        [Test]
        public void ShouldReadHeaders()
        {
            ExcelStreamReader.Execute(fileName, reader =>

                reader.ForEachSheet(sheet =>
                {
                    sheet.ForEachRow((schema, row) =>
                    {
                        Console.WriteLine("New Row");
                        for (int i = 0; i < schema.Columns.Count; i++)
                        {
                            Console.WriteLine(row[i]);
                        }
                    });
                })
            );
        }
    }
}
