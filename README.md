# OpenXMLExcelStreamReader
Read like a stream from open xml Excel files. Avoid loading the entire file into memory.


Usage:

        ExcelStreamReader.Execute(@"c:\temp\BigExcelFile.xlsx", reader =>

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
