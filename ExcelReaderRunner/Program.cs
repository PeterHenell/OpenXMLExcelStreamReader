using ExcelReaderTest;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReaderRunner
{
    class Program
    {
        static void Main(string[] args)
        {
            string fileName = @"C:\Users\pehe\Downloads\REPORT 73.xlsx";

            var ddlManager = new DDLManager();
            var bulkdInsertManager = new BulkInsertManager();
            String connString = GetConnectionString();
            Console.WriteLine("Start time" + DateTime.Now);

            ExcelStreamReader.Execute(fileName, reader =>

                reader.ForEachSheet(sheet =>
                {
                    var schema = sheet.GetSchema();

                    ddlManager.CreateTable(schema.TableName, schema, connString);
                    bulkdInsertManager.BulkInsertTo(schema, schema.TableName, sheet, connString);
                })
            );

            Console.WriteLine("End time" + DateTime.Now);
        }
        private static string GetConnectionString()
        {
            var builder = new SqlConnectionStringBuilder();
            builder.DataSource = "localhost";
            builder.IntegratedSecurity = true;
            builder.InitialCatalog = "testdb";
            return builder.ToString();
        }
    }
}
