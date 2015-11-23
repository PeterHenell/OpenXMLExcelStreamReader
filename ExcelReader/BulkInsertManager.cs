using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReaderTest
{
    public class BulkInsertManager
    {
        public void BulkInsertTo(DataTable schema, string tableName, IDataReader dataReader, string connectionString)
        {
            using (var bulkCopy = new SqlBulkCopy(connectionString))
            {
                bulkCopy.BulkCopyTimeout = 9000000;
                bulkCopy.BatchSize = 1000;
                bulkCopy.DestinationTableName = string.Format("[{0}]", tableName);
                for (int ordinal = 0; ordinal < schema.Columns.Count; ordinal++)
                {
                    var mapping = new SqlBulkCopyColumnMapping
                    {
                        DestinationOrdinal = ordinal,
                        SourceOrdinal = ordinal
                    };
                    bulkCopy.ColumnMappings.Add(mapping);
                }

                bulkCopy.WriteToServer(dataReader);
            }
        }
    }
}
