using System;
using System.Data;
using Microsoft.Data.SqlClient;
using System.IO;
using OfficeOpenXml;

class ExcelLoaderWithEPPlus
{
    static void Main()
    {
        ExcelPackage.License.SetNonCommercialPersonal("Yersultan"); 

        string folderPath = @"C:\Users\yersu\Desktop\Для интервью\Тех Задание";
        string connectionStringMaster = "Server=localhost;Database=master;Trusted_Connection=True;TrustServerCertificate=True;";

        foreach (var excelFilePath in Directory.GetFiles(folderPath, "*.xlsx"))
        {
            string fileName = Path.GetFileNameWithoutExtension(excelFilePath);
            string dbName = "ExcelDataDB";

            Console.WriteLine($"Обрабатывается файл: {fileName}");

            CreateDatabaseIfNotExists(connectionStringMaster, dbName);

            DataTable dataTable = LoadExcelIntoDataTable(excelFilePath);

            string dbConnectionString = $"Server=localhost;Database={dbName};Trusted_Connection=True;TrustServerCertificate=True;";
            using (var connection = new SqlConnection(dbConnectionString))
            {
                connection.Open();

                if (!TableExists(connection, fileName))
                {
                    CreateTable(connection, fileName, dataTable);
                    Console.WriteLine($"Создана таблица: {fileName}");
                }

                InsertData(connection, fileName, dataTable);
                Console.WriteLine($"Данные загружены в: {fileName}");
            }
        }

        Console.WriteLine("Все Excel-файлы обработаны!");
    }

    static void CreateDatabaseIfNotExists(string connectionStringMaster, string dbName)
    {
        using (var connection = new SqlConnection(connectionStringMaster))
        {
            connection.Open();
            string sql = $"IF DB_ID('{dbName}') IS NULL CREATE DATABASE [{dbName}];";
            using (var command = new SqlCommand(sql, connection))
            {
                command.ExecuteNonQuery();
            }
        }
    }

    static DataTable LoadExcelIntoDataTable(string filePath)
    {
        var dataTable = new DataTable();

        using (var package = new ExcelPackage(new FileInfo(filePath)))
        {
            var worksheet = package.Workbook.Worksheets[0]; // первый лист
            int colCount = worksheet.Dimension.End.Column;
            int rowCount = worksheet.Dimension.End.Row;

            
            for (int col = 1; col <= colCount; col++)
            {
                dataTable.Columns.Add(worksheet.Cells[1, col].Text);
            }

      
            for (int row = 2; row <= rowCount; row++)
            {
                var newRow = dataTable.NewRow();
                for (int col = 1; col <= colCount; col++)
                {
                    newRow[col - 1] = worksheet.Cells[row, col].Text;
                }
                dataTable.Rows.Add(newRow);
            }
        }

        return dataTable;
    }

    static bool TableExists(SqlConnection connection, string tableName)
    {
        string query = @"SELECT COUNT(*) FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = @tableName";
        using (var command = new SqlCommand(query, connection))
        {
            command.Parameters.AddWithValue("@tableName", tableName);
            int count = (int)command.ExecuteScalar();
            return count > 0;
        }
    }

    static void CreateTable(SqlConnection connection, string tableName, DataTable dataTable)
    {
        string sql = $"CREATE TABLE [{tableName}] (";
        foreach (DataColumn column in dataTable.Columns)
        {
            sql += $"[{column.ColumnName}] NVARCHAR(MAX),";
        }
        sql = sql.TrimEnd(',') + ")";
        using (var command = new SqlCommand(sql, connection))
        {
            command.ExecuteNonQuery();
        }
    }

    static void InsertData(SqlConnection connection, string tableName, DataTable dataTable)
    {
        using (var bulkCopy = new SqlBulkCopy(connection))
        {
            bulkCopy.DestinationTableName = tableName;
            bulkCopy.WriteToServer(dataTable);
        }
    }
}