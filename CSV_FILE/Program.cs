using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualBasic.FileIO;
using Bytescout.Spreadsheet;
using System.Configuration;

namespace CSV_FILE
{
    static class Program
    {
        #region IMPORT CSV FILE TO THE DATABASE
        static void Main(string[] args)
        {
            try
            {
                var connectionString = ConfigurationManager.ConnectionStrings["csvFileContext"].ConnectionString;
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    #region DATABASE QUERY 
                    // Drop test database if exists
                    ExecuteQueryWithoutResult(connection, "IF DB_ID('CSV_FILE_TEST') IS NOT NULL DROP DATABASE CSV_FILE");
                    // Create empty database
                    ExecuteQueryWithoutResult(connection, "CREATE DATABASE CSV_FILE_TEST");
                    // Switch to created database
                    ExecuteQueryWithoutResult(connection, "USE CSV_FILE");
                    // Create a table for CSV data (Depends on the name of your column in CSV File)
                    ExecuteQueryWithoutResult(connection,
                    "CREATE TABLE [dbo].[cr_expences](iCtr INT,cCRCode CHAR(50),cExpenses CHAR(50),iCost DECIMAL(9, 2))");
                    #endregion

                    using (Spreadsheet document = new Spreadsheet())
                    {
                        document.LoadFromFile(ConfigurationManager.AppSettings["path"],",");
                        Worksheet worksheet = document.Workbook.Worksheets[0];

                        for (int row = 1; row <= worksheet.UsedRangeRowMax; row++)
                        {
                            //Depends on how many column of you CSV File
                            String insertCommand = string.Format("INSERT [crCashReceipt_Expenses] VALUES('{0}','{1}','{2}','{3}')", worksheet.Cell(row, 0).Value, worksheet.Cell(row, 1).Value, worksheet.Cell(row, 2).Value, worksheet.Cell(row, 3).Value);
                            ExecuteQueryWithoutResult(connection, insertCommand);
                            Console.WriteLine($"Uploaded data:{row}");
                        }

                    }
                    Console.WriteLine();
                    Console.WriteLine("Successfully uploaded.");
                    Console.ReadKey();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
                Console.ReadKey();
            }
        }
        static void ExecuteQueryWithoutResult(SqlConnection connection, string query)
        {
            using(SqlCommand command =new SqlCommand(query, connection))
            {
                command.ExecuteNonQuery();
            }
        }
        #endregion
    }
}
