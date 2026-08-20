namespace Lewiss.Paradox.Fabric.Core;

using System.Data;
using System.Data.Odbc;
using Microsoft.VisualBasic;

class Program
{
    static void Main(string[] args)
    {
        if (!ArchitectureCheck()) System.Environment.Exit(1);

        var connectionString = @"Driver={Microsoft Paradox Driver (*.db )};DriverID=538;Fil=Paradox 4.X;DefaultDir=K:\Data;Dbq=K:\Data\CUTS\LOOKUP;CollatingSequence=International;";
        var tableName = "MPFabric";
        var query = $"SELECT * FROM {tableName}";

        ExecuteQuery(connectionString, query);

        System.Console.ReadLine();

    }

    private static bool ArchitectureCheck()
    {
        if (!Environment.Is64BitProcess)
        {

            return true;
        }

        System.Console.WriteLine("Error: Process must be running in 32 bit.");

        return true;
    }

    private static void ExecuteQuery(string connectionString, string query)
    {


        using (OdbcConnection connection = new OdbcConnection(connectionString))
        using (OdbcCommand command = new OdbcCommand(query, connection))
        {
            try
            {

                connection.Open();
                var line = 0;

                using (OdbcDataReader reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {


                        for (var i = 0; i < reader.FieldCount; i++)
                        {
                            string columnName = reader.GetName(i);
                            System.Console.WriteLine($"{columnName} : {reader[i]}");
                        }

                        line += 1;

                        if (line > 1) break;
                    }
                }

            }
            catch (OdbcException ex)
            {
                System.Console.WriteLine($"Odbc: ${ex.Message}");
            }
            catch (Exception ex)
            {
                System.Console.WriteLine($"General: {ex.Message}");
            }




        }


    }


}
