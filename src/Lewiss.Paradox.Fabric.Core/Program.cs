namespace Lewiss.Paradox.Fabric.Core;

using System.Data.Odbc;
using System.Net.Http.Json;
using System.Text.Json;
using ClosedXML;
using ClosedXML.Excel;

// dotnet publish -c Release --runtime win-x86 --self-contained true -p:PublishSingleFile=true

class Program
{
    async static Task Main(string[] args)
    {
        try
        {
            if (!ArchitectureCheck()) System.Environment.Exit(1);

            var connectionString = @"Driver={Microsoft Paradox Driver (*.db )};DriverID=538;Fil=Paradox 4.X;DefaultDir=K:\Data;Dbq=K:\Data\CUTS\LOOKUP;CollatingSequence=International;";
            var tableName = "MPFabric";

            var headerList = ReadTableHeader(connectionString, tableName);

            var fabricData = GetValues(connectionString, tableName);

            var filePath = @"K:\Data\Process\Development\fabric\fabric.xlsx";
            var fileBase64 = GetExcelFileAsBase64(filePath, headerList, fabricData);

            await WriteWorkbookToSharePoint(fileBase64);

        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed: {ex.Message}");
        }



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

    private static List<string> ReadTableHeader(string connectionString, string tableName)
    {
        using (OdbcConnection connection = new OdbcConnection(connectionString))
        using (OdbcCommand command = new OdbcCommand($"SELECT * FROM {tableName}", connection))
        {
            try
            {

                System.Console.WriteLine("Reading table header");

                connection.Open();

                List<string> output = [];


                using (OdbcDataReader reader = command.ExecuteReader())
                {

                    if (!reader.Read())
                    {
                        System.Console.WriteLine("Failed to read first row");
                    }

                    for (var i = 0; i < reader.FieldCount; i++)
                    {
                        output.Add(reader.GetName(i));
                    }



                }

                return output;

            }
            catch (OdbcException ex)
            {
                System.Console.WriteLine($"Odbc: ${ex.Message}");
                throw;

            }
            catch (Exception ex)
            {
                System.Console.WriteLine($"General: {ex.Message}");
                throw;


            }

        }


    }

    private static List<List<string>> GetValues(string connectionString, string tableName)
    {
        List<List<string>> output = [];

        var query = $"SELECT * FROM {tableName}";

        using (OdbcConnection connection = new OdbcConnection(connectionString))
        using (OdbcCommand command = new OdbcCommand(query, connection))
        {
            try
            {

                System.Console.WriteLine("Reading fabric values");

                connection.Open();



                using (OdbcDataReader reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {

                        List<string> newEntry = [];

                        for (var i = 0; i < reader.FieldCount; i++)
                        {
                            if (reader.IsDBNull(i))
                            {
                                newEntry.Add(string.Empty);
                            }
                            else
                            {
                                newEntry.Add(reader.GetValue(i).ToString() ?? string.Empty);
                            }
                        }

                        output.Add(newEntry);

                    }
                }

            }
            catch (OdbcException ex)
            {
                System.Console.WriteLine($"Odbc: ${ex.Message}");
                throw;

            }
            catch (Exception ex)
            {
                System.Console.WriteLine($"General: {ex.Message}");
                throw;


            }




        }

        return output;

    }


    private static string GetExcelFileAsBase64(string filePath, List<string> header, List<List<string>> data)
    {
        using (var workbook = new XLWorkbook())
        {

            try
            {
                System.Console.WriteLine("Saving workbook");
                var worksheet = workbook.Worksheets.Add("A");

                WriteWorksheetHeader(worksheet, header);
                WriteWorksheetData(worksheet, data);



                using (var memoryStream = new MemoryStream())
                {
                    workbook.SaveAs(memoryStream);
                    return Convert.ToBase64String(memoryStream.ToArray());
                }

            }
            catch (Exception ex)
            {
                System.Console.WriteLine($"General: {ex.Message}");
                throw;

            }

        }


    }


    private static void WriteWorksheetHeader(IXLWorksheet worksheet, List<string> header)
    {
        var rowStart = 1;
        var columnStart = 1;

        for (var i = 0; i < header.Count; i++)
        {
            worksheet.Cell(rowStart, i + columnStart).Value = header[i];
        }
    }

    private static void WriteWorksheetData(IXLWorksheet worksheet, List<List<string>> data)
    {
        var rowStart = 2;
        var columnStart = 1;

        for (var i = 0; i < data.Count; i++)
        {
            var row = rowStart + i;

            for (var j = 0; j < data[i].Count; j++)
            {
                var column = columnStart + j;
                worksheet.Cell(row, column).Value = data[i][j];
            }
        }
    }


    private async static Task WriteWorkbookToSharePoint(string fileBase64)
    {
        var sharepointFolder = "01VFVMOAB5NYKNEK4WOJELK3E4XZK5ZKFJ";
        var fileName = "FABRICS.xlsx";

        using (HttpClientHandler handler = new HttpClientHandler() { UseCookies = false })
        using (HttpClient client = new HttpClient(handler))
        {
            try
            {
                System.Console.WriteLine("Saving to Sharepoint");
                var payload = new
                {
                    action = "uploadAndNotify",
                    filename = fileName,
                    fileBase64 = fileBase64,
                    folderId = sharepointFolder,
                    custName = "",
                    custAddress = "",
                    custMeasurer = "",
                    custDate = "",
                    custSalesConsultant = "",
                    consultantEmail = "",
                    photos = ""
                };

                var url = "https://lewiss-measure-pro.netlify.app/.netlify/functions/graph";


                var cookie = "a8f9590f-d26e-407d-bce1-39705a940a4d=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJoYXNoIjoiWHg1NmVkMXdCQmJ2MHUtR0NxVUtSQWc0ZlNGZXVVZ0RETV9nRDBsSWFhZz0iLCJzaXRlX2lkIjoiYThmOTU5MGYtZDI2ZS00MDdkLWJjZTEtMzk3MDVhOTQwYTRkIn0.EFQyaYl9keL8wbuLGUf2mXG6EBQsu2fUCyQ9o2lrT8Y";
                var request = new HttpRequestMessage(HttpMethod.Post, url);

                request.Headers.TryAddWithoutValidation("Cookie", cookie);
                request.Content = JsonContent.Create(payload);


                HttpResponseMessage response = await client.SendAsync(request);

                response.EnsureSuccessStatusCode();

                string responseBody = await response.Content.ReadAsStringAsync();
                JSONOutput(responseBody);
            }
            catch (HttpRequestException ex)
            {
                System.Console.WriteLine($"Http: {ex.Message}");
                throw;


            }
            catch (Exception ex)
            {
                System.Console.WriteLine($"General: {ex.Message}");
                throw;

            }
        }


    }

    private static void JSONOutput(string json)
    {
        using (var jsonDoc = JsonDocument.Parse(json))
        {
            var options = new JsonSerializerOptions() { WriteIndented = true };
            var output = JsonSerializer.Serialize(jsonDoc, options);
            System.Console.WriteLine(output);
        }
    }
}
