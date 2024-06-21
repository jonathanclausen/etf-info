using System;
using ExcelDataReader;
using System.Data;
using System.IO;
using System.Net;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Extensions.Logging;

namespace ExcelExtractor
{
  public class ExcelExtractor
  {
    [FunctionName("ExcelExtractor")]
    public void Run([TimerTrigger("0 0 * * * *")]TimerInfo myTimer, ILogger log)
    {
      System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

      log.LogInformation($"C# Timer trigger function executed at: {DateTime.Now}");
      string url = "https://skat.dk/media/hdylxyjl/abis-listen-27052024.xlsx";
      // Create a WebClient to download the file
      using (WebClient client = new WebClient())
      {
        // Download the data as a byte array
        byte[] data = client.DownloadData(url);

        // Read the Excel file from the byte array
        using (MemoryStream stream = new MemoryStream(data))
        {
          using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
          {
            // Use the AsDataSet extension method to get a DataSet
            DataSet result = reader.AsDataSet();

            // Assuming you want the first table in the workbook
            DataTable dataTable = result.Tables[0];

            // Display the first few rows of the DataTable
            foreach (DataRow row in dataTable.Rows)
            {
              foreach (var item in row.ItemArray)
              {
                log.LogInformation($"{item} ");
              }
             
            }
          }
        }
      }
    }
  }
}

