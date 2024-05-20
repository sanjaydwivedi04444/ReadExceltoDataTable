using Azure.Storage.Blobs;
using Azure.Storage.Blobs.Models;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadExcel
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string blobName= Environment.GetEnvironmentVariable("BlobName");
            foreach (DataRow dr in GetDataTable(blobName).Rows)
            {
                Console.WriteLine(dr.ItemArray[0] + "," + dr.ItemArray[1] + "," + dr.ItemArray[2] + "," + dr.ItemArray[3] + "," + dr.ItemArray[4] + "," + dr.ItemArray[5] + "," + dr.ItemArray[6]);
            }
            Console.WriteLine(solution(""));
            Console.ReadLine();
        }
        public static int solution(string S)
        {
            // Implement your solution here
            // Convert the binary string to an integer
            int V = Convert.ToInt32(S, 2);
            int operationsCount = 0;

            // Perform operations until V becomes 0
            while (V > 0)
            {
                if (V % 2 == 0)
                {
                    V /= 2;
                }
                else
                {
                    V -= 1;
                }
                operationsCount++;
            }

            // Return the number of operations
            return operationsCount;
        }
        public static int CellReferenceToIndex(Cell cell)
        {
            int index = -1;
            string reference = cell.CellReference.ToString().ToUpper();
            foreach (char ch in reference)
            {
                if (Char.IsLetter(ch))
                {
                    int value = (int)ch - (int)'A';
                    index = (index + 1) * 26 + value;
                }
                else
                    return index;
            }
            return index;
        }
        public static DataTable GetDataTable(string blobName)
        {
            //Read Excel From Storage
            string connectionString = Environment.GetEnvironmentVariable("ConnectionString");
            string sourceContainerName =Environment.GetEnvironmentVariable("SourceContainerName");
            var input = new List<BlobOutput>();
            var dt = new DataTable();
            foreach (BlobOutput item in ReadExcel(connectionString, sourceContainerName, blobName).Result)
            {
                using (SpreadsheetDocument spreadSheetDocument = SpreadsheetDocument.Open(item.BlobContent, false))
                {
                    WorkbookPart workbookPart = spreadSheetDocument.WorkbookPart;
                    IEnumerable<Sheet> sheets = spreadSheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
                    string relationshipId = sheets.First().Id.Value;
                    WorksheetPart worksheetPart = (WorksheetPart)spreadSheetDocument.WorkbookPart.GetPartById(relationshipId);
                    Worksheet workSheet = worksheetPart.Worksheet;
                    SheetData sheetData = workSheet.GetFirstChild<SheetData>();
                    IEnumerable<Row> rows = sheetData.Descendants<Row>();

                    foreach (Cell cell in rows.ElementAt(0))
                    {
                        dt.Columns.Add(GetCellValue(spreadSheetDocument, cell));
                    }

                    foreach (Row row in rows) //this will also include your header row...
                    {
                        DataRow tempRow = dt.NewRow();

                        for (int i = 0; i < row.Descendants<Cell>().Count(); i++)
                        {
                            Cell cell = row.Descendants<Cell>().ElementAt(i);
                            int index = CellReferenceToIndex(cell);
                            tempRow[index] = GetCellValue(spreadSheetDocument, cell);
                        }

                        dt.Rows.Add(tempRow);
                    }
                }
                dt.Rows.RemoveAt(0); // Delete extra row ...so i'm taking it out here.
            }
            //End Excel
            return dt;
        }
        public static string GetCellValue(SpreadsheetDocument document, Cell cell)
        {
            SharedStringTablePart stringTablePart = document.WorkbookPart.SharedStringTablePart;
            string value = cell.CellValue.InnerXml;

            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                return stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;
            }
            else
            {
                return value;
            }
        }
        public static async Task<List<BlobOutput>> ReadExcel(string connectionString, string containerName, string blobName)
        {
            BlobServiceClient _blobServiceClient = new BlobServiceClient(connectionString);
            var downloadedData = new List<BlobOutput>();
            try
            {
                // Create service and container client for blob  
                BlobContainerClient blobContainerClient = _blobServiceClient.GetBlobContainerClient(containerName);

                // List all blobs in the container  
                //await foreach (BlobItem item in blobContainerClient.GetBlobsAsync())
                //{
                // Download the blob's contents and save it to a file  
                BlobClient blobClient = blobContainerClient.GetBlobClient(blobName);
                BlobDownloadInfo downloadedInfo = await blobClient.DownloadAsync();
                downloadedData.Add(new BlobOutput { BlobName = blobName, BlobContent = downloadedInfo.Content });
                //}
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return downloadedData;
        }

    }
}
