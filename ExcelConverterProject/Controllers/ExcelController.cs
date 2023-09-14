using System.IO;
using System.Net;
using Microsoft.AspNetCore.Mvc;
using NPOI.XSSF.UserModel;
using ExcelDataReader;
using System.Data;
using OfficeOpenXml;
[Route("api/[controller]")]
[ApiController]
public class ExcelController : ControllerBase
{
    [HttpPost]
    public IActionResult Post([FromBody] LinkRequest request)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        using (var client = new WebClient())
        {
            // Download the file data from the provided URL
            var data = client.DownloadData(request.Link);

            using (var reader = ExcelReaderFactory.CreateReader(new MemoryStream(data)))
            {
                var dataSet = reader.AsDataSet();
                var dataTable = dataSet.Tables[0];

                // Convert the DataTable to a list of dynamic objects
                var rows = dataTable.AsEnumerable().Skip(1) // Skip the header row
                    .Select(row => new
                    {
                        Segment = row["Column0"],
                        Country = row["Column1"],
                        Product = row["Column2"],
                        DiscountBand = row["Column3"],
                        UnitsSold = row["Column4"],
                        ManufacturingPrice = row["Column5"],
                        SalePrice = row["Column6"],
                        GrossSales = row["Column7"],
                        Discounts = row["Column8"],
                        Sales = row["Column9"],
                        COGS = row["Column10"],
                        Profit = row["Column11"],
                        Date = row["Column12"],
                        MonthNumber = row["Column13"],
                        MonthName = row["Column14"],
                        Year = row["Column15"]
                    })
                    .ToList();

                // Return the list of rows as a JSON response
                return new JsonResult(rows);
            }
        }
    }
}

public class LinkRequest
{
    public string Link { get; set; }
}