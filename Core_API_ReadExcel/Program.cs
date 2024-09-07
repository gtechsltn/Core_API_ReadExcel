using ExcelDataReader;
using Microsoft.Extensions.Hosting;
using System.Data;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
// Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();

var app = builder.Build();

// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

app.UseHttpsRedirection();

app.MapGet("/readexcel/{file}", async (HttpContext context, IWebHostEnvironment env, string file) =>
{
    // 1. Get the directory path
    string dirPath = Path.Combine(env.ContentRootPath, "Files");
    // 2. Read Files
    var files = Directory.GetFiles(dirPath);

    // 3. Get The FileName from the received 'file'
    var fileName = files.FirstOrDefault(f => f.Contains(file));

    // 4. If file is not found return NotFound
    if (fileName == null)
    {
        return Results.NotFound($"File {file}.xlsx is not available");
    }


    // 5. Register the Provider

    System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
    // 6. Define the DataSet
    DataSet dataSet;
    // 7. Read the Excel File on the server
    using (var stream =  File.OpenRead(fileName))
    {
        // Create a reader for the Excel file
        using (var reader = ExcelReaderFactory.CreateReader(stream))
        {
            var conf = new ExcelDataSetConfiguration
            {
                ConfigureDataTable = _ => new ExcelDataTableConfiguration
                {
                    // 8. Use the first row as the header
                    UseHeaderRow = true
                }
            };
            // 9. Convert theworksheet to DataSet
            dataSet = reader.AsDataSet(conf);
        }
    }

    var dataResult = new List<Dictionary<string, object>>();
   
    // 10. Loop through the tables
    foreach (DataTable table in dataSet.Tables)
    {
        // 11. Loop through the rows
        foreach (DataRow row in table.Rows)
        {
            var rowDataDict = new Dictionary<string, object>();
            foreach (DataColumn col in table.Columns)
            {
                //12 Add the column name and value to the dictionary
                rowDataDict[col.ColumnName] = row[col];
            }
            // 13 Add the dictionary to the JSON Result
            dataResult.Add(rowDataDict);
        }
    }

    return Results.Ok(dataResult);
});

app.Run();
