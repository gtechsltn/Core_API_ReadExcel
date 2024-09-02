using ExcelDataReader;
using Microsoft.Extensions.Hosting;

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

    // 5. Read the Excel File
    System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
    using var stream = new FileStream(fileName, FileMode.Open);
    IExcelDataReader reader;
    // 6. Reading from a OpenXml Excel file (2007 format; *.xlsx)
    reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
    // 7. DataSet - The result of each spreadsheet will be created in the result.Tables
    var result = reader.AsDataSet();
    // 8. Close the data reader
    reader.Close();
    // 9. Get the first table
    var table = result.Tables[0];

    var headerRow = table.Rows[0];

    // Read all columns for Header row

    context.Response.ContentType = "application/json";
    await context.Response.WriteAsJsonAsync(table);


    /*
      var files = Directory.GetFiles(dirPath);
    var fileNames = new List<string>();

    foreach (var filePath in files)
    {
        fileNames.Add(Path.GetFileName(filePath));
    }

    context.Response.ContentType = "application/json";
    await context.Response.WriteAsJsonAsync(fileNames);
     */





     return Results.Ok();
});

app.Run();
