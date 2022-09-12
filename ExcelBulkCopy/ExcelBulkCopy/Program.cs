using Sylvan.Data;
using Sylvan.Data.Excel;
using System.Data.SqlClient;

var conn = new SqlConnection("Data Source=.;Initial Catalog=mydb;Integrated Security=true;");
conn.Open();

// create a schema that maps the excel headers to a different name
var schema = new ExcelSchema(true, Schema.Parse("a>Account,c>Processing Date:date,e>Value:int?"));

var opts = new ExcelDataReaderOptions
{
    Schema = schema
};

var edr = ExcelDataReader.Create("data.xlsx", opts);

// locate desired sheet
while (edr.WorksheetName != "MyData")
{
    if (!edr.NextResult())
    {
        throw new Exception("Couldn't find sheet");
    }
}

// select the three columns to load
var dataToLoad = edr.Select("Account", "Processing Date", "Value");

var bc = new SqlBulkCopy(conn);
bc.DestinationTableName = "MyData";
bc.WriteToServer(dataToLoad);
