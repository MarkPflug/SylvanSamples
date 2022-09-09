using Sylvan.Data;
using Sylvan.Data.Excel;
using System.Data.Common;
using System.Data.SqlClient;

// create a schema that maps the excel headers to a different name
var schema = new MySchema("a>Account,c>Processing Date:date,e>Value:int?");

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

// bulk copy it into the table

// SQL TABLE DEFINITION:
//create table MyData (
//Account varchar(100),
//[processing date] datetime2,
//Value int
//)

var conn = new SqlConnection("Data Source=.;Initial Catalog=mydb;Integrated Security=true;");
conn.Open();
var bc = new SqlBulkCopy(conn);
bc.DestinationTableName = "MyData";
bc.WriteToServer(dataToLoad);

class MySchema : IExcelSchemaProvider
{
    Schema schema;

    public MySchema(string schemaSpec)
    {
        this.schema = Schema.Parse(schemaSpec);
    }

    public DbColumn? GetColumn(string sheetName, string? name, int ordinal)
    {
        foreach (var col in schema)
        {
            if (string.Equals(col.BaseColumnName, name, StringComparison.OrdinalIgnoreCase))
            {
                return col;
            }
        }
        return null;
    }

    public bool HasHeaders(string sheetName)
    {
        return true;
    }
}
