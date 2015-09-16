using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace WebApplication
{
    public partial class Default : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!this.IsPostBack)
            {
                string constr = ConfigurationManager.ConnectionStrings["NorthwindMDF"].ConnectionString;
                using (SqlConnection con = new SqlConnection(constr))
                {
                    using (SqlCommand cmd = new SqlCommand("SELECT * FROM [Alphabetical list of products]"))
                    {
                        using (SqlDataAdapter sda = new SqlDataAdapter())
                        {
                            cmd.Connection = con;
                            sda.SelectCommand = cmd;
                            using (DataTable dt = new DataTable())
                            {
                                sda.Fill(dt);
                                GridView.DataSource = dt;
                                GridView.DataBind();
                            }
                        }
                    }
                }
            }
        }
        public DataTable GetProducts()
        {
            using (var conn = new SqlConnection(ConfigurationManager.ConnectionStrings["NorthwindMDF"].ConnectionString))
            using (var cmd = new SqlCommand("SELECT * FROM [Alphabetical list of products]", conn))
            using (var adapter = new SqlDataAdapter(cmd))
            {
                var products = new DataTable();
                adapter.Fill(products);
                return products;
            }
        }
        protected void btnExportExcel_Click(object sender, EventArgs e)
        {
            var products = GetProducts();
            ExcelPackage excel = new ExcelPackage();
            var workSheet = excel.Workbook.Worksheets.Add("Products");
            var totalCols = products.Columns.Count;
            var totalRows = products.Rows.Count;

            for (var col = 1; col <= totalCols; col++)
            {
                workSheet.Cells[1, col].Value = products.Columns[col - 1].ColumnName;
                
                //Set column width
                workSheet.Column(col).Width = 30;
            }
            for (var row = 1; row <= totalRows; row++)
            {
                for (var col = 0; col < totalCols; col++)
                {
                    object value = products.Rows[row - 1][col];
                    workSheet.Cells[row + 1, col + 1].Value = value;
                    //If the value is DateTime, format to DateTime
                    if (value.GetType() == typeof(DateTime))
                    {
                        workSheet.Cells[row + 1, col + 1].Style.Numberformat.Format = "MM/dd/yyyy hh:mm:ss AM/PM";
                    }
                }
            }
            using (var memoryStream = new MemoryStream())
            {
                
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;  filename=products.xlsx");
                excel.SaveAs(memoryStream);
                memoryStream.WriteTo(Response.OutputStream);
                Response.Flush();
                Response.End();
            }
            
        }
    }
}