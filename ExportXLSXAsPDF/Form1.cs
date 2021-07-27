using ClosedXML.Excel;
using System;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

namespace ExportXLSXAsPDF
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public string ExportPath
        {
            get
            {
                var path = Path.Combine(Application.StartupPath, "Export");
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }
                return path;
            }
        }

        public string ExcelFileName
        {
            get
            {
                return "Sample.xlsx";
            }
        }
        public string PdfFileName
        {
            get
            {
                return "Sample.pdf";
            }
        }

        public DataTable getData()
        {
            //Creating DataTable  
            DataTable dt = new DataTable();
            //Setiing Table Name  
            dt.TableName = "EmployeeData";
            //Add Columns  
            dt.Columns.Add("ID", typeof(int));
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("City", typeof(string));
            //Add Rows in DataTable  
            dt.Rows.Add(1, "Anoop Kumar Sharma", "Delhi");
            dt.Rows.Add(2, "Andrew", "U.P.");
            dt.AcceptChanges();
            return dt;
        }

        public void WriteDataToExcel()
        {
            DataTable dt = getData();

            string fileName = Path.Combine(ExportPath, ExcelFileName);
            using (XLWorkbook wb = new XLWorkbook())
            {
                //Add DataTable in worksheet  
                wb.Worksheets.Add(dt);
                wb.SaveAs(fileName);
            }
        }

        public void ExportAsPdf()
        {
            string fileNameExcel = Path.Combine(ExportPath, ExcelFileName);
            string fileNamePdf = Path.Combine(ExportPath, PdfFileName);
            if (File.Exists(fileNameExcel))
            {
                //PDF
                using (Spire.Xls.Workbook workbook = new Spire.Xls.Workbook())
                {
                    workbook.LoadFromFile(fileNameExcel);


                    workbook.SaveToFile(fileNamePdf, Spire.Xls.FileFormat.PDF);
                }


                if (File.Exists(fileNamePdf))
                {
                    Process.Start(fileNamePdf);
                }
            }
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            btnRun.Enabled = false;
            Application.DoEvents();

            try
            {
                //Export Excel
                WriteDataToExcel();

                //Generate Pdf
                ExportAsPdf();
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }

            btnRun.Enabled = true;
            Application.DoEvents();

        }
    }
}
