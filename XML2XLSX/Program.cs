using Microsoft.Office.Interop.Excel;
using System;
using System.Data;
using System.IO;
using System.Text;

namespace XML2XLSX
{
    class Program
    {
        static void Main(string[] args)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            
            DataSet ds = new DataSet();

            Console.WriteLine("XML file name with full path:");
            string filepath = Console.ReadLine();
            //Convert the XML into Dataset
            ds.ReadXml(filepath);

            Console.WriteLine("Start!");

            try
            {
                Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
                Workbook xlWorkBook = ExcelApp.Workbooks.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet);
                foreach (System.Data.DataTable dt in ds.Tables)
                {
                    System.Data.DataTable dtDataTable1 = dt;


                    for (int i = 1; i > 0; i--)
                    {

                        Sheets xlSheets = null;
                        Worksheet xlWorksheet = null;
                        //Create Excel sheet
                        xlSheets = ExcelApp.Sheets;
                        xlWorksheet = (Worksheet)xlSheets.Add(xlSheets[1], Type.Missing, Type.Missing, Type.Missing);
                        xlWorksheet.Name = dt.TableName;
                        for (int j = 1; j < dtDataTable1.Columns.Count + 1; j++)
                        {
                            ExcelApp.Cells[i, j] = dtDataTable1.Columns[j - 1].ColumnName;
                            //ExcelApp.Cells[1, j].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);
                            //ExcelApp.Cells[i, j].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.WhiteSmoke);
                        }
                        // for the data of the excel
                        for (int k = 0; k < dtDataTable1.Rows.Count; k++)
                        {
                            for (int l = 0; l < dtDataTable1.Columns.Count; l++)
                            {
                                ExcelApp.Cells[k + 2, l + 1] = dtDataTable1.Rows[k].ItemArray[l].ToString();
                            }
                        }
                        ExcelApp.Columns.AutoFit();
                    }
                    //((Worksheet)ExcelApp.ActiveWorkbook.Sheets[ExcelApp.ActiveWorkbook.Sheets.Count]).Delete();
                    Console.WriteLine(dt.TableName + " created.");
                }
                ExcelApp.Visible = true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }


            /*
            foreach (DataTable dt in ds.Tables)
            {
                // Create an Excel object
                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();

                //Microsoft.Office.Interop.Excel.Workbook workbook = excel.Workbooks.Open(Filename: str + dt.TableName + ".xlsx");
                //Create worksheet object
                Microsoft.Office.Interop.Excel.Worksheet worksheet;
                worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Application.Worksheets.Add();

                //((Microsoft.Office.Interop.Excel._Worksheet)worksheet).Activate();
                //Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.NewWindow();

                worksheet.Name = dt.TableName;
                // Column Headings
                int iColumn = 0;

                foreach (DataColumn c in dt.Columns)
                {
                    iColumn++;
                    excel.Cells[1, iColumn] = c.ColumnName;
                }

                // Row Data
                int iRow = worksheet.UsedRange.Rows.Count - 1;

                foreach (DataRow dr in dt.Rows)
                {
                    iRow++;

                    // Row's Cell Data
                    iColumn = 0;
                    foreach (DataColumn c in dt.Columns)
                    {
                        iColumn++;
                        excel.Cells[iRow + 1, iColumn] = dr[c.ColumnName];
                    }
                }

                //((Microsoft.Office.Interop.Excel._Worksheet)worksheet).Activate();
                //((Microsoft.Office.Interop.Excel._Worksheet)worksheet).SaveAs(dt.TableName);
                worksheet = null;

                //Save the workbook
                workbook.Save();

                //Close the Workbook
                workbook.Close();

                // Finally Quit the Application
                ((Microsoft.Office.Interop.Excel._Application)excel).Quit();
            }
            */
            Console.WriteLine("End!");

            Console.ReadLine();
        }
    }
}
