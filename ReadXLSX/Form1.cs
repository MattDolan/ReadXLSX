using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ReadXLSX
{
    public partial class Form1 : Form
    {
        string strXLSXfile = "";
        int iStartRow = 0;
        int iEndRow = 0;
        string strColumns = "";
        string[] strColumnToOutput = null;
        string strOutputFile = "";
        bool bComplete = false;

        public Form1()
        {
            InitializeComponent();

            /* Get the arguments */

            string[] args = Environment.GetCommandLineArgs();
            if (args.Length != 2)
            {
                MessageBox.Show("No command line argument provided. " + args[0]);
                return;
            }
            
            /* Break them up since we know the expected format */

            string[] argsSplit = args[1].Split('|');
            if (argsSplit.Length != 4)
            {
                MessageBox.Show("No command line argument provided.");
                return;
            }

            /* Get the input XLSX file and make sure it is good */

            strXLSXfile = argsSplit[0].Trim();
            if (strXLSXfile == "")
            {
                MessageBox.Show("No input XLSX file provided.");
                return;
            }
            if (File.Exists(strXLSXfile) == false)
            {
                MessageBox.Show("Input file provided could not be found." + Environment.NewLine + strXLSXfile);
                return;
            }

            /* Get the start and end row */

            string strRows = argsSplit[1];

            int iIndex = strRows.IndexOf('-');
            if (iIndex == -1)
            {
                MessageBox.Show("Start and End row argument is not formatted correctly. Should be 2-155");
                return;
            }

            iStartRow = Convert.ToInt16(strRows.Substring(0, iIndex));
            iIndex = iIndex + 1;

            if (iIndex == strRows.Length)
            {
                iEndRow = 0;
            }
            else
            {
                iEndRow = Convert.ToInt16(strRows.Substring(iIndex, strRows.Length - iIndex));
            }

            /* Get the columns to be output */

            strColumns = argsSplit[2].Trim();
            strColumnToOutput = strColumns.Split(',');
            if (strColumnToOutput.Length == 0)
            {
                MessageBox.Show("No column were provided for outputing.");
                return;
            }

            /* Get the output file and make sure it is good */
            
            strOutputFile = argsSplit[3].Trim();
            if (strOutputFile == "")
            {
                MessageBox.Show("No output XLSX file provided.");
                return;
            }
            if (File.Exists(strOutputFile) == true)
            {
                MessageBox.Show("Output file provided already exists.  It will be overwritten." + Environment.NewLine + strOutputFile);
                File.Delete(strOutputFile);
            }

            OutputXLSX();

        }

        private void OutputXLSX()
        {
            string strOutputData = "";
            FileInfo XLSXfile = new FileInfo(strXLSXfile);
            
            /* Open and read the XlSX file. */

            try
            {
                using (ExcelPackage package = new ExcelPackage(XLSXfile))
                {

                    /* Get the work book in the file */

                    ExcelWorkbook workBook = package.Workbook;
                    if (workBook != null)
                    {
                        if (workBook.Worksheets.Count > 0)
                        {

                            /* Get the first worksheet */

                            //ExcelWorksheet Worksheet = workBook.Worksheets.First();
                            var worksheet = package.Workbook.Worksheets[1];

                            if (iEndRow == 0)
                            {
                                /* Find the "real" last used row. */

                                var rowRun = worksheet.Dimension.End.Row;
                                while (rowRun >= 1)
                                {
                                    var range = worksheet.Cells[rowRun, 1, rowRun, worksheet.Dimension.End.Column];
                                    if (range.Any(c => !string.IsNullOrEmpty(c.Text)))
                                    {
                                        break;
                                    }
                                    rowRun--;
                                }
                                iEndRow = rowRun;
                            }

                            /* Create the output file */

                            using (StreamWriter writer = File.AppendText(strOutputFile))
                            {

                                /* Loop through the worksheet and output the values we need. */
                                /* Go from the start row to the end row */

                                for (int row = iStartRow; row <= iEndRow; row++)
                                {
                                    /* Build a string with the values from each column */

                                    strOutputData = "";
                                    foreach (string col in strColumnToOutput)
                                    {
                                        string strValue = worksheet.Cells[col + row.ToString()].Value == null ? string.Empty : worksheet.Cells[col + row.ToString()].Value.ToString();
                                        strOutputData = strOutputData + strValue + "|";
                                    }
                                    writer.WriteLine(strOutputData);
                                }
                                writer.Dispose();
                            }

                        }
                    }
                }
            }
            catch (IOException ex)
            {
                MessageBox.Show("Error opening spreadsheet. Is it already open? Close it and try again." + Environment.NewLine + ex.Message);
                return;
            }
            bComplete = true;
        }

        private void Exit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            if (bComplete == true)
            {
                Application.Exit();
            }
        }
    }
}
