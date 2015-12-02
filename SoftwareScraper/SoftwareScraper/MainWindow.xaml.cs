using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;

namespace SoftwareScraper
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, EventArgs e)
        {
            // Create DataSet
            DataSet ds = new DataSet("SoftwareScraper");

            // Create DataTable with columns
            System.Data.DataTable dt = createDT();

            // TODO: Modify to add support for x64 programs as well
            string uninstallKey = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall";
            using (RegistryKey rk = Registry.LocalMachine.OpenSubKey(uninstallKey))
            {
                foreach (string skName in rk.GetSubKeyNames())
                {
                    using (RegistryKey sk = rk.OpenSubKey(skName))
                    {
                        try
                        {
                            // Copy over registry values to DataTable
                            if (!String.IsNullOrEmpty(sk.GetValue("DisplayName").ToString()))
                            {
                                dt.Rows.Add(sk.GetValue("DisplayName"), sk.GetValue("EstimatedSize"));
                            }
                        }
                        catch (Exception ex)
                        { }
                    }
                }

                dt.DefaultView.Sort = "DisplayName";
                dt = dt.DefaultView.ToTable();

                // Add DataTable to DataSet
                ds.Tables.Add(dt);

                // Create the Excel .xlsx file
                try
                {
                    string excelFilename = @"C:\Users\scott\Desktop\Sample.xlsx";
                    ExportDataSet(ds, excelFilename);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Couldn't create Excel file.\r\nException: " + ex.Message);
                    return;
                }

                //label1.Text += " (" + lstDisplayHardware.Items.Count.ToString() + ")";

            }
        }

        private System.Data.DataTable createDT()
        {
            // Initialize DataTable
            System.Data.DataTable dt = new System.Data.DataTable("Installed Software");
            dt.Columns.Add("DisplayName");
            dt.Columns.Add("EstimatedSize");

            return dt;
        }

        private static void ExportDataSet(DataSet ds, string destination)
        {
            var workbook = new ClosedXML.Excel.XLWorkbook();
            foreach (DataTable dt in ds.Tables)
            {
                var worksheet = workbook.Worksheets.Add(dt.TableName);
                worksheet.Cell(1, 1).InsertTable(dt);
                worksheet.Columns().AdjustToContents();
            }
            workbook.SaveAs(destination);
            workbook.Dispose();
        }
    }
}
