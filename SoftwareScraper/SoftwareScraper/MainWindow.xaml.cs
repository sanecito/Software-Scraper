// This code is under the terms of the GNU General Public License v3.0 as published by the Free Software Foundation. 
// A copy of the GNU General Public License v3.0 can be found at http://www.gnu.org/licenses/.

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
using System.IO;

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
                                dt.Rows.Add(sk.GetValue("DisplayName"),
                                    sk.GetValue("DisplayVersion"),
                                    sk.GetValue("Publisher"),
                                    sk.GetValue("VersionMajor"),
                                    sk.GetValue("VersionMinor"),
                                    sk.GetValue("Version"),
                                    sk.GetValue("InstallDate"),
                                    sk.GetValue("InstallLocation"),
                                    sk.GetValue("InstallSource"),
                                    sk.GetValue("EstimatedSize"),
                                    sk.GetValue("Readme"),
                                    sk.GetValue("UninstallString"));
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
                    ExportDataSet(ds);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Couldn't create Excel file.\r\nException: " + ex.Message);

                    return;
                }
            }
        }

        private System.Data.DataTable createDT()
        {
            // Initialize DataTable
            System.Data.DataTable dt = new System.Data.DataTable("Installed Software");
            dt.Columns.Add("DisplayName");
            dt.Columns.Add("DisplayVersion");
            dt.Columns.Add("Publisher");
            dt.Columns.Add("VersionMajor");
            dt.Columns.Add("VersionMinor");
            dt.Columns.Add("Version");
            dt.Columns.Add("InstallDate");
            dt.Columns.Add("InstallLocation");
            dt.Columns.Add("InstallSource");
            dt.Columns.Add("EstimatedSize");
            dt.Columns.Add("Readme");
            dt.Columns.Add("UninstallString");
            
            return dt;
        }

        private static void ExportDataSet(DataSet ds)
        {
            var workbook = new ClosedXML.Excel.XLWorkbook();
            foreach (DataTable dt in ds.Tables)
            {
                var worksheet = workbook.Worksheets.Add(dt.TableName);
                worksheet.Cell(1, 1).InsertTable(dt);
                worksheet.Columns().AdjustToContents();
            }

            SaveFileDialog sFD = new SaveFileDialog();
            sFD.Filter = "Excel file (.xlsx)|*.xlsx";
            if (sFD.ShowDialog() == true)
            {
                workbook.SaveAs(sFD.FileName);
                workbook.Dispose();
            }        
        }
    }
}