using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace SoftwareScraper
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void action_btn_get_Click(object sender, EventArgs e)
        {
            // TODO: Modify to add support for x64 as well
            string uninstallKey = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall";
            using (RegistryKey rk = Registry.LocalMachine.OpenSubKey(uninstallKey))
            {
                foreach (string skName in rk.GetSubKeyNames())
                {
                    using (RegistryKey sk = rk.OpenSubKey(skName))
                    {
                        try
                        {

                            var displayName = sk.GetValue("DisplayName");
                            var size = sk.GetValue("EstimatedSize");

                            ListViewItem item;
                            if (displayName != null)
                            {
                                if (size != null)
                                    item = new ListViewItem(new string[] {displayName.ToString(), 
                                                       size.ToString()});
                                else
                                    item = new ListViewItem(new string[] { displayName.ToString() });
                                lstDisplayHardware.Items.Add(item);
                            }
                        }
                        catch (Exception ex)
                        { }
                    }
                }
                label1.Text += " (" + lstDisplayHardware.Items.Count.ToString() + ")";
            }  
        }

    }
}
