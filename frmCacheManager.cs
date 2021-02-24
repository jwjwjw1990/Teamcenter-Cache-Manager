using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SolidEdgeCommunity;
using Microsoft.Win32;
using System.IO;
using System.Text.RegularExpressions;

namespace Teamcenter_Cache_Manager
{
    public partial class frmCacheManager : Form
    {

        RegistryKey key = null;
        string workspaceFolder = null;
        string prefixWorkspaces = null;

        SolidEdgeFramework.Application app = null;
        public frmCacheManager()
        {
            InitializeComponent();
        }

        private void frmCacheManager_Load(object sender, EventArgs e)
        {
            this.comboBox1.KeyPress += new System.Windows.Forms.KeyPressEventHandler(CheckKeyPress);
            this.KeyPress += new System.Windows.Forms.KeyPressEventHandler(CheckKeyPress);

            try
            {
                OleMessageFilter.Register();
                app = SolidEdgeUtils.Connect(false, true);
            }
            catch (System.Exception ex)
            {
                Console.WriteLine(ex.Message);
                throw;
            }

            bool tcMode = false;

            if (app != null)
            {
                app.SolidEdgeTCE.GetTeamCenterMode(out tcMode);

                if (tcMode == false)
                    return;

                string keyname = app.RegistryPath + @"\General";
                key = RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.CurrentUser, RegistryView.Registry64); 
                key = key.OpenSubKey(keyname, true);
                string seecPath = key.GetValue(@"SEEC cache").ToString();
                string currentWS = key.GetValue(@"SEEC_Active_Project").ToString();
                prefixWorkspaces = currentWS.Substring(0, currentWS.Length - (currentWS.Split('\\').Last().Length + 1)) + @"\";

                workspaceFolder = Directory.GetParent(seecPath + @"\" + currentWS).FullName;

                string[] workspaces = Directory .GetDirectories(workspaceFolder)
                                                .OrderByDescending(Directory.GetLastWriteTime)
                                                .Select(Path.GetFileName)
                                                .ToArray();

                comboBox1.Items.AddRange(workspaces);
                comboBox1.Text = currentWS.Split('\\').Last();

            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SetWorkspace();
        }

        private void SetWorkspace()
        {
            if (!comboBox1.Text.All(c => Char.IsLetterOrDigit(c) || c == ' ' || c == '_' || c == '-'))
            {
                MessageBox.Show("De ingegeven naam bevat tekens die niet geldig zijn. Alleen alphanumerieke waardes of getallen zijn toegestaan", "Foute invoer");
                return;
            }

            app.SolidEdgeTCE.CreateNewProject(comboBox1.Text);
            app.DoIdle();
            this.Close();
        }

        private void CheckKeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Return)

            {
                SetWorkspace();
            }

            if (e.KeyChar == (char)Keys.Escape)

            {
                this.Close();
            }
        }

        private void comboBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)

            {
                SetWorkspace();
            }

            if (e.KeyCode == Keys.Escape)

            {
                this.Close();
            }
        }

        private void frmCacheManager_FormClosing(object sender, FormClosingEventArgs e)
        {
            OleMessageFilter.Unregister();
            app = null;
        }
    }
}
