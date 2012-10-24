using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;


namespace CPCSSNReport
{
    public partial class ReportMain : Form
    {
        WordAccess wordaccess;

        public ReportMain()
        {
            InitializeComponent();            
        }

        private void cmdByProvider_Click(object sender, EventArgs e)
        {
            if (cboProviders.SelectedItems.Count == 1)
            {
                wordaccess.CreateDocByProvider(cboProviders.SelectedItem.ToString());            
            }
            else
            {
                HashSet<String> providers = new HashSet<string>();
                foreach (var v in cboProviders.SelectedItems)
                {
                    if (v.ToString() != "All")
                    {
                        providers.Add(v.ToString());
                    }
                }
                wordaccess.CreateDocByGroup(providers);
            }
            
            MessageBox.Show("Provider reports are created");
        }

        private void cmdAll_Click(object sender, EventArgs e)
        {            
            wordaccess.CreateDocByAll();
            MessageBox.Show("Summary report is created");
        }

        private void cmdByPractice_Click(object sender, EventArgs e)
        {            
            wordaccess.CreateDocByPractice(cboPractices.SelectedItem.ToString());
            MessageBox.Show("Practice reports are created");
        }

        private void cmdConnectDB_Click(object sender, EventArgs e)
        {
            if (wordaccess == null)
            {
                string outputPath = txtOutput.Text;
                if (outputPath.Length == 0)
                {
                    System.Reflection.Assembly asm = System.Reflection.Assembly.GetExecutingAssembly();
                    string fullProcessPath = asm.Location;
                    outputPath = System.IO.Path.GetDirectoryName(fullProcessPath);
                }
                wordaccess = new WordAccess(txtDBCurrent.Text, txtDBPrev.Text, txtTemplate.Text, txtExcelTemplate.Text, outputPath);
            }

            if (!cboProviders.Items.Contains("All"))
                cboProviders.Items.Add("All");
            foreach (string s in wordaccess.GetProviders())
            {
                if (!cboProviders.Items.Contains(s))
                {
                    cboProviders.Items.Add(s);
                }
            }
            cboProviders.SelectedItem = "All";

            if (!cboPractices.Items.Contains("All"))
                cboPractices.Items.Add("All");
            foreach (string s in wordaccess.GetPractices())
            {
                if (!cboPractices.Items.Contains(s))
                {
                    cboPractices.Items.Add(s);
                }
            }
            cboPractices.SelectedItem = "All";
        }

        private void textBox1_MouseDown(object sender, MouseEventArgs e)
        {
            OpenFileDialog ofDlg = new OpenFileDialog();
            if (ofDlg.ShowDialog() == DialogResult.OK)
            {
                txtTemplate.Text = ofDlg.FileName;
            }
        }

        private void txtDBCurrent_MouseDown(object sender, MouseEventArgs e)
        {
            OpenFileDialog ofDlg = new OpenFileDialog();
            if (ofDlg.ShowDialog() == DialogResult.OK)
            {
                txtDBCurrent.Text = ofDlg.FileName;
            }
        }

        private void txtDBPrev_MouseDown(object sender, MouseEventArgs e)
        {
            OpenFileDialog ofDlg = new OpenFileDialog();
            if (ofDlg.ShowDialog() == DialogResult.OK)
            {
                txtDBPrev.Text = ofDlg.FileName;
            }
        }

        private void cmdAppend_Click(object sender, EventArgs e)
        {
            wordaccess.OpenBook();
            wordaccess.CreateSheetByAll();
            MessageBox.Show("Summary data are appended");
        }

        private void cmdAppendByProvider_Click(object sender, EventArgs e)
        {
            if (cboProviders.SelectedItems.Count == 1)
            {
                wordaccess.OpenBook();
                wordaccess.CreateSheetByProvider(cboProviders.SelectedItem.ToString());
            }
            else
            {
                HashSet<String> providers = new HashSet<string>();
                foreach (var v in cboProviders.SelectedItems)
                {
                    if (v.ToString() != "All")
                    {
                        providers.Add(v.ToString());
                    }
                }
                wordaccess.CreateSheetByGroup(providers);
            }
                        
            MessageBox.Show("Provider data are appended");
        }

        private void cmdAppendByPractice_Click(object sender, EventArgs e)
        {
            wordaccess.OpenBook();
            wordaccess.CreateSheetByPractice(cboPractices.SelectedItem.ToString());
            MessageBox.Show("Practice data are appended");
        }

        private void cmdSaveExcel_Click(object sender, EventArgs e)
        {
            wordaccess.CloseBook();
        }

        private void txtExcelTemplate_MouseDown(object sender, MouseEventArgs e)
        {
            OpenFileDialog ofDlg = new OpenFileDialog();
            if (ofDlg.ShowDialog() == DialogResult.OK)
            {
                txtExcelTemplate.Text = ofDlg.FileName;
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (wordaccess != null)
            {
                wordaccess.QuitBook();
            }
        }

        private void txtOutput_MouseDown(object sender, MouseEventArgs e)
        {
            FolderBrowserDialog ofDlg = new FolderBrowserDialog();
            if (ofDlg.ShowDialog() == DialogResult.OK)
            {
                txtOutput.Text = ofDlg.SelectedPath;
            }
        }
    }
}
