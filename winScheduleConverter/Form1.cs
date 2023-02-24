using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using winScheduleConverter.Services;

namespace winScheduleConverter
{
    public partial class fmwinSchConverter : Form
    {
        private string ntxtfile = "";
        private string noutfile = "";
        public fmwinSchConverter()
        {
            InitializeComponent();
        }
        private void fmwinSchConverter_Load(object sender, EventArgs e)
        {
        
            ntxtfile = "";
            noutfile = "";
            txtFile.Text = ntxtfile;
            opfSchFile.Title = "Browse Excel Files";
            opfSchFile.DefaultExt = "xlsx";
            opfSchFile.FileName = "";
            opfSchFile.Filter = "Excel files *.xlsx|*.xlsx";
            opfSchFile.CheckFileExists = true;
            opfSchFile.CheckPathExists = true;
            txtFile.Text = ntxtfile;
            txtProgress.Text = "";
            txtoutputfile.Text = "";
            txtYear.Text = DateTime.Now.Year.ToString();
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            noutfile = "";
            txtProgress.Text = "";
            txtoutputfile.Text = "";
            txtProgress.Visible = false;
            txtoutputfile.Visible = false;
            if (opfSchFile.ShowDialog() == DialogResult.OK)
            {
                txtFile.Text = opfSchFile.FileName;
                ntxtfile = txtFile.Text;
            }
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
           

            try
            {
                if (string.IsNullOrEmpty(txtFile.Text))
                {
                    MessageBox.Show("Please select file!");
                    return; 
                }
                if (string.IsNullOrEmpty(txtYear.Text))
                {
                    MessageBox.Show("Please enter start of the input year as per sheet!");
                    return;
                }

                noutfile = "";
                btnBrowse.Enabled = false;
                btnRun.Enabled = false;
                txtProgress.Visible = false;
                txtoutputfile.Visible = false;
             
                //txtProgress.Visible = true;
                //Thread.Sleep(1000);
                TaskManagerService tm = new TaskManagerService();
                noutfile = Path.GetDirectoryName(ntxtfile) + "\\" + Path.GetFileNameWithoutExtension(ntxtfile) + "_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv";
                var result = tm.ConvertSchedule(ntxtfile, noutfile,Convert.ToInt32(txtYear.Text));
                if (result)
                {
                    MessageBox.Show("File converted successfully!");
                }
                else
                {
                    MessageBox.Show("Unable to convert file!");
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show("Error occoured:\n"+ex.Message);
            }
            finally
            {
            
                
                btnBrowse.Enabled = true;
                btnRun.Enabled = true;
                txtProgress.Visible = true;
                txtoutputfile.Visible = true;
                txtProgress.Text = "Saved at:";
                txtoutputfile.Text = noutfile;
                //txtProgress.Visible = false;
            }
        }


    }
}
