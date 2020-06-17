using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace MyOutlookAddIn
{
    public partial class SearchMailInfoForm : Form
    {
        public SearchMailInfoForm()
        {
            InitializeComponent();
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            //FolderBrowserDialog fileName = new FolderBrowserDialog();
            SaveFileDialog fileDialog = new SaveFileDialog();
            //fileDialog.InitialDirectory = "C:\\";
            fileDialog.Filter = "Excel files(*.xlsx)|*.xlsx";
            fileDialog.FileName = "MailDate";
            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = fileDialog.FileName;
                this.textBox1.Text = filePath;
            }
            fileDialog.InitialDirectory = this.textBox1.Text;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (this.textBox1.Text.Trim().Equals(""))
            {

            }
            else
            {
                SingleSearchInfo singleSearchInfo = SingleSearchInfo.GetInstance();
                singleSearchInfo.fromDateTime = this.dateTimePicker1.Value;
                singleSearchInfo.toDateTime = this.dateTimePicker2.Value;
                singleSearchInfo.keyWord = textBox2.Text.Trim();
                singleSearchInfo.folderPath = this.textBox1.Text.Trim();

                MailInfo.Save_MailInfo();
            }
            this.Close();
        }

        public void setDateTimePicker1(DateTime dateTime)
        {
            this.dateTimePicker1.Value = dateTime;
        }

        public void setDateTimePicker2(DateTime dateTime)
        {
            this.dateTimePicker2.Value = dateTime;
        }

        public void setTextBox1Text(string text)
        {
            this.textBox1.Text = text.Trim();
        }

        public void setTextBox2Text(string text)
        {
            this.textBox2.Text = text.Trim();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            button2_Click(sender, e);
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            this.Close();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
