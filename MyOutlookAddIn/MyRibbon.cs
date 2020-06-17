using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;

namespace MyOutlookAddIn
{
    public partial class MyRibbon
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            SearchMailInfoForm mailInfoForm = new SearchMailInfoForm();
            SingleSearchInfo singleSearchInfo = SingleSearchInfo.GetInstance();
            if (singleSearchInfo.keyWord != "")
            {
                mailInfoForm.setTextBox2Text(singleSearchInfo.keyWord);
            }
            if (singleSearchInfo.folderPath != "")
            {
                mailInfoForm.setTextBox1Text(singleSearchInfo.folderPath);
            }
            mailInfoForm.Show();
        }
    }
}
