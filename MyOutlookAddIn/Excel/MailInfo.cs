using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace MyOutlookAddIn
{
    class MailInfo
    {
        public static List<string> _lineInfo = new List<string>();

        public static void Save_MailInfo()
        {
            SingleMailInfo singleMailInfo = SingleMailInfo.GetInstance();
            SingleSearchInfo singleSearchInfo = SingleSearchInfo.GetInstance();
            // 打开Excel
            ExcelManage excelManage = new ExcelManage();
            // 向Excel中写入标题行
            _lineInfo.Clear();
            _lineInfo.Add("No.");
            _lineInfo.Add("DateTime.");
            _lineInfo.Add("Sender.");
            _lineInfo.Add("Receiver.");
            _lineInfo.Add("CC.");
            _lineInfo.Add("Subject.");
            excelManage.WriteToSheet(_lineInfo);
            for (int i = 0; i < singleMailInfo.mailDateTime.Count; i++)
            {
                try
                {
                    // 在日期内
                    bool isGreatFrom = DateTime.Compare(Convert.ToDateTime(singleMailInfo.mailDateTime[i]), singleSearchInfo.fromDateTime) >= 0;
                    bool isLessTo = DateTime.Compare(Convert.ToDateTime(singleMailInfo.mailDateTime[i]), singleSearchInfo.toDateTime) <= 0;
                    if (isGreatFrom && isLessTo)
                    {
                        // 包含关键字
                        if (singleMailInfo.mailSubject[i].Contains(singleSearchInfo.keyWord))
                        {
                            // 写入Excel
                            _lineInfo.Clear();
                            _lineInfo.Add("No.");
                            _lineInfo.Add(singleMailInfo.mailDateTime[i]);
                            _lineInfo.Add(singleMailInfo.mailAddresser[i]);
                            _lineInfo.Add(singleMailInfo.mailTo[i]);
                            _lineInfo.Add(singleMailInfo.mailCC[i]);
                            _lineInfo.Add(singleMailInfo.mailSubject[i]);
                            excelManage.WriteToSheet(_lineInfo);
                        }
                    }
                }
                catch { }
            }
            SingleMailInfo.ClearMailConts();
            // 设置Excel格式
            excelManage.SetFormat();
            // 保存Excel
            //string excelPath = singleMailInfo.folderPath + "\\MailDate.xlsx";
            excelManage.SaveExcel(singleSearchInfo.folderPath);
            // 关闭Excel
            excelManage.Close();
        }

        public static void Add_MailInfo()
        {
            NPOIManage nPOIManage = new NPOIManage();
            SingleMailInfo singleMailInfo = SingleMailInfo.GetInstance();
            for (int i = 0; i < singleMailInfo.mailDateTime.Count; i++)
            {
                try
                {
                    // 包含关键字
                    if (singleMailInfo.mailSubject[i].Contains(Config.GetKeyWord()))
                    {
                        // 写入Excel
                        _lineInfo.Clear();
                        _lineInfo.Add("No.");
                        _lineInfo.Add(singleMailInfo.mailDateTime[i]);
                        _lineInfo.Add(singleMailInfo.mailAddresser[i]);
                        _lineInfo.Add(singleMailInfo.mailTo[i]);
                        _lineInfo.Add(singleMailInfo.mailCC[i]);
                        _lineInfo.Add(singleMailInfo.mailSubject[i]);
                        nPOIManage.AddLineToSheet(_lineInfo);
                    }
                }
                catch { }
            }
            SingleMailInfo.ClearMailConts();
            nPOIManage.Close();
        }
    }
}
