using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System.IO;

namespace MyOutlookAddIn
{
    class NPOIManage
    {
        private XSSFWorkbook wb = null;
        private ISheet sheet = null;
        private IRow row = null;

        public NPOIManage()
        {
            wb = new XSSFWorkbook(Config.GetFolderPath());
            sheet = wb.GetSheetAt(0);
        }

        public void AddLineToSheet(List<string> str)
        {
            row = sheet.CreateRow(sheet.LastRowNum + 1);
            for (int i = 0; i < str.Count; i++)
            {
                if (0 == i)
                {
                    row.CreateCell(i).SetCellValue(sheet.LastRowNum.ToString());
                }
                else
                {
                    row.CreateCell(i).SetCellValue(str[i]);
                }
            }
        }

        public void Close()
        {
            string tempFilePath = "~~" + Config.GetFolderPath();
            FileStream fileStream = File.Create(tempFilePath);
            wb.Write(fileStream);
            wb.Close();
            fileStream.Close();
            File.Delete(Config.GetFolderPath());
            File.Move(tempFilePath, Config.GetFolderPath());
        }
    }
}
