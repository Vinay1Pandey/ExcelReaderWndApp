using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
namespace ReadExcelFileApp
{
    public class WriteReminderExcel
    {
        public static void UpdateReminder(string CustID,string filePath,DataTable dt1)
        {
           DataTable dt = dt1.Select("");
            Application xlApp = new Application();
            xlApp.Workbooks.Open(filePath);
            Worksheet wrksheet = new Worksheet();
            .CopyToDataTable().DefaultView.ToTable(true, "No# of reminders sent");
        }
    }
}
