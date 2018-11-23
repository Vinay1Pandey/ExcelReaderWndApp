using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;

namespace ReadExcelFileApp
{
    public class WriteReminderExcel
    {
        public static void UpdateReminder(string CustID, string filePath, string NoOfRem)
        {
            try
            {
                using (OleDbConnection objConn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties='Excel 12.0;HDR=YES;';"))
                {
                    //string connetionString = null;
                    // OleDbConnection connection;
                    OleDbDataAdapter oledbAdapter = new OleDbDataAdapter();
                    string sql = null;
                    //connetionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Your mdb filename;";
                    //connection = new OleDbConnection(connetionString);
                    string sheetName = "Revenue Data$";
                    int i = Convert.ToInt32(NoOfRem) + 1;
                    sql = "update [" + sheetName + "] SET [No# of reminders sent] = '"+ i.ToString() +"' where [Customer Id] = '" + CustID+"'";


                    objConn.Open();
                    oledbAdapter.UpdateCommand = objConn.CreateCommand();
                    oledbAdapter.UpdateCommand.CommandText = sql;
                    oledbAdapter.UpdateCommand.ExecuteNonQuery();
                    //MessageBox.Show("Row(s) Updated !! ");

                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
            
               
        }

        public static void UpDateStatus(string CustID, string filePath, string status)
        {
            try
            {
                using (OleDbConnection objConn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties='Excel 12.0;HDR=YES;';"))
                {
                    //string connetionString = null;
                    // OleDbConnection connection;
                    OleDbDataAdapter oledbAdapter = new OleDbDataAdapter();
                    string sql = null;
                    //connetionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Your mdb filename;";
                    //connection = new OleDbConnection(connetionString);
                    string sheetName = "Base Data$";
                    sql = "update [" + sheetName + "] SET [F20] = '" +status+ "' where [F2] = '" + CustID + "'";



                    objConn.Open();
                    oledbAdapter.UpdateCommand = objConn.CreateCommand();
                    oledbAdapter.UpdateCommand.CommandText = sql;
                    oledbAdapter.UpdateCommand.ExecuteNonQuery();
                    //MessageBox.Show("Row(s) Updated !! ");

                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }


        }
    }
}
