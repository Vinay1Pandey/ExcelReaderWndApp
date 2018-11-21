using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Data.SqlClient;
using System.Configuration;
using System.Data.OleDb;
using Excel = Microsoft.Office.Interop.Excel;
using System.Globalization;
using System.Net.Mail;

namespace ReadExcelFileApp
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            dataGridView1.Visible = false;
        }
        public static DataTable ConvertExcelToDataTableRevenue(string FileName)
        {
            DataTable dtResult = null;
            int totalSheet = 0; //No of sheets on excel file  
            using (OleDbConnection objConn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + FileName + ";Extended Properties='Excel 12.0;HDR=YES;IMEX=1;';"))
            {
                objConn.Open();
                OleDbCommand cmd = new OleDbCommand();
                OleDbDataAdapter oleda = new OleDbDataAdapter();
                DataSet ds = new DataSet();
                DataTable dt = objConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                string sheetName = string.Empty;
                if (dt != null)
                {
                    var tempDataTable = (from dataRow in dt.AsEnumerable()
                                         where !dataRow["TABLE_NAME"].ToString().Contains("FilterDatabase")
                                         select dataRow).CopyToDataTable();
                    dt = tempDataTable;
                    totalSheet = dt.Rows.Count;
                    //sheetName = dt.Rows[0]["TABLE_NAME"].ToString();
                    sheetName = "Revenue Data$";
                }
                cmd.Connection = objConn;
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT * FROM [" + sheetName + "]";
                oleda = new OleDbDataAdapter(cmd);
                oleda.Fill(ds, "excelData");
                dtResult = ds.Tables["excelData"];
                objConn.Close();
                return dtResult; //Returning Dattable  
            }
        }
        public static DataTable ConvertExcelToDataTableBaseData(string FileName)
        {
            DataTable dtResult = null;
            int totalSheet = 0; //No of sheets on excel file  
            using (OleDbConnection objConn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + FileName + ";Extended Properties='Excel 12.0;HDR=YES;IMEX=1;';"))
            {
                objConn.Open();
                OleDbCommand cmd = new OleDbCommand();
                OleDbDataAdapter oleda = new OleDbDataAdapter();
                DataSet ds = new DataSet();
                DataTable dt = objConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                string sheetName = string.Empty;
                if (dt != null)
                {
                    var tempDataTable = (from dataRow in dt.AsEnumerable()
                                         where !dataRow["TABLE_NAME"].ToString().Contains("FilterDatabase")
                                         select dataRow).CopyToDataTable();
                    dt = tempDataTable;
                    totalSheet = dt.Rows.Count;
                    //sheetName = dt.Rows[0]["TABLE_NAME"].ToString();
                    sheetName = "Base Data$";
                }
                cmd.Connection = objConn;
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT * FROM [" + sheetName + "]";
                oleda = new OleDbDataAdapter(cmd);
                oleda.Fill(ds, "excelData");
                dtResult = ds.Tables["excelData"];
                objConn.Close();
                return dtResult; //Returning Dattable  
            }
        }
        public static DataTable ConvertExcelToDataTableDisputes(string FileName)
        {
            DataTable dtResult = null;
            int totalSheet = 0; //No of sheets on excel file  
            using (OleDbConnection objConn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + FileName + ";Extended Properties='Excel 12.0;HDR=YES;IMEX=1;';"))
            {
                objConn.Open();
                OleDbCommand cmd = new OleDbCommand();
                OleDbDataAdapter oleda = new OleDbDataAdapter();
                DataSet ds = new DataSet();
                DataTable dt = objConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                string sheetName = string.Empty;
                if (dt != null)
                {
                    var tempDataTable = (from dataRow in dt.AsEnumerable()
                                         where !dataRow["TABLE_NAME"].ToString().Contains("FilterDatabase")
                                         select dataRow).CopyToDataTable();
                    dt = tempDataTable;
                    totalSheet = dt.Rows.Count;
                    //sheetName = dt.Rows[0]["TABLE_NAME"].ToString();
                    sheetName = "Disputes$";
                }
                cmd.Connection = objConn;
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT * FROM [" + sheetName + "]";
                oleda = new OleDbDataAdapter(cmd);
                oleda.Fill(ds, "excelData");
                dtResult = ds.Tables["excelData"];
                objConn.Close();
                return dtResult; //Returning Dattable  
            }
        }
       
        private void btnChoose_Click(object sender, EventArgs e)
        {
            string filePath = string.Empty;
            string fileExt = string.Empty;
            OpenFileDialog file = new OpenFileDialog();//open dialog to choose file
            if (file.ShowDialog() == DialogResult.OK)//if there is a file choosen by the user
            {
                filePath = file.FileName;//get the path of the file
                fileExt = Path.GetExtension(filePath);//get the file extension
                if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo(".xlsx") == 0)
                {
                    try
                    {
                        DataTable dtExcel1 = new DataTable();
                        DataTable dtExcel2 = new DataTable();
                        DataTable dtExcel3 = new DataTable();                     
                        dtExcel1 = ConvertExcelToDataTableBaseData(filePath);
                        dtExcel2 = ConvertExcelToDataTableRevenue(filePath);
                        dtExcel3 = ConvertExcelToDataTableDisputes(filePath);
                        dataGridView1.Visible = true;
                        dataGridView1.DataSource = dtExcel3;
                        processData(dtExcel1, dtExcel2, dtExcel3);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message.ToString());
                    }
                }
                else
                {
                    MessageBox.Show("Please choose .xls or .xlsx file only.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error);//custom messageBox to show error
                }
            }
        }
        public static void Email(int i,List<string> EmailID)
        {
            try
            {
                MailMessage mail = new MailMessage();
                SmtpClient SmtpServer = new SmtpClient("smtp-mail.outlook.com");
                mail.From = new MailAddress("vinay.pandey@gridinfocom.com");
                mail.Subject = "Reminder Request";

                switch (i)
                {
                    case 0:
                        mail.Body = "EM0";
                        break;
                    case 1:
                        foreach (var item in EmailID)
                        {
                            mail.To.Add(item);
                        }
                        mail.Body = "Your telephone bill numbered NNNN dated dd/mm/yy for an amount of ₹ xx,xx,xxx.xx " +
                            "(Rupees aaa bbb ccc ddd and paise eee only) is due for payment since [due date]. Please pay " +
                            "up this outstanding amount at an early date, to help us in providing you continued better " +
                            "service.Thank you for your prompt action." +
                            "Regards,";
                        break;
                    case 2:
                        foreach (var item in EmailID)
                        {
                            mail.To.Add(item);
                        }
                        mail.Body = "EM2";
                        break;
                    case 3:
                        foreach (var item in EmailID)
                        {
                            mail.To.Add(item);
                        }
                        mail.Body = "EM3";
                        break;
                    case 4:
                        foreach (var item in EmailID)
                        {
                            mail.To.Add(item);
                        }
                        mail.Body = "EM4";
                        break;
                    case 5:
                        foreach (var item in EmailID)
                        {
                            mail.To.Add(item);
                        }
                        mail.Body = "EM5";
                        break;
                    case 6:
                        foreach (var item in EmailID)
                        {
                            mail.To.Add(item);
                        }
                        mail.Body = "EM6";
                        break;
                }
                SmtpServer.Port = 587;
                SmtpServer.Credentials = new System.Net.NetworkCredential("vinay.pandey@gridinfocom.com", "jj");
                SmtpServer.EnableSsl = true;
                SmtpServer.Send(mail);
                MessageBox.Show("mail Send");
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }
        public void processData(DataTable dt1, DataTable dt2, DataTable dt3)
        {
            try
             {
                double NetAmtOS = 0;
                List<object> lstLastPaymentDate = new List<object>();
                lstLastPaymentDate = (from DataRow dr in dt1.Rows
                                      select (dr["F18"])).ToList();
                lstLastPaymentDate.Remove("Last Payment Date");
                var stringList = lstLastPaymentDate.OfType<string>();

                List<object> lstCustID = new List<object>();
                lstCustID = (from DataRow dr in dt1.Rows
                             select (dr["F2"])).ToList();
                lstCustID.Remove("Customer Id");
                var stringList1 = lstCustID.OfType<string>();

                List<object> lstEmailID = new List<object>();
                lstEmailID = (from DataRow dr in dt1.Rows
                             select (dr["F19"])).ToList();
                lstEmailID.Remove("Contact e-mail");
                var stringListEmailID = lstEmailID.OfType<string>();
                

                List<object> lstBill = new List<object>();
                lstBill = (from DataRow dr in dt2.Rows
                           select (dr["Bill Amount (₹)"])).ToList();
                var intListBill = lstBill.OfType<double>().ToList();


                List<object> lstDeposits = new List<object>();
                lstDeposits = (from DataRow dr in dt1.Rows
                               select (dr["F15"])).ToList();
                lstDeposits.Remove("Security Deposit Amount (₹)");
                var stringListDeposits = lstDeposits.OfType<string>().ToList();


                List<object> lstCRLimits = new List<object>();
                lstCRLimits = (from DataRow dr in dt1.Rows
                               select (dr["F16"])).ToList();
                lstCRLimits.Remove("Credit Limit (₹)");
                var stringListCRLimits = lstCRLimits.OfType<string>().ToList();

                List<object> lstConnectionType = new List<object>();
                lstConnectionType = (from DataRow dr in dt1.Rows
                                     select (dr["F11"])).ToList();
                lstConnectionType.Remove("Connection Type");
                var stringListConnectionType = lstConnectionType.OfType<string>().ToList();

                List<string> lstCostumerCategory = new List<string>();
                for (int i = 0; i < dt1.Rows.Count; i++)
                {
                    lstCostumerCategory.Add(dt1.Rows[i]["F10"].ToString());
                }
                //lstCostumerCategory = (from DataRow dr in dt1.Rows
                //                   select (string)dr[18]).ToList();
                lstCostumerCategory.Remove("Customer Category ");
                //var stringListCostumerCategory = lstCostumerCategory.OfType<string>().ToList();

                List<string> lstDefaults = new List<string>();
                for (int i = 0; i < dt1.Rows.Count; i++)
                {
                    lstDefaults.Add(dt1.Rows[i]["F17"].ToString());
                }
                lstDefaults.Remove("Defaults / Year");

                List<string> lstCustomerType = new List<string>();
                for (int i = 0; i < dt1.Rows.Count; i++)
                {
                    lstCustomerType.Add(dt1.Rows[i]["F4"].ToString());
                }
                lstCustomerType.Remove("Customer Type");

                List<string> lstAVGR = new List<string>();
                for (int i = 0; i < dt1.Rows.Count; i++)
                {
                    lstAVGR.Add(dt1.Rows[i]["F8"].ToString());
                }
                lstAVGR.Remove("Avg Revenue Score (AVGR)");
                List<string> lstLOYT = new List<string>();
                for (int i = 0; i < dt1.Rows.Count; i++)
                {
                    lstLOYT.Add(dt1.Rows[i]["F7"].ToString());
                }
                lstLOYT.Remove("Loyalty Score (LOYT)");

                List<string> lstDisputes = new List<string>();
                for (int i = 0; i < dt3.Rows.Count; i++)
                {
                    lstDisputes.Add(dt3.Rows[i]["Status"].ToString());
                }
                lstDisputes.Remove("Status");

                List<string> lstDisputesCust = new List<string>();
                for (int i = 0; i < dt3.Rows.Count; i++)
                {
                    lstDisputesCust.Add(dt3.Rows[i]["Customer Id"].ToString());
                }
                lstDisputesCust.Remove("Customer Id");
                List<string> lstDisputeEmail = new List<string>();
                foreach (var item in lstDisputesCust)
                {
                    var dt = dt1.Select("F2 = '"+item+"'").CopyToDataTable().DefaultView.ToTable(true, "F19");
                    lstDisputeEmail.Add(dt.Rows[0][0].ToString());
                }
                List<string> lstNoDisputeEmails = new List<string>();
                var lst= stringList1.Except(lstDisputesCust).ToList();
                foreach (var item in lst)
                {
                    var dt = dt1.Select("F2 = '" + item + "'").CopyToDataTable().DefaultView.ToTable(true, "F19");
                    lstNoDisputeEmails.Add(dt.Rows[0][0].ToString());
                }

                foreach (var item in stringList)
                {

                    if (item != "")
                    {
                        string dtS = DateTime.ParseExact(item, "dd/MM/yy", CultureInfo.InvariantCulture).ToShortDateString();
                        //string dtS = string.Format("{0:MM/dd/yy}", Convert.ToDateTime(item).ToShortDateString());
                        string dtNow = string.Format("{0:MM/dd/yyyy}", DateTime.Now);
                        double date = 0;
                        date = (Convert.ToDateTime(dtNow) - Convert.ToDateTime(dtS)).TotalDays;
                        date = 31;
                        if (date < 30)
                        {
                            //EM=0;
                            //Email(0,stringListEmailID.ToList());
                        }
                        else if (date > 30 && date < 60)
                        {
                            for (int i = 0; i < intListBill.Count; i++)
                            {
                                //if(!intListCRLimits.Contains("Credit Limit(₹)"))
                                //{
                                NetAmtOS = intListBill[i] - Convert.ToDouble(stringListDeposits[i]) - Convert.ToDouble(stringListCRLimits[i]);
                                NetAmtOS = 999;
                                //}
                                if (NetAmtOS <= 0)
                                {
                                    //Email(0,stringListEmailID.ToList());
                                }
                                else if (0 < NetAmtOS && NetAmtOS < 1000)
                                {
                                    if (stringListConnectionType[i] == "Private")
                                    {
                                        if (lstCostumerCategory[i+1] == "CIP")
                                        {
                                            if (Convert.ToInt32(lstDefaults[i+1]) > 0)
                                            {
                                                //cust
                                                if(lstCustomerType[i].ToUpper().Equals("INDIVIDUAL"))
                                                {
                                                    //avgr
                                                    if(Convert.ToInt32(lstAVGR[i])<3)
                                                    {
                                                        //loyt
                                                        if(Convert.ToInt32(lstLOYT[i])>1)
                                                        {
                                                            //dispute

                                                            //Email(0,lstDisputeEmail);

                                                            //Email(3,lstNoDisputeEmails);

                                                        }
                                                        else if(Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //Email(0,lstDisputeEmail);

                                                            //Email(4,lstNoDisputeEmails);
                                                        }
                                                    }
                                                    else if(3<Convert.ToInt32(lstAVGR[i]) && Convert.ToInt32(lstAVGR[i])<6)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //Email(0,lstDisputeEmail);

                                                            //Email(1,lstNoDisputeEmails);
                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //Email(0,lstDisputeEmail);

                                                            //Email(2,lstNoDisputeEmails);
                                                        }
                                                    }
                                                    else if(Convert.ToInt32(lstAVGR[i])>6)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //Email(0,lstDisputeEmail);

                                                            //Email(1,lstNoDisputeEmails);
                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //Email(0,lstDisputeEmail);

                                                            //Email(2,lstNoDisputeEmails);
                                                        }
                                                    }
                                                }
                                                else if(lstCustomerType[i].ToUpper().Equals("BUSINESS") || 
                                                    lstCustomerType[i].ToUpper().Equals("SERVICE") || lstCustomerType[i].ToUpper().Equals("EMER"))
                                                {
                                                    //avgr
                                                    if (Convert.ToInt32(lstAVGR[i]) < 3)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //Email(0,lstDisputeEmail);

                                                            //Email(3,lstNoDisputeEmails);
                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //Email(0,lstDisputeEmail);

                                                            //Email(4,lstNoDisputeEmails);
                                                        }
                                                    }
                                                    else if (3 < Convert.ToInt32(lstAVGR[i]) && Convert.ToInt32(lstAVGR[i]) < 6)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //Email(0,lstDisputeEmail);

                                                            //Email(1,lstNoDisputeEmails);
                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //Email(0,lstDisputeEmail);

                                                            //Email(2,lstNoDisputeEmails);
                                                        }
                                                    }
                                                    else if (Convert.ToInt32(lstAVGR[i]) > 6)
                                                    {
                                                        //Email(0,stringListEmailID.ToList());
                                                    }
                                                }
                                            }
                                            else if (Convert.ToInt32(lstDefaults[i]) == 0)
                                            {
                                                //cust
                                                if (lstCustomerType[i].ToUpper().Equals("INDIVIDUAL"))
                                                {
                                                    //avgr
                                                    if (Convert.ToInt32(lstAVGR[i]) < 3)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //Email(0,lstDisputeEmail);

                                                            //Email(1,lstNoDisputeEmails);
                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //Email(0,lstDisputeEmail);

                                                            //Email(2,lstNoDisputeEmails);
                                                        }
                                                    }
                                                    else if (3 < Convert.ToInt32(lstAVGR[i]) && Convert.ToInt32(lstAVGR[i]) < 6)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //Email(0,lstDisputeEmail);

                                                            //Email(1,lstNoDisputeEmails);
                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //Email(0,lstDisputeEmail);

                                                            //Email(2,lstNoDisputeEmails);
                                                        }
                                                    }
                                                    else if (Convert.ToInt32(lstAVGR[i]) > 6)
                                                    {
                                                        //Email(0,lstDisputeEmail);
                                                    }
                                                }
                                                else if (lstCustomerType[i].ToUpper().Equals("BUSINESS") ||
                                                    lstCustomerType[i].ToUpper().Equals("SERVICE") || lstCustomerType[i].ToUpper().Equals("EMER"))
                                                {
                                                    //Email(0,lstDisputeEmail);
                                                }
                                            }
                                        }
                                        else if (lstCostumerCategory[i] == "VIP")
                                        {
                                            if (Convert.ToInt32(lstDefaults[i]) > 0)
                                            {
                                                //cust
                                                if (lstCustomerType[i].ToUpper().Equals("INDIVIDUAL"))
                                                {
                                                    //avgr
                                                    if (Convert.ToInt32(lstAVGR[i]) < 3)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //Email(0,lstDisputeEmail);

                                                            //Email(3,lstNoDisputeEmails);
                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //Email(0,lstDisputeEmail);

                                                            //Email(4,lstNoDisputeEmails);
                                                        }
                                                    }
                                                    else if (3 < Convert.ToInt32(lstAVGR[i]) && Convert.ToInt32(lstAVGR[i]) < 6)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //Email(0,lstDisputeEmail);

                                                            //Email(1,lstNoDisputeEmails);
                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //Email(0,lstDisputeEmail);

                                                            //Email(2,lstNoDisputeEmails);
                                                        }
                                                    }
                                                    else if (Convert.ToInt32(lstAVGR[i]) > 6)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //Email(0,lstDisputeEmail);

                                                            //Email(1,lstNoDisputeEmails);
                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //Email(0,lstDisputeEmail);

                                                            //Email(2,lstNoDisputeEmails);
                                                        }
                                                    }
                                                }
                                                else if (lstCustomerType[i].ToUpper().Equals("BUSINESS") ||
                                                    lstCustomerType[i].ToUpper().Equals("SERVICE") || lstCustomerType[i].ToUpper().Equals("EMER"))
                                                {
                                                    //avgr
                                                    if (Convert.ToInt32(lstAVGR[i]) < 3)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //Email(0,lstDisputeEmail);

                                                            //Email(3,lstNoDisputeEmails);
                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //Email(0,lstDisputeEmail);

                                                            //Email(4,lstNoDisputeEmails);
                                                        }
                                                    }
                                                    else if (3 < Convert.ToInt32(lstAVGR[i]) && Convert.ToInt32(lstAVGR[i]) < 6)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //Email(0,lstDisputeEmail);

                                                            //Email(1,lstNoDisputeEmails);
                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //Email(0,lstDisputeEmail);

                                                            //Email(2,lstNoDisputeEmails);
                                                        }
                                                    }
                                                    else if (Convert.ToInt32(lstAVGR[i]) > 6)
                                                    {
                                                        //Email(0,lstDisputeEmail);
                                                    }
                                                }
                                            }
                                            else if (Convert.ToInt32(lstDefaults[i]) == 0)
                                            {
                                                //cust
                                                if (lstCustomerType[i].ToUpper().Equals("INDIVIDUAL"))
                                                {
                                                    //avgr
                                                    if (Convert.ToInt32(lstAVGR[i]) < 3)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //Email(0,lstDisputeEmail);

                                                            //Email(1,lstNoDisputeEmails);
                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //Email(0,lstDisputeEmail);

                                                            //Email(2,lstNoDisputeEmails);
                                                        }
                                                    }
                                                    else if (3 < Convert.ToInt32(lstAVGR[i]) && Convert.ToInt32(lstAVGR[i]) < 6)
                                                    {
                                                        //Email(0,lstDisputeEmail);
                                                    }
                                                    else if (Convert.ToInt32(lstAVGR[i]) > 6)
                                                    {
                                                        //Email(0,lstDisputeEmail);
                                                    }
                                                }
                                                else if (lstCustomerType[i].ToUpper().Equals("BUSINESS") ||
                                                    lstCustomerType[i].ToUpper().Equals("SERVICE") || lstCustomerType[i].ToUpper().Equals("EMER"))
                                                {
                                                    //Email(0,lstDisputeEmail);
                                                }
                                            }
                                        }
                                        else if (lstCostumerCategory[i] == "VVIP")
                                        {
                                            if (Convert.ToInt32(lstDefaults[i]) > 0)
                                            {
                                                //cust
                                                if (lstCustomerType[i].ToUpper().Equals("INDIVIDUAL"))
                                                {
                                                    //avgr
                                                    if (Convert.ToInt32(lstAVGR[i]) < 3)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //Email(0,lstDisputeEmail);

                                                            //Email(3,lstNoDisputeEmails);
                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //Email(0,lstDisputeEmail);

                                                            //Email(4,lstNoDisputeEmails);
                                                        }
                                                    }
                                                    else if (3 < Convert.ToInt32(lstAVGR[i]) && Convert.ToInt32(lstAVGR[i]) < 6)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //Email(0,lstDisputeEmail);

                                                            //Email(1,lstNoDisputeEmails);
                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //Email(0,lstDisputeEmail);

                                                            //Email(2,lstNoDisputeEmails);
                                                        }
                                                    }
                                                    else if (Convert.ToInt32(lstAVGR[i]) > 6)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //Email(0,lstDisputeEmail);

                                                            //Email(1,lstNoDisputeEmails);
                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //Email(0,lstDisputeEmail);

                                                            //Email(2,lstNoDisputeEmails);
                                                        }
                                                    }
                                                }
                                                else if (lstCustomerType[i].ToUpper().Equals("BUSINESS") ||
                                                    lstCustomerType[i].ToUpper().Equals("SERVICE") || lstCustomerType[i].ToUpper().Equals("EMER"))
                                                {
                                                    //avgr
                                                    if (Convert.ToInt32(lstAVGR[i]) < 3)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //Email(0,lstDisputeEmail);

                                                            //Email(1,lstNoDisputeEmails);
                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //Email(0,lstDisputeEmail);

                                                            //Email(2,lstNoDisputeEmails);
                                                        }
                                                    }
                                                    else if (3 < Convert.ToInt32(lstAVGR[i]) && Convert.ToInt32(lstAVGR[i]) < 6)
                                                    {
                                                        //Email(0,lstDisputeEmail);
                                                    }
                                                    else if (Convert.ToInt32(lstAVGR[i]) > 6)
                                                    {
                                                        //Email(0,lstDisputeEmail);
                                                    }
                                                }
                                            }
                                            else if (Convert.ToInt32(lstDefaults[i]) == 0)
                                            {
                                                //cust
                                                if (lstCustomerType[i].ToUpper().Equals("INDIVIDUAL"))
                                                {
                                                    //avgr
                                                    if (Convert.ToInt32(lstAVGR[i]) < 3)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //Email(0,lstDisputeEmail);

                                                            //Email(0,lstNoDisputeEmails);
                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //Email(0,lstDisputeEmail);

                                                            //Email(1,lstNoDisputeEmails);
                                                        }
                                                    }
                                                    else if (3 < Convert.ToInt32(lstAVGR[i]) && Convert.ToInt32(lstAVGR[i]) < 6)
                                                    {
                                                        //Email(0,lstDisputeEmail);
                                                    }
                                                    else if (Convert.ToInt32(lstAVGR[i]) > 6)
                                                    {
                                                        //Email(0,lstDisputeEmail);
                                                    }
                                                }
                                                else if (lstCustomerType[i].ToUpper().Equals("BUSINESS") ||
                                                    lstCustomerType[i].ToUpper().Equals("SERVICE") || lstCustomerType[i].ToUpper().Equals("EMER"))
                                                {
                                                    //Email(0,lstDisputeEmail);
                                                }
                                            }
                                        }
                                        else if (lstCostumerCategory[i] == "General")
                                        {
                                            if (Convert.ToInt32(lstDefaults[i]) > 0)
                                            {
                                                //cust
                                                if (lstCustomerType[i].ToUpper().Equals("INDIVIDUAL"))
                                                {
                                                    //avgr
                                                    if (Convert.ToInt32(lstAVGR[i]) < 3)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //Email(0,lstDisputeEmail);

                                                            //Email(5,lstNoDisputeEmails);
                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //Email(0,lstDisputeEmail);

                                                            //Email(6,lstNoDisputeEmails);
                                                        }
                                                    }
                                                    else if (3 < Convert.ToInt32(lstAVGR[i]) && Convert.ToInt32(lstAVGR[i]) < 6)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //Email(0,lstDisputeEmail);

                                                            //Email(3,lstNoDisputeEmails);
                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //Email(0,lstDisputeEmail);

                                                            //Email(4,lstNoDisputeEmails);
                                                        }
                                                    }
                                                    else if (Convert.ToInt32(lstAVGR[i]) > 6)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //Email(0,lstDisputeEmail);

                                                            //Email(3,lstNoDisputeEmails);
                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //Email(0,lstDisputeEmail);

                                                            //Email(4,lstNoDisputeEmails);
                                                        }
                                                    }
                                                }
                                                else if (lstCustomerType[i].ToUpper().Equals("BUSINESS") ||
                                                    lstCustomerType[i].ToUpper().Equals("SERVICE") || lstCustomerType[i].ToUpper().Equals("EMER"))
                                                {
                                                    //avgr
                                                    if (Convert.ToInt32(lstAVGR[i]) < 3)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //Email(0,lstDisputeEmail);

                                                            //Email(5,lstNoDisputeEmails);
                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //Email(0,lstDisputeEmail);

                                                            //Email(6,lstNoDisputeEmails);
                                                        }
                                                    }
                                                    else if (3 < Convert.ToInt32(lstAVGR[i]) && Convert.ToInt32(lstAVGR[i]) < 6)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //Email(0,lstDisputeEmail);

                                                            //Email(3,lstNoDisputeEmails);
                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //Email(0,lstDisputeEmail);

                                                            //Email(4,lstNoDisputeEmails);
                                                        }
                                                    }
                                                    else if (Convert.ToInt32(lstAVGR[i]) > 6)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //Email(0,lstDisputeEmail);

                                                            //Email(1,lstNoDisputeEmails);
                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //Email(0,lstDisputeEmail);

                                                            //Email(2,lstNoDisputeEmails);
                                                        }
                                                    }
                                                }
                                            }
                                            else if (Convert.ToInt32(lstDefaults[i]) == 0)
                                            {
                                                //cust
                                                if (lstCustomerType[i].ToUpper().Equals("INDIVIDUAL"))
                                                {
                                                    //avgr
                                                    if (Convert.ToInt32(lstAVGR[i]) < 3)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //Email(0,lstDisputeEmail);

                                                            //Email(5,lstNoDisputeEmails);
                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //Email(0,lstDisputeEmail);

                                                            //Email(6,lstNoDisputeEmails);
                                                        }
                                                    }
                                                    else if (3 < Convert.ToInt32(lstAVGR[i]) && Convert.ToInt32(lstAVGR[i]) < 6)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //Email(0,lstDisputeEmail);

                                                            //Email(3,lstNoDisputeEmails);
                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //Email(0,lstDisputeEmail);

                                                            //Email(4,lstNoDisputeEmails);
                                                        }
                                                    }
                                                    else if (Convert.ToInt32(lstAVGR[i]) > 6)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //Email(0,lstDisputeEmail);

                                                            //Email(3,lstNoDisputeEmails);
                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //Email(0,lstDisputeEmail);

                                                            //Email(4,lstNoDisputeEmails);
                                                        }
                                                    }
                                                }
                                                else if (lstCustomerType[i].ToUpper().Equals("BUSINESS") ||
                                                    lstCustomerType[i].ToUpper().Equals("SERVICE") || lstCustomerType[i].ToUpper().Equals("EMER"))
                                                {
                                                    //avgr
                                                    if (Convert.ToInt32(lstAVGR[i]) < 3)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //Email(0,lstDisputeEmail);

                                                            //Email(5,lstNoDisputeEmails);
                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //Email(0,lstDisputeEmail);

                                                            //Email(6,lstNoDisputeEmails);
                                                        }
                                                    }
                                                    else if (3 < Convert.ToInt32(lstAVGR[i]) && Convert.ToInt32(lstAVGR[i]) < 6)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //Email(0,lstDisputeEmail);

                                                            //Email(3,lstNoDisputeEmails);
                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //Email(0,lstDisputeEmail);

                                                            //Email(4,lstNoDisputeEmails);
                                                        }
                                                    }
                                                    else if (Convert.ToInt32(lstAVGR[i]) > 6)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //Email(0,lstDisputeEmail);

                                                            //Email(1,lstNoDisputeEmails);
                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //Email(0,lstDisputeEmail);

                                                            //Email(2,lstNoDisputeEmails);
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        if (lstCostumerCategory[i].ToUpper().Equals("CIP"))
                                        {
                                        }
                                        else if (lstCostumerCategory[i].ToUpper().Equals("VIP"))
                                        {
                                        }
                                        else if (lstCostumerCategory[i].ToUpper().Equals("VVIP"))
                                        {
                                        }
                                        else if (lstCostumerCategory[i].ToUpper().Equals("GENERAL"))
                                        {
                                        }
                                    }

                                }
                                else if (1000 < NetAmtOS && NetAmtOS < 3000)
                                {

                                }
                                else if (3000 < NetAmtOS && NetAmtOS < 10000)
                                {

                                }
                                else if (10000 < NetAmtOS && NetAmtOS < 25000)
                                {

                                }
                                else if (25000 < NetAmtOS && NetAmtOS < 50000)
                                {

                                }
                                else if (50000 < NetAmtOS && NetAmtOS < 100000)
                                {

                                }
                                else if (100000 < NetAmtOS && NetAmtOS < 500000)
                                {

                                }
                                else if (NetAmtOS > 500000)
                                {

                                }

                            }
                        }
                        else if (60 < date && date < 90)
                        {

                        }
                        else if (date > 90)
                        {

                        }
                    }

                }


            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }

        }
        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();//to close the window(Form1)
        }       
    }
}
