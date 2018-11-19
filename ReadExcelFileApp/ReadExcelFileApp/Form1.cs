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
                        DataTable dtExcel4 = new DataTable();
                        dtExcel1 = ReadExcel1(filePath, fileExt);//read excel file
                        dtExcel2 = ReadExcel2(filePath, fileExt);
                        dtExcel3 = ReadExcel3(filePath, fileExt);
                        dtExcel4 = ReadExcel4(filePath, fileExt);

                        dataGridView1.Visible = true;
                        dataGridView1.DataSource = dtExcel1;
                        processData(dtExcel1,dtExcel2,dtExcel3,dtExcel4);
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
        public void processData(DataTable dt1,DataTable dt2,DataTable dt3,DataTable dt4)
        {
            try
            {
                List<object> lstLastPaymentDate = new List<object>();
                lstLastPaymentDate = (from DataRow dr in dt1.Rows
                                      select (object)dr[26]).ToList();
                
                List<object> lstCustID = new List<object>();
                lstCustID = (from DataRow dr in dt1.Rows
                             select (object)dr[10]).ToList();

                List<object> lstBill = new List<object>();
                lstBill = (from DataRow dr in dt2.Rows
                           select (object)dr[12]).ToList();

                List<object> lstDeposits = new List<object>();
                lstDeposits = (from DataRow dr in dt1.Rows
                               select (object)dr[23]).ToList();

                List<object> lstCRLimits = new List<object>();
                lstCRLimits = (from DataRow dr in dt1.Rows
                               select (object)dr[24]).ToList();
                foreach(var item in lstLastPaymentDate)
                {
                    if (!lstLastPaymentDate.Contains(DBNull.Value))
                    {
                        double date=0;
                        date = (DateTime.Now - Convert.ToDateTime(item)).TotalDays;
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

        public DataTable ReadExcel1(string fileName, string fileExt)
        {
            string conn = string.Empty;
            DataTable dtexcel = new DataTable();
            if (fileExt.CompareTo(".xls") == 0)//compare the extension of the file
                conn = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';";//for below excel 2007
            else
                conn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0;HDR=Yes;IMEX=1';";//for above excel 2007
            using (OleDbConnection con = new OleDbConnection(conn))
            {
                try
                {
                    con.Open();
                    OleDbDataAdapter oleAdpt = new OleDbDataAdapter();
                    dtexcel = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    string SheetName = dtexcel.Rows[0]["TABLE_NAME"].ToString();
                    OleDbCommand cmdExcel = new OleDbCommand();
                    cmdExcel.Connection = con;
                    SheetName = "Base Data$";
                    cmdExcel.CommandText = "SELECT * From [" + SheetName + "]";
                    oleAdpt.SelectCommand = cmdExcel;
                    oleAdpt.Fill(dtexcel);//fill excel data into dataTable
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString());
                }
            }
            return dtexcel;
        }
        public DataTable ReadExcel2(string fileName, string fileExt)
        {
            string conn = string.Empty;
            DataTable dtexcel = new DataTable();
            if (fileExt.CompareTo(".xls") == 0)//compare the extension of the file
                conn = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';";//for below excel 2007
            else
                conn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0;HDR=Yes;IMEX=1';";//for above excel 2007
            using (OleDbConnection con = new OleDbConnection(conn))
            {
                try
                {
                    con.Open();
                    OleDbDataAdapter oleAdpt = new OleDbDataAdapter();
                    dtexcel = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    string SheetName = dtexcel.Rows[0]["TABLE_NAME"].ToString();
                    OleDbCommand cmdExcel = new OleDbCommand();
                    cmdExcel.Connection = con;
                    SheetName = "Revenue Data$";
                    cmdExcel.CommandText = "SELECT * From [" + SheetName + "]";
                    oleAdpt.SelectCommand = cmdExcel;
                    oleAdpt.Fill(dtexcel);//fill excel data into dataTable
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString());
                }
            }
            return dtexcel;
        }
        public DataTable ReadExcel3(string fileName, string fileExt)
        {
            string conn = string.Empty;
            DataTable dtexcel = new DataTable();
            if (fileExt.CompareTo(".xls") == 0)//compare the extension of the file
                conn = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';";//for below excel 2007
            else
                conn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0;HDR=Yes;IMEX=1';";//for above excel 2007
            using (OleDbConnection con = new OleDbConnection(conn))
            {
                try
                {
                    con.Open();
                    OleDbDataAdapter oleAdpt = new OleDbDataAdapter();
                    dtexcel = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    string SheetName = dtexcel.Rows[0]["TABLE_NAME"].ToString();
                    OleDbCommand cmdExcel = new OleDbCommand();
                    cmdExcel.Connection = con;
                    SheetName = "Disputes$";
                    cmdExcel.CommandText = "SELECT * From [" + SheetName + "]";
                    oleAdpt.SelectCommand = cmdExcel;
                    oleAdpt.Fill(dtexcel);//fill excel data into dataTable
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString());
                }
            }
            return dtexcel;
        }
        public DataTable ReadExcel4(string fileName, string fileExt)
        {
            string conn = string.Empty;
            DataTable dtexcel = new DataTable();
            if (fileExt.CompareTo(".xls") == 0)//compare the extension of the file
                conn = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';";//for below excel 2007
            else
                conn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0;HDR=Yes;IMEX=1';";//for above excel 2007
            using (OleDbConnection con = new OleDbConnection(conn))
            {
                try
                {
                    con.Open();
                    OleDbDataAdapter oleAdpt = new OleDbDataAdapter();
                    dtexcel = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    string SheetName = dtexcel.Rows[0]["TABLE_NAME"].ToString();
                    OleDbCommand cmdExcel = new OleDbCommand();
                    cmdExcel.Connection = con;
                    SheetName = "Data_Points$";
                    cmdExcel.CommandText = "SELECT * From [" + SheetName + "]";
                    oleAdpt.SelectCommand = cmdExcel;
                    oleAdpt.Fill(dtexcel);//fill excel data into dataTable
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString());
                }
            }
            return dtexcel;
        }
    }
}
