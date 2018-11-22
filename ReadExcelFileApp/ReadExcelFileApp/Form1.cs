﻿using System;
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
                        dtExcel1 = ReadExcel.ConvertExcelToDataTableBaseData(filePath);
                        dtExcel2 = ReadExcel.ConvertExcelToDataTableRevenue(filePath);
                        dtExcel3 = ReadExcel.ConvertExcelToDataTableDisputes(filePath);
                        dataGridView1.Visible = true;
                        dataGridView1.DataSource = dtExcel1;
                        ProcessData.processData(dtExcel1, dtExcel2, dtExcel3);
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
        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();//to close the window(Form1)
        }       
    }
}
