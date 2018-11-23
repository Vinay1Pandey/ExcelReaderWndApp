using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ReadExcelFileApp
{
    public class ProcessData
    {
        public static void processData(DataTable dt1, DataTable dt2, DataTable dt3,string filepath)
        {
            try
            {
                double NetAmtOS = 0;
                List<object> lstCustID = new List<object>();
                lstCustID = (from DataRow dr in dt1.Rows
                             select (dr["F2"])).ToList();
                lstCustID.Remove("Customer Id");
                var stringList1 = lstCustID.OfType<string>();

                List<string> lstDisputesCust = new List<string>();
                for (int i = 0; i < dt3.Rows.Count; i++)
                {
                    lstDisputesCust.Add(dt3.Rows[i]["Customer Id"].ToString());
                }
                lstDisputesCust.Remove("Customer Id");

                var lstNoDisputeCustID = stringList1.Except(lstDisputesCust).ToList();
                List<string> lstNoDisputeEmails = new List<string>();

                foreach (var item in lstNoDisputeCustID)
                {
                    var dt = dt1.Select("F2 = '" + item + "'").CopyToDataTable().DefaultView.ToTable(true, "F19");
                    lstNoDisputeEmails.Add(dt.Rows[0][0].ToString());
                }
                List<string> lstMobiles = new List<string>();
                foreach (var item in lstNoDisputeCustID)
                {
                    var dt = dt1.Select("F2 = '" + item + "'").CopyToDataTable().DefaultView.ToTable(true, "F14");
                    lstMobiles.Add(dt.Rows[0][0].ToString());
                }
                lstMobiles.Remove("Mobile");

                List<string> lstBB = new List<string>();
                foreach (var item in lstNoDisputeCustID)
                {
                    var dt = dt1.Select("F2 = '" + item + "'").CopyToDataTable().DefaultView.ToTable(true, "F13");
                    lstBB.Add(dt.Rows[0][0].ToString());
                }
                lstBB.Remove("BB");

                List<string> lstFixedLine = new List<string>();
                foreach (var item in lstNoDisputeCustID)
                {
                    var dt = dt1.Select("F2 = '" + item + "'").CopyToDataTable().DefaultView.ToTable(true, dt1.Columns[11].ColumnName);
                    lstFixedLine.Add(dt.Rows[0][0].ToString());
                }
                lstFixedLine.Remove("Fixed Line");

                List<string> lstLastPaymentDate = new List<string>();
                foreach (var item in lstNoDisputeCustID)
                {
                    var dt = dt1.Select("F2 = '" + item + "'").CopyToDataTable().DefaultView.ToTable(true, "F18");
                    lstLastPaymentDate.Add(dt.Rows[0][0].ToString());
                }
                lstLastPaymentDate.Remove("Last Payment Date");
                var stringList = lstLastPaymentDate;

                List<string> lstCustName = new List<string>();
                foreach (var item in lstNoDisputeCustID)
                {
                    var dt = dt1.Select("F2 = '" + item + "'").CopyToDataTable().DefaultView.ToTable(true, "F3");
                    lstCustName.Add(dt.Rows[0][0].ToString());
                }
                lstCustName.Remove("Customer Name");
                var stringListCustName = lstCustName;

                List<string> lstBillDate = new List<string>();
                foreach (var item in lstNoDisputeCustID)
                {
                    var dt = dt2.Select("[Customer Id] = '" + item + "'").CopyToDataTable().DefaultView.ToTable(true, "Bill Date");
                    lstBillDate.Add(dt.Rows[0][0].ToString());
                }
                lstBillDate.Remove("Bill Date");
                var stringListBillDate = lstBillDate;

                List<string> lstEmailID = new List<string>();
                foreach (var item in lstNoDisputeCustID)
                {
                    var dt = dt1.Select("F2='" + item + "'").CopyToDataTable().DefaultView.ToTable(true, "F19");
                    lstEmailID.Add(dt.Rows[0][0].ToString());
                }
                lstEmailID.Remove("Contact e-mail");
                var stringListEmailID = lstEmailID;

                List<string> lstBill = new List<string>();
                foreach (var item in lstNoDisputeCustID)
                {
                    var dt = dt2.Select("[Customer Id]='" + item + "'").CopyToDataTable().DefaultView.ToTable(true, "Bill Amount (₹)");
                    lstBill.Add(dt.Rows[0][0].ToString());
                }
                //var intListBill = lstBill.OfType<double>().ToList();

                List<string> lstDeposits = new List<string>();
                foreach (var item in lstNoDisputeCustID)
                {
                    var dt = dt1.Select("F2='" + item + "'").CopyToDataTable().DefaultView.ToTable(true, "F15");
                    lstDeposits.Add(dt.Rows[0][0].ToString());
                }
                lstDeposits.Remove("Security Deposit Amount (₹)");
                var stringListDeposits = lstDeposits;

                List<string> lstCRLimits = new List<string>();
                foreach (var item in lstNoDisputeCustID)
                {
                    var dt = dt1.Select("F2='" + item + "'").CopyToDataTable().DefaultView.ToTable(true, "F16");
                    lstCRLimits.Add(dt.Rows[0][0].ToString());
                }
                lstCRLimits.Remove("Credit Limit (₹)");
                var stringListCRLimits = lstCRLimits;

                List<string> lstConnectionType = new List<string>();
                foreach (var item in lstNoDisputeCustID)
                {
                    var dt = dt1.Select("F2='" + item + "'").CopyToDataTable().DefaultView.ToTable(true, "F11");
                    lstConnectionType.Add(dt.Rows[0][0].ToString());
                }
                lstConnectionType.Remove("Connection Type");
                var stringListConnectionType = lstConnectionType;

                 List<string> lstCostumerCategory = new List<string>();
                foreach (var item in lstNoDisputeCustID)
                {
                    var dt = dt1.Select("F2='" + item + "'").CopyToDataTable().DefaultView.ToTable(true, "F10");
                    lstCostumerCategory.Add(dt.Rows[0][0].ToString());
                }
                lstCostumerCategory.Remove("Connection Type");


                List<string> lstDefaults = new List<string>();
                foreach (var item in lstNoDisputeCustID)
                {
                    var dt = dt1.Select("F2='" + item + "'").CopyToDataTable().DefaultView.ToTable(true, "F17");
                    lstDefaults.Add(dt.Rows[0][0].ToString());
                }
                lstDefaults.Remove("Defaults / Year");

                List<string> lstCustomerType = new List<string>();
                foreach (var item in lstNoDisputeCustID)
                {
                    var dt = dt1.Select("F2='" + item + "'").CopyToDataTable().DefaultView.ToTable(true, "F4");
                    lstCustomerType.Add(dt.Rows[0][0].ToString());
                }
                lstCustomerType.Remove("Customer Type");

                List<string> lstAVGR = new List<string>();
                foreach (var item in lstNoDisputeCustID)
                {
                    var dt = dt1.Select("F2='" + item + "'").CopyToDataTable().DefaultView.ToTable(true, "F8");
                    lstAVGR.Add(dt.Rows[0][0].ToString());
                }
                lstAVGR.Remove("Avg Revenue Score (AVGR)");

                List<string> lstLOYT = new List<string>();
                foreach (var item in lstNoDisputeCustID)
                {
                    var dt = dt1.Select("F2='" + item + "'").CopyToDataTable().DefaultView.ToTable(true, "F7");
                    lstLOYT.Add(dt.Rows[0][0].ToString());
                }
                lstLOYT.Remove("Loyalty Score (LOYT)");

                List<string> lstGadgets = new List<string>();
                for (int i = 0; i < lstBB.Count(); i++)
                {
                    if(lstBB[i]=="Yes")
                    {
                        lstGadgets.Add("BroadBand");
                    }
                    else if(lstMobiles[i]=="Yes")
                    {
                        lstGadgets.Add("Mobiles");
                    }
                    else if(lstFixedLine[i]=="Yes")
                    {
                        lstGadgets.Add("FixedLine");
                    }
                }
                List<string> lstReminder = new List<string>();
                foreach (var item in lstNoDisputeCustID)
                {
                    var dt = dt2.Select("[Customer Id]='" + item + "'").CopyToDataTable().DefaultView.ToTable(true, "No# of reminders sent");
                    lstReminder.Add(dt.Rows[0][0].ToString());
                }
                double date = 0;
                //SendEmail.Email(3, lstNoDisputeEmails, stringList.ToList(), lstBill, lstReminder, stringList.ToList());

                //lstReminder.Remove("No# of reminders sent");

                //lstGadgets.Remove("");


                //foreach (var item in stringList)
                for (int i = 0; i < stringList.Count; i++)
                    {

                    if (stringList[i] != "")
                    {
                        string dtS = DateTime.ParseExact(stringList[i], "dd/MM/yy", CultureInfo.InvariantCulture).ToShortDateString();
                        //string dtS = string.Format("{0:MM/dd/yy}", Convert.ToDateTime(item).ToShortDateString());
                        string dtNow = string.Format("{0:MM/dd/yyyy}", DateTime.Now);
                        
                        date = (Convert.ToDateTime(dtNow) - Convert.ToDateTime(dtS)).TotalDays;
                        //date = 31;
                        if (date <= 30)
                        {
                            //EM=0;
                            //SendEmail.Email(0,stringListEmailID.ToList());
                           // WriteReminderExcel.UpDateStatus(lstNoDisputeCustID[i], filepath, "Out standing days are less than 30");
                        }
                        else if (date > 30 && date <= 60)
                        {
                            //for (int i = 0; i < lstBill.Count; i++)
                            //{
                                //if(!intListCRLimits.Contains("Credit Limit(₹)"))
                                //{
                                NetAmtOS = Convert.ToDouble(lstBill[i]) - Convert.ToDouble(stringListDeposits[i]) - Convert.ToDouble(stringListCRLimits[i]);
                                //NetAmtOS = 999;
                                //}
                                if (NetAmtOS <= 0)
                                {
                                //EM0
                               // WriteReminderExcel.UpDateStatus(lstNoDisputeCustID[i], filepath, "Net Amount is less than or equal to 0");
                            }
                                else if (0 < NetAmtOS && NetAmtOS <= 1000)
                                {
                                    if (stringListConnectionType[i] == "Private")
                                    {
                                        if (lstCostumerCategory[i] == "CIP")
                                        {
                                            if (Convert.ToInt32(lstDefaults[i]) > 0)
                                            {
                                                //cust
                                                if (lstCustomerType[i].ToUpper().Equals("INDIVIDUAL"))
                                                {
                                                    //avgr
                                                    if (Convert.ToInt32(lstAVGR[i]) <= 3)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute

                                                            //SendEmail.Email(0, lstDisputeEmail);

                                                            SendEmail.Email(3, lstNoDisputeEmails[i], stringList.ToList()[i], lstBill[i],lstReminder[i], stringList.ToList()[i], lstCustName[i],lstGadgets[i]);
                                                            WriteReminderExcel.UpdateReminder(lstNoDisputeCustID[i], filepath, lstReminder[i]);WriteReminderExcel.UpDateStatus(lstNoDisputeCustID[i], filepath, "Reminder Sent");
                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0, lstDisputeEmail);

                                                            SendEmail.Email(4, lstNoDisputeEmails[i], stringList.ToList()[i], lstBill[i], lstReminder[i], stringList.ToList()[i], lstCustName[i], lstGadgets[i]);
                                                            WriteReminderExcel.UpdateReminder(lstNoDisputeCustID[i], filepath, lstReminder[i]);WriteReminderExcel.UpDateStatus(lstNoDisputeCustID[i], filepath, "Reminder Sent");

                                                        }
                                                    }
                                                    else if (3 < Convert.ToInt32(lstAVGR[i]) && Convert.ToInt32(lstAVGR[i]) <= 6)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0, lstDisputeEmail);

                                                            SendEmail.Email(1, lstNoDisputeEmails[i], stringList.ToList()[i], lstBill[i], lstReminder[i], stringList.ToList()[i], lstCustName[i], lstGadgets[i]);
                                                            WriteReminderExcel.UpdateReminder(lstNoDisputeCustID[i], filepath, lstReminder[i]);WriteReminderExcel.UpDateStatus(lstNoDisputeCustID[i], filepath, "Reminder Sent");

                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0, lstDisputeEmail);

                                                            SendEmail.Email(2, lstNoDisputeEmails[i], stringList.ToList()[i], lstBill[i], lstReminder[i], stringList.ToList()[i], lstCustName[i], lstGadgets[i]);
                                                            WriteReminderExcel.UpdateReminder(lstNoDisputeCustID[i], filepath, lstReminder[i]);WriteReminderExcel.UpDateStatus(lstNoDisputeCustID[i], filepath, "Reminder Sent");

                                                        }
                                                    }
                                                    else if (Convert.ToInt32(lstAVGR[i]) > 6)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0, lstDisputeEmail);

                                                            SendEmail.Email(1, lstNoDisputeEmails[i], stringList.ToList()[i], lstBill[i], lstReminder[i], stringList.ToList()[i], lstCustName[i], lstGadgets[i]);
                                                            WriteReminderExcel.UpdateReminder(lstNoDisputeCustID[i], filepath, lstReminder[i]);WriteReminderExcel.UpDateStatus(lstNoDisputeCustID[i], filepath, "Reminder Sent");

                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0, lstDisputeEmail);

                                                            SendEmail.Email(2, lstNoDisputeEmails[i], stringList.ToList()[i], lstBill[i], lstReminder[i], stringList.ToList()[i], lstCustName[i], lstGadgets[i]);
                                                            WriteReminderExcel.UpdateReminder(lstNoDisputeCustID[i], filepath, lstReminder[i]);WriteReminderExcel.UpDateStatus(lstNoDisputeCustID[i], filepath, "Reminder Sent");

                                                        }
                                                    }
                                                }
                                                else if (lstCustomerType[i].ToUpper().Equals("BUSINESS") ||
                                                    lstCustomerType[i].ToUpper().Equals("SERVICE") || lstCustomerType[i].ToUpper().Equals("EMER"))
                                                {
                                                    //avgr
                                                    if (Convert.ToInt32(lstAVGR[i]) <= 3)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0, lstDisputeEmail);

                                                            SendEmail.Email(3, lstNoDisputeEmails[i], stringList.ToList()[i], lstBill[i], lstReminder[i], stringList.ToList()[i], lstCustName[i], lstGadgets[i]);
                                                            WriteReminderExcel.UpdateReminder(lstNoDisputeCustID[i], filepath, lstReminder[i]);WriteReminderExcel.UpDateStatus(lstNoDisputeCustID[i], filepath, "Reminder Sent");

                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0, lstDisputeEmail);

                                                            SendEmail.Email(4, lstNoDisputeEmails[i], stringList.ToList()[i], lstBill[i], lstReminder[i], stringList.ToList()[i], lstCustName[i], lstGadgets[i]);
                                                            WriteReminderExcel.UpdateReminder(lstNoDisputeCustID[i], filepath, lstReminder[i]);WriteReminderExcel.UpDateStatus(lstNoDisputeCustID[i], filepath, "Reminder Sent");

                                                        }
                                                    }
                                                    else if (3 < Convert.ToInt32(lstAVGR[i]) && Convert.ToInt32(lstAVGR[i]) <= 6)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0, lstDisputeEmail);

                                                            SendEmail.Email(1, lstNoDisputeEmails[i], stringList.ToList()[i], lstBill[i], lstReminder[i], stringList.ToList()[i], lstCustName[i], lstGadgets[i]);
                                                            WriteReminderExcel.UpdateReminder(lstNoDisputeCustID[i], filepath, lstReminder[i]);WriteReminderExcel.UpDateStatus(lstNoDisputeCustID[i], filepath, "Reminder Sent");

                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0, lstDisputeEmail);

                                                            SendEmail.Email(2, lstNoDisputeEmails[i], stringList.ToList()[i], lstBill[i], lstReminder[i], stringList.ToList()[i], lstCustName[i], lstGadgets[i]);
                                                            WriteReminderExcel.UpdateReminder(lstNoDisputeCustID[i], filepath, lstReminder[i]);WriteReminderExcel.UpDateStatus(lstNoDisputeCustID[i], filepath, "Reminder Sent");

                                                        }
                                                    }
                                                    else if (Convert.ToInt32(lstAVGR[i]) > 6)
                                                    {
                                                        //SendEmail.Email(0, stringListEmailID.ToList());
                                                    }
                                                }
                                            }
                                            else if (Convert.ToInt32(lstDefaults[i]) == 0)
                                            {
                                                //cust
                                                if (lstCustomerType[i].ToUpper().Equals("INDIVIDUAL"))
                                                {
                                                    //avgr
                                                    if (Convert.ToInt32(lstAVGR[i]) <= 3)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0, lstDisputeEmail);

                                                            SendEmail.Email(1, lstNoDisputeEmails[i], stringList.ToList()[i], lstBill[i], lstReminder[i], stringList.ToList()[i], lstCustName[i], lstGadgets[i]);
                                                            WriteReminderExcel.UpdateReminder(lstNoDisputeCustID[i], filepath, lstReminder[i]);WriteReminderExcel.UpDateStatus(lstNoDisputeCustID[i], filepath, "Reminder Sent");

                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0, lstDisputeEmail);

                                                            SendEmail.Email(2, lstNoDisputeEmails[i], stringList.ToList()[i], lstBill[i], lstReminder[i], stringList.ToList()[i], lstCustName[i], lstGadgets[i]);
                                                            WriteReminderExcel.UpdateReminder(lstNoDisputeCustID[i], filepath, lstReminder[i]);WriteReminderExcel.UpDateStatus(lstNoDisputeCustID[i], filepath, "Reminder Sent");

                                                        }
                                                    }
                                                    else if (3 < Convert.ToInt32(lstAVGR[i]) && Convert.ToInt32(lstAVGR[i]) <= 6)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0, lstDisputeEmail);

                                                            SendEmail.Email(1, lstNoDisputeEmails[i], stringList.ToList()[i], lstBill[i], lstReminder[i], stringList.ToList()[i], lstCustName[i], lstGadgets[i]);
                                                            WriteReminderExcel.UpdateReminder(lstNoDisputeCustID[i], filepath, lstReminder[i]);WriteReminderExcel.UpDateStatus(lstNoDisputeCustID[i], filepath, "Reminder Sent");

                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0, lstDisputeEmail);

                                                            SendEmail.Email(2, lstNoDisputeEmails[i], stringList.ToList()[i], lstBill[i], lstReminder[i], stringList.ToList()[i], lstCustName[i], lstGadgets[i]);
                                                            WriteReminderExcel.UpdateReminder(lstNoDisputeCustID[i], filepath, lstReminder[i]);WriteReminderExcel.UpDateStatus(lstNoDisputeCustID[i], filepath, "Reminder Sent");

                                                        }
                                                    }
                                                    else if (Convert.ToInt32(lstAVGR[i]) > 6)
                                                    {
                                                        //SendEmail.Email(0, lstDisputeEmail);
                                                    }
                                                }
                                                else if (lstCustomerType[i].ToUpper().Equals("BUSINESS") ||
                                                    lstCustomerType[i].ToUpper().Equals("SERVICE") || lstCustomerType[i].ToUpper().Equals("EMER"))
                                                {
                                                    //SendEmail.Email(0, lstDisputeEmail);
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
                                                    if (Convert.ToInt32(lstAVGR[i]) <= 3)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0, lstDisputeEmail);

                                                            SendEmail.Email(3, lstNoDisputeEmails[i], stringList.ToList()[i], lstBill[i], lstReminder[i], stringList.ToList()[i], lstCustName[i], lstGadgets[i]);
                                                            WriteReminderExcel.UpdateReminder(lstNoDisputeCustID[i], filepath, lstReminder[i]);WriteReminderExcel.UpDateStatus(lstNoDisputeCustID[i], filepath, "Reminder Sent");

                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0, lstDisputeEmail);

                                                            SendEmail.Email(4, lstNoDisputeEmails[i], stringList.ToList()[i], lstBill[i], lstReminder[i], stringList.ToList()[i], lstCustName[i], lstGadgets[i]);
                                                            WriteReminderExcel.UpdateReminder(lstNoDisputeCustID[i], filepath, lstReminder[i]);WriteReminderExcel.UpDateStatus(lstNoDisputeCustID[i], filepath, "Reminder Sent");

                                                        }
                                                    }
                                                    else if (3 < Convert.ToInt32(lstAVGR[i]) && Convert.ToInt32(lstAVGR[i]) <= 6)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0, lstDisputeEmail);

                                                            SendEmail.Email(1, lstNoDisputeEmails[i], stringList.ToList()[i], lstBill[i], lstReminder[i], stringList.ToList()[i], lstCustName[i], lstGadgets[i]);
                                                            WriteReminderExcel.UpdateReminder(lstNoDisputeCustID[i], filepath, lstReminder[i]);WriteReminderExcel.UpDateStatus(lstNoDisputeCustID[i], filepath, "Reminder Sent");

                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0, lstDisputeEmail);

                                                            SendEmail.Email(2, lstNoDisputeEmails[i], stringList.ToList()[i], lstBill[i], lstReminder[i], stringList.ToList()[i], lstCustName[i], lstGadgets[i]);
                                                            WriteReminderExcel.UpdateReminder(lstNoDisputeCustID[i], filepath, lstReminder[i]);WriteReminderExcel.UpDateStatus(lstNoDisputeCustID[i], filepath, "Reminder Sent");

                                                        }
                                                    }
                                                    else if (Convert.ToInt32(lstAVGR[i]) > 6)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0, lstDisputeEmail);

                                                            SendEmail.Email(1, lstNoDisputeEmails[i], stringList.ToList()[i], lstBill[i], lstReminder[i], stringList.ToList()[i], lstCustName[i], lstGadgets[i]);
                                                            WriteReminderExcel.UpdateReminder(lstNoDisputeCustID[i], filepath, lstReminder[i]);WriteReminderExcel.UpDateStatus(lstNoDisputeCustID[i], filepath, "Reminder Sent");

                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0, lstDisputeEmail);

                                                            SendEmail.Email(2, lstNoDisputeEmails[i], stringList.ToList()[i], lstBill[i], lstReminder[i], stringList.ToList()[i], lstCustName[i], lstGadgets[i]);
                                                            WriteReminderExcel.UpdateReminder(lstNoDisputeCustID[i], filepath, lstReminder[i]);WriteReminderExcel.UpDateStatus(lstNoDisputeCustID[i], filepath, "Reminder Sent");

                                                        }
                                                    }
                                                }
                                                else if (lstCustomerType[i].ToUpper().Equals("BUSINESS") ||
                                                    lstCustomerType[i].ToUpper().Equals("SERVICE") || lstCustomerType[i].ToUpper().Equals("EMER"))
                                                {
                                                    //avgr
                                                    if (Convert.ToInt32(lstAVGR[i]) <= 3)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0, lstDisputeEmail);

                                                            SendEmail.Email(3, lstNoDisputeEmails[i], stringList.ToList()[i], lstBill[i], lstReminder[i], stringList.ToList()[i], lstCustName[i], lstGadgets[i]);
                                                            WriteReminderExcel.UpdateReminder(lstNoDisputeCustID[i], filepath, lstReminder[i]);WriteReminderExcel.UpDateStatus(lstNoDisputeCustID[i], filepath, "Reminder Sent");

                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0, lstDisputeEmail);

                                                            SendEmail.Email(4, lstNoDisputeEmails[i], stringList.ToList()[i], lstBill[i], lstReminder[i], stringList.ToList()[i], lstCustName[i], lstGadgets[i]);
                                                            WriteReminderExcel.UpdateReminder(lstNoDisputeCustID[i], filepath, lstReminder[i]);WriteReminderExcel.UpDateStatus(lstNoDisputeCustID[i], filepath, "Reminder Sent");

                                                        }
                                                    }
                                                    else if (3 < Convert.ToInt32(lstAVGR[i]) && Convert.ToInt32(lstAVGR[i]) <= 6)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0, lstDisputeEmail);

                                                            SendEmail.Email(1, lstNoDisputeEmails[i], stringList.ToList()[i], lstBill[i], lstReminder[i], stringList.ToList()[i], lstCustName[i], lstGadgets[i]);
                                                            WriteReminderExcel.UpdateReminder(lstNoDisputeCustID[i], filepath, lstReminder[i]);WriteReminderExcel.UpDateStatus(lstNoDisputeCustID[i], filepath, "Reminder Sent");

                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0, lstDisputeEmail);

                                                            SendEmail.Email(2, lstNoDisputeEmails[i], stringList.ToList()[i], lstBill[i], lstReminder[i], stringList.ToList()[i], lstCustName[i], lstGadgets[i]);
                                                            WriteReminderExcel.UpdateReminder(lstNoDisputeCustID[i], filepath, lstReminder[i]);WriteReminderExcel.UpDateStatus(lstNoDisputeCustID[i], filepath, "Reminder Sent");

                                                        }
                                                    }
                                                    else if (Convert.ToInt32(lstAVGR[i]) > 6)
                                                    {
                                                        //SendEmail.Email(0, lstDisputeEmail);
                                                    }
                                                }
                                            }
                                            else if (Convert.ToInt32(lstDefaults[i]) == 0)
                                            {
                                                //cust
                                                if (lstCustomerType[i].ToUpper().Equals("INDIVIDUAL"))
                                                {
                                                    //avgr
                                                    if (Convert.ToInt32(lstAVGR[i]) <= 3)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0, lstDisputeEmail);

                                                            SendEmail.Email(1, lstNoDisputeEmails[i], stringList.ToList()[i], lstBill[i], lstReminder[i], stringList.ToList()[i], lstCustName[i], lstGadgets[i]);
                                                            WriteReminderExcel.UpdateReminder(lstNoDisputeCustID[i], filepath, lstReminder[i]);WriteReminderExcel.UpDateStatus(lstNoDisputeCustID[i], filepath, "Reminder Sent");

                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0, lstDisputeEmail);

                                                            SendEmail.Email(2, lstNoDisputeEmails[i], stringList.ToList()[i], lstBill[i], lstReminder[i], stringList.ToList()[i], lstCustName[i], lstGadgets[i]);
                                                            WriteReminderExcel.UpdateReminder(lstNoDisputeCustID[i], filepath, lstReminder[i]);WriteReminderExcel.UpDateStatus(lstNoDisputeCustID[i], filepath, "Reminder Sent");

                                                        }
                                                    }
                                                    else if (3 < Convert.ToInt32(lstAVGR[i]) && Convert.ToInt32(lstAVGR[i]) <= 6)
                                                    {
                                                        //SendEmail.Email(0, lstDisputeEmail);
                                                    }
                                                    else if (Convert.ToInt32(lstAVGR[i]) > 6)
                                                    {
                                                        //SendEmail.Email(0, lstDisputeEmail);
                                                    }
                                                }
                                                else if (lstCustomerType[i].ToUpper().Equals("BUSINESS") ||
                                                    lstCustomerType[i].ToUpper().Equals("SERVICE") || lstCustomerType[i].ToUpper().Equals("EMER"))
                                                {
                                                    //SendEmail.Email(0, lstDisputeEmail);
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
                                                    if (Convert.ToInt32(lstAVGR[i]) <= 3)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0, lstDisputeEmail);

                                                            SendEmail.Email(3, lstNoDisputeEmails[i], stringList.ToList()[i], lstBill[i], lstReminder[i], stringList.ToList()[i], lstCustName[i], lstGadgets[i]);
                                                            WriteReminderExcel.UpdateReminder(lstNoDisputeCustID[i], filepath, lstReminder[i]);WriteReminderExcel.UpDateStatus(lstNoDisputeCustID[i], filepath, "Reminder Sent");

                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0, lstDisputeEmail);

                                                            SendEmail.Email(4, lstNoDisputeEmails[i], stringList.ToList()[i], lstBill[i], lstReminder[i], stringList.ToList()[i], lstCustName[i], lstGadgets[i]);
                                                            WriteReminderExcel.UpdateReminder(lstNoDisputeCustID[i], filepath, lstReminder[i]);WriteReminderExcel.UpDateStatus(lstNoDisputeCustID[i], filepath, "Reminder Sent");

                                                        }
                                                    }
                                                    else if (3 < Convert.ToInt32(lstAVGR[i]) && Convert.ToInt32(lstAVGR[i]) <= 6)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0, lstDisputeEmail);

                                                            SendEmail.Email(1, lstNoDisputeEmails[i], stringList.ToList()[i], lstBill[i], lstReminder[i], stringList.ToList()[i], lstCustName[i], lstGadgets[i]);
                                                            WriteReminderExcel.UpdateReminder(lstNoDisputeCustID[i], filepath, lstReminder[i]);WriteReminderExcel.UpDateStatus(lstNoDisputeCustID[i], filepath, "Reminder Sent");

                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0, lstDisputeEmail);

                                                            SendEmail.Email(2, lstNoDisputeEmails[i], stringList.ToList()[i], lstBill[i], lstReminder[i], stringList.ToList()[i], lstCustName[i], lstGadgets[i]);
                                                            WriteReminderExcel.UpdateReminder(lstNoDisputeCustID[i], filepath, lstReminder[i]);WriteReminderExcel.UpDateStatus(lstNoDisputeCustID[i], filepath, "Reminder Sent");

                                                        }
                                                    }
                                                    else if (Convert.ToInt32(lstAVGR[i]) > 6)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0, lstDisputeEmail);

                                                            SendEmail.Email(1, lstNoDisputeEmails[i], stringList.ToList()[i], lstBill[i], lstReminder[i], stringList.ToList()[i], lstCustName[i], lstGadgets[i]);
                                                            WriteReminderExcel.UpdateReminder(lstNoDisputeCustID[i], filepath, lstReminder[i]);WriteReminderExcel.UpDateStatus(lstNoDisputeCustID[i], filepath, "Reminder Sent");

                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0, lstDisputeEmail);

                                                            SendEmail.Email(2, lstNoDisputeEmails[i], stringList.ToList()[i], lstBill[i], lstReminder[i], stringList.ToList()[i], lstCustName[i], lstGadgets[i]);
                                                            WriteReminderExcel.UpdateReminder(lstNoDisputeCustID[i], filepath, lstReminder[i]);WriteReminderExcel.UpDateStatus(lstNoDisputeCustID[i], filepath, "Reminder Sent");

                                                        }
                                                    }
                                                }
                                                else if (lstCustomerType[i].ToUpper().Equals("BUSINESS") ||
                                                    lstCustomerType[i].ToUpper().Equals("SERVICE") || lstCustomerType[i].ToUpper().Equals("EMER"))
                                                {
                                                    //avgr
                                                    if (Convert.ToInt32(lstAVGR[i]) <= 3)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0, lstDisputeEmail);

                                                            SendEmail.Email(1, lstNoDisputeEmails[i], stringList.ToList()[i], lstBill[i], lstReminder[i], stringList.ToList()[i], lstCustName[i], lstGadgets[i]);
                                                            WriteReminderExcel.UpdateReminder(lstNoDisputeCustID[i], filepath, lstReminder[i]);WriteReminderExcel.UpDateStatus(lstNoDisputeCustID[i], filepath, "Reminder Sent");

                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0, lstDisputeEmail);

                                                            SendEmail.Email(2, lstNoDisputeEmails[i], stringList.ToList()[i], lstBill[i], lstReminder[i], stringList.ToList()[i], lstCustName[i], lstGadgets[i]);
                                                            WriteReminderExcel.UpdateReminder(lstNoDisputeCustID[i], filepath, lstReminder[i]);WriteReminderExcel.UpDateStatus(lstNoDisputeCustID[i], filepath, "Reminder Sent");

                                                        }
                                                    }
                                                    else if (3 < Convert.ToInt32(lstAVGR[i]) && Convert.ToInt32(lstAVGR[i]) <= 6)
                                                    {
                                                        //SendEmail.Email(0, lstDisputeEmail);
                                                    }
                                                    else if (Convert.ToInt32(lstAVGR[i]) > 6)
                                                    {
                                                        //SendEmail.Email(0, lstDisputeEmail);
                                                    }
                                                }
                                            }
                                            else if (Convert.ToInt32(lstDefaults[i]) == 0)
                                            {
                                                //cust
                                                if (lstCustomerType[i].ToUpper().Equals("INDIVIDUAL"))
                                                {
                                                    //avgr
                                                    if (Convert.ToInt32(lstAVGR[i]) <= 3)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0, lstDisputeEmail);

                                                            //SendEmail.Email(0, lstNoDisputeEmails);
                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0, lstDisputeEmail);

                                                            SendEmail.Email(1, lstNoDisputeEmails[i], stringList.ToList()[i], lstBill[i], lstReminder[i], stringList.ToList()[i], lstCustName[i], lstGadgets[i]);
                                                            WriteReminderExcel.UpdateReminder(lstNoDisputeCustID[i], filepath, lstReminder[i]);WriteReminderExcel.UpDateStatus(lstNoDisputeCustID[i], filepath, "Reminder Sent");

                                                        }
                                                    }
                                                    else if (3 < Convert.ToInt32(lstAVGR[i]) && Convert.ToInt32(lstAVGR[i]) <= 6)
                                                    {
                                                        //SendEmail.Email(0, lstDisputeEmail);
                                                    }
                                                    else if (Convert.ToInt32(lstAVGR[i]) > 6)
                                                    {
                                                        //SendEmail.Email(0, lstDisputeEmail);
                                                    }
                                                }
                                                else if (lstCustomerType[i].ToUpper().Equals("BUSINESS") ||
                                                    lstCustomerType[i].ToUpper().Equals("SERVICE") || lstCustomerType[i].ToUpper().Equals("EMER"))
                                                {
                                                    //SendEmail.Email(0, lstDisputeEmail);
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
                                                    if (Convert.ToInt32(lstAVGR[i]) <= 3)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0, lstDisputeEmail);

                                                            SendEmail.Email(5, lstNoDisputeEmails[i], stringList.ToList()[i], lstBill[i], lstReminder[i], stringList.ToList()[i], lstCustName[i], lstGadgets[i]);
                                                            WriteReminderExcel.UpdateReminder(lstNoDisputeCustID[i], filepath, lstReminder[i]);WriteReminderExcel.UpDateStatus(lstNoDisputeCustID[i], filepath, "Reminder Sent");

                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0, lstDisputeEmail);

                                                            SendEmail.Email(6, lstNoDisputeEmails[i], stringList.ToList()[i], lstBill[i], lstReminder[i], stringList.ToList()[i], lstCustName[i], lstGadgets[i]);
                                                            WriteReminderExcel.UpdateReminder(lstNoDisputeCustID[i], filepath, lstReminder[i]);WriteReminderExcel.UpDateStatus(lstNoDisputeCustID[i], filepath, "Reminder Sent");

                                                        }
                                                    }
                                                    else if (3 < Convert.ToInt32(lstAVGR[i]) && Convert.ToInt32(lstAVGR[i]) <= 6)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0, lstDisputeEmail);

                                                            SendEmail.Email(3, lstNoDisputeEmails[i], stringList.ToList()[i], lstBill[i], lstReminder[i], stringList.ToList()[i], lstCustName[i], lstGadgets[i]);
                                                            WriteReminderExcel.UpdateReminder(lstNoDisputeCustID[i], filepath, lstReminder[i]);WriteReminderExcel.UpDateStatus(lstNoDisputeCustID[i], filepath, "Reminder Sent");

                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0, lstDisputeEmail);

                                                            SendEmail.Email(4, lstNoDisputeEmails[i], stringList.ToList()[i], lstBill[i], lstReminder[i], stringList.ToList()[i], lstCustName[i], lstGadgets[i]);
                                                            WriteReminderExcel.UpdateReminder(lstNoDisputeCustID[i], filepath, lstReminder[i]);WriteReminderExcel.UpDateStatus(lstNoDisputeCustID[i], filepath, "Reminder Sent");

                                                        }
                                                    }
                                                    else if (Convert.ToInt32(lstAVGR[i]) > 6)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0, lstDisputeEmail);

                                                            SendEmail.Email(3, lstNoDisputeEmails[i], stringList.ToList()[i], lstBill[i], lstReminder[i], stringList.ToList()[i], lstCustName[i], lstGadgets[i]);
                                                            WriteReminderExcel.UpdateReminder(lstNoDisputeCustID[i], filepath, lstReminder[i]);WriteReminderExcel.UpDateStatus(lstNoDisputeCustID[i], filepath, "Reminder Sent");

                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0, lstDisputeEmail);

                                                            SendEmail.Email(4, lstNoDisputeEmails[i], stringList.ToList()[i], lstBill[i], lstReminder[i], stringList.ToList()[i], lstCustName[i], lstGadgets[i]);
                                                            WriteReminderExcel.UpdateReminder(lstNoDisputeCustID[i], filepath, lstReminder[i]);WriteReminderExcel.UpDateStatus(lstNoDisputeCustID[i], filepath, "Reminder Sent");

                                                        }
                                                    }
                                                }
                                                else if (lstCustomerType[i].ToUpper().Equals("BUSINESS") ||
                                                    lstCustomerType[i].ToUpper().Equals("SERVICE") || lstCustomerType[i].ToUpper().Equals("EMER"))
                                                {
                                                    //avgr
                                                    if (Convert.ToInt32(lstAVGR[i]) <= 3)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0, lstDisputeEmail);

                                                            SendEmail.Email(5, lstNoDisputeEmails[i], stringList.ToList()[i], lstBill[i], lstReminder[i], stringList.ToList()[i], lstCustName[i], lstGadgets[i]);
                                                            WriteReminderExcel.UpdateReminder(lstNoDisputeCustID[i], filepath, lstReminder[i]);WriteReminderExcel.UpDateStatus(lstNoDisputeCustID[i], filepath, "Reminder Sent");

                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0, lstDisputeEmail);

                                                            SendEmail.Email(6, lstNoDisputeEmails[i], stringList.ToList()[i], lstBill[i], lstReminder[i], stringList.ToList()[i], lstCustName[i], lstGadgets[i]);
                                                            WriteReminderExcel.UpdateReminder(lstNoDisputeCustID[i], filepath, lstReminder[i]);WriteReminderExcel.UpDateStatus(lstNoDisputeCustID[i], filepath, "Reminder Sent");

                                                        }
                                                    }
                                                    else if (3 < Convert.ToInt32(lstAVGR[i]) && Convert.ToInt32(lstAVGR[i]) <= 6)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0, lstDisputeEmail);

                                                            SendEmail.Email(3, lstNoDisputeEmails[i], stringList.ToList()[i], lstBill[i], lstReminder[i], stringList.ToList()[i], lstCustName[i], lstGadgets[i]);
                                                            WriteReminderExcel.UpdateReminder(lstNoDisputeCustID[i], filepath, lstReminder[i]);WriteReminderExcel.UpDateStatus(lstNoDisputeCustID[i], filepath, "Reminder Sent");

                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0, lstDisputeEmail);

                                                            SendEmail.Email(4, lstNoDisputeEmails[i], stringList.ToList()[i], lstBill[i], lstReminder[i], stringList.ToList()[i], lstCustName[i], lstGadgets[i]);
                                                            WriteReminderExcel.UpdateReminder(lstNoDisputeCustID[i], filepath, lstReminder[i]);WriteReminderExcel.UpDateStatus(lstNoDisputeCustID[i], filepath, "Reminder Sent");

                                                        }
                                                    }
                                                    else if (Convert.ToInt32(lstAVGR[i]) > 6)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0, lstDisputeEmail);

                                                            SendEmail.Email(1, lstNoDisputeEmails[i], stringList.ToList()[i], lstBill[i], lstReminder[i], stringList.ToList()[i], lstCustName[i], lstGadgets[i]);
                                                            WriteReminderExcel.UpdateReminder(lstNoDisputeCustID[i], filepath, lstReminder[i]);WriteReminderExcel.UpDateStatus(lstNoDisputeCustID[i], filepath, "Reminder Sent");

                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                           // SendEmail.Email(0, lstDisputeEmail);

                                                            SendEmail.Email(2, lstNoDisputeEmails[i], stringList.ToList()[i], lstBill[i], lstReminder[i], stringList.ToList()[i], lstCustName[i], lstGadgets[i]);
                                                            WriteReminderExcel.UpdateReminder(lstNoDisputeCustID[i], filepath, lstReminder[i]);WriteReminderExcel.UpDateStatus(lstNoDisputeCustID[i], filepath, "Reminder Sent");

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
                                                    if (Convert.ToInt32(lstAVGR[i]) <= 3)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                           // SendEmail.Email(0, lstDisputeEmail);

                                                            SendEmail.Email(5, lstNoDisputeEmails[i], stringList.ToList()[i], lstBill[i], lstReminder[i], stringList.ToList()[i], lstCustName[i], lstGadgets[i]);
                                                            WriteReminderExcel.UpdateReminder(lstNoDisputeCustID[i], filepath, lstReminder[i]);WriteReminderExcel.UpDateStatus(lstNoDisputeCustID[i], filepath, "Reminder Sent");

                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0, lstDisputeEmail);

                                                            SendEmail.Email(6, lstNoDisputeEmails[i], stringList.ToList()[i], lstBill[i], lstReminder[i], stringList.ToList()[i], lstCustName[i], lstGadgets[i]);
                                                            WriteReminderExcel.UpdateReminder(lstNoDisputeCustID[i], filepath, lstReminder[i]);WriteReminderExcel.UpDateStatus(lstNoDisputeCustID[i], filepath, "Reminder Sent");

                                                        }
                                                    }
                                                    else if (3 < Convert.ToInt32(lstAVGR[i]) && Convert.ToInt32(lstAVGR[i]) <= 6)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                           // SendEmail.Email(0, lstDisputeEmail);

                                                            SendEmail.Email(3, lstNoDisputeEmails[i], stringList.ToList()[i], lstBill[i], lstReminder[i], stringList.ToList()[i], lstCustName[i], lstGadgets[i]);
                                                            WriteReminderExcel.UpdateReminder(lstNoDisputeCustID[i], filepath, lstReminder[i]);WriteReminderExcel.UpDateStatus(lstNoDisputeCustID[i], filepath, "Reminder Sent");

                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0, lstDisputeEmail);

                                                            SendEmail.Email(4, lstNoDisputeEmails[i], stringList.ToList()[i], lstBill[i], lstReminder[i], stringList.ToList()[i], lstCustName[i], lstGadgets[i]);
                                                            WriteReminderExcel.UpdateReminder(lstNoDisputeCustID[i], filepath, lstReminder[i]);WriteReminderExcel.UpDateStatus(lstNoDisputeCustID[i], filepath, "Reminder Sent");

                                                        }
                                                    }
                                                    else if (Convert.ToInt32(lstAVGR[i]) > 6)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0, lstDisputeEmail);

                                                            SendEmail.Email(3, lstNoDisputeEmails[i], stringList.ToList()[i], lstBill[i], lstReminder[i], stringList.ToList()[i], lstCustName[i], lstGadgets[i]);
                                                            WriteReminderExcel.UpdateReminder(lstNoDisputeCustID[i], filepath, lstReminder[i]);WriteReminderExcel.UpDateStatus(lstNoDisputeCustID[i], filepath, "Reminder Sent");

                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0, lstDisputeEmail);

                                                            SendEmail.Email(4, lstNoDisputeEmails[i], stringList.ToList()[i], lstBill[i], lstReminder[i], stringList.ToList()[i], lstCustName[i], lstGadgets[i]);
                                                            WriteReminderExcel.UpdateReminder(lstNoDisputeCustID[i], filepath, lstReminder[i]);WriteReminderExcel.UpDateStatus(lstNoDisputeCustID[i], filepath, "Reminder Sent");

                                                        }
                                                    }
                                                }
                                                else if (lstCustomerType[i].ToUpper().Equals("BUSINESS") ||
                                                    lstCustomerType[i].ToUpper().Equals("SERVICE") || lstCustomerType[i].ToUpper().Equals("EMER"))
                                                {
                                                    //avgr
                                                    if (Convert.ToInt32(lstAVGR[i]) <= 3)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0, lstDisputeEmail);

                                                            SendEmail.Email(5, lstNoDisputeEmails[i], stringList.ToList()[i], lstBill[i], lstReminder[i], stringList.ToList()[i], lstCustName[i], lstGadgets[i]);
                                                            WriteReminderExcel.UpdateReminder(lstNoDisputeCustID[i], filepath, lstReminder[i]);WriteReminderExcel.UpDateStatus(lstNoDisputeCustID[i], filepath, "Reminder Sent");

                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0, lstDisputeEmail);

                                                            SendEmail.Email(6, lstNoDisputeEmails[i], stringList.ToList()[i], lstBill[i], lstReminder[i], stringList.ToList()[i], lstCustName[i], lstGadgets[i]);
                                                            WriteReminderExcel.UpdateReminder(lstNoDisputeCustID[i], filepath, lstReminder[i]);WriteReminderExcel.UpDateStatus(lstNoDisputeCustID[i], filepath, "Reminder Sent");

                                                        }
                                                    }
                                                    else if (3 < Convert.ToInt32(lstAVGR[i]) && Convert.ToInt32(lstAVGR[i]) <= 6)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0, lstDisputeEmail);

                                                            SendEmail.Email(3, lstNoDisputeEmails[i], stringList.ToList()[i], lstBill[i], lstReminder[i], stringList.ToList()[i], lstCustName[i], lstGadgets[i]);
                                                            WriteReminderExcel.UpdateReminder(lstNoDisputeCustID[i], filepath, lstReminder[i]);WriteReminderExcel.UpDateStatus(lstNoDisputeCustID[i], filepath, "Reminder Sent");

                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0, lstDisputeEmail);

                                                            SendEmail.Email(4, lstNoDisputeEmails[i], stringList.ToList()[i], lstBill[i], lstReminder[i], stringList.ToList()[i], lstCustName[i], lstGadgets[i]);
                                                            WriteReminderExcel.UpdateReminder(lstNoDisputeCustID[i], filepath, lstReminder[i]);WriteReminderExcel.UpDateStatus(lstNoDisputeCustID[i], filepath, "Reminder Sent");

                                                        }
                                                    }
                                                    else if (Convert.ToInt32(lstAVGR[i]) > 6)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0, lstDisputeEmail);

                                                            SendEmail.Email(1, lstNoDisputeEmails[i], stringList.ToList()[i], lstBill[i], lstReminder[i], stringList.ToList()[i], lstCustName[i], lstGadgets[i]);
                                                            WriteReminderExcel.UpdateReminder(lstNoDisputeCustID[i], filepath, lstReminder[i]);WriteReminderExcel.UpDateStatus(lstNoDisputeCustID[i], filepath, "Reminder Sent");

                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0, lstDisputeEmail);

                                                            SendEmail.Email(2, lstNoDisputeEmails[i], stringList.ToList()[i], lstBill[i], lstReminder[i], stringList.ToList()[i], lstCustName[i], lstGadgets[i]);
                                                            WriteReminderExcel.UpdateReminder(lstNoDisputeCustID[i], filepath, lstReminder[i]);WriteReminderExcel.UpDateStatus(lstNoDisputeCustID[i], filepath, "Reminder Sent");

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
                                else if (1000 < NetAmtOS && NetAmtOS <= 3000)
                                {

                                }
                                else if (3000 < NetAmtOS && NetAmtOS <= 10000)
                                {

                                }
                                else if (10000 < NetAmtOS && NetAmtOS <= 25000)
                                {

                                }
                                else if (25000 < NetAmtOS && NetAmtOS <= 50000)
                                {

                                }
                                else if (50000 < NetAmtOS && NetAmtOS <= 100000)
                                {

                                }
                                else if (100000 < NetAmtOS && NetAmtOS <= 500000)
                                {

                                }
                                else if (NetAmtOS > 500000)
                                {

                                }

                            }
                        }
                        else if (60 < date && date <= 90)
                        {

                        }
                        else if (date > 90)
                        {

                        }
                    }
                MessageBox.Show("Process Complete!");
                //}


            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.ToString());
            }

        }
    }
}
