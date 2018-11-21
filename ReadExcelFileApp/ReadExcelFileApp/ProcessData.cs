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
        public static void processData(DataTable dt1, DataTable dt2, DataTable dt3)
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
                    var dt = dt1.Select("F2 = '" + item + "'").CopyToDataTable().DefaultView.ToTable(true, "F19");
                    lstDisputeEmail.Add(dt.Rows[0][0].ToString());
                }
                List<string> lstNoDisputeEmails = new List<string>();
                var lst = stringList1.Except(lstDisputesCust).ToList();
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
                            //SendEmail.Email(0,stringListEmailID.ToList());
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
                                    //SendEmail.Email(0,stringListEmailID.ToList());
                                }
                                else if (0 < NetAmtOS && NetAmtOS < 1000)
                                {
                                    if (stringListConnectionType[i] == "Private")
                                    {
                                        if (lstCostumerCategory[i + 1] == "CIP")
                                        {
                                            if (Convert.ToInt32(lstDefaults[i + 1]) > 0)
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

                                                            //SendEmail.Email(0,lstDisputeEmail);

                                                            //SendEmail.Email(3,lstNoDisputeEmails);

                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0,lstDisputeEmail);

                                                            //SendEmail.Email(4,lstNoDisputeEmails);
                                                        }
                                                    }
                                                    else if (3 < Convert.ToInt32(lstAVGR[i]) && Convert.ToInt32(lstAVGR[i]) < 6)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0,lstDisputeEmail);

                                                            //SendEmail.Email(1,lstNoDisputeEmails);
                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0,lstDisputeEmail);

                                                            //SendEmail.Email(2,lstNoDisputeEmails);
                                                        }
                                                    }
                                                    else if (Convert.ToInt32(lstAVGR[i]) > 6)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0,lstDisputeEmail);

                                                            //SendEmail.Email(1,lstNoDisputeEmails);
                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0,lstDisputeEmail);

                                                            //SendEmail.Email(2,lstNoDisputeEmails);
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
                                                            //SendEmail.Email(0,lstDisputeEmail);

                                                            //SendEmail.Email(3,lstNoDisputeEmails);
                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0,lstDisputeEmail);

                                                            //SendEmail.Email(4,lstNoDisputeEmails);
                                                        }
                                                    }
                                                    else if (3 < Convert.ToInt32(lstAVGR[i]) && Convert.ToInt32(lstAVGR[i]) < 6)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0,lstDisputeEmail);

                                                            //SendEmail.Email(1,lstNoDisputeEmails);
                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0,lstDisputeEmail);

                                                            //SendEmail.Email(2,lstNoDisputeEmails);
                                                        }
                                                    }
                                                    else if (Convert.ToInt32(lstAVGR[i]) > 6)
                                                    {
                                                        //SendEmail.Email(0,stringListEmailID.ToList());
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
                                                            //SendEmail.Email(0,lstDisputeEmail);

                                                            //SendEmail.Email(1,lstNoDisputeEmails);
                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0,lstDisputeEmail);

                                                            //SendEmail.Email(2,lstNoDisputeEmails);
                                                        }
                                                    }
                                                    else if (3 < Convert.ToInt32(lstAVGR[i]) && Convert.ToInt32(lstAVGR[i]) < 6)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0,lstDisputeEmail);

                                                            //SendEmail.Email(1,lstNoDisputeEmails);
                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0,lstDisputeEmail);

                                                            //SendEmail.Email(2,lstNoDisputeEmails);
                                                        }
                                                    }
                                                    else if (Convert.ToInt32(lstAVGR[i]) > 6)
                                                    {
                                                        //SendEmail.Email(0,lstDisputeEmail);
                                                    }
                                                }
                                                else if (lstCustomerType[i].ToUpper().Equals("BUSINESS") ||
                                                    lstCustomerType[i].ToUpper().Equals("SERVICE") || lstCustomerType[i].ToUpper().Equals("EMER"))
                                                {
                                                    //SendEmail.Email(0,lstDisputeEmail);
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
                                                            //SendEmail.Email(0,lstDisputeEmail);

                                                            //SendEmail.Email(3,lstNoDisputeEmails);
                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0,lstDisputeEmail);

                                                            //SendEmail.Email(4,lstNoDisputeEmails);
                                                        }
                                                    }
                                                    else if (3 < Convert.ToInt32(lstAVGR[i]) && Convert.ToInt32(lstAVGR[i]) < 6)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0,lstDisputeEmail);

                                                            //SendEmail.Email(1,lstNoDisputeEmails);
                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0,lstDisputeEmail);

                                                            //SendEmail.Email(2,lstNoDisputeEmails);
                                                        }
                                                    }
                                                    else if (Convert.ToInt32(lstAVGR[i]) > 6)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0,lstDisputeEmail);

                                                            //SendEmail.Email(1,lstNoDisputeEmails);
                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0,lstDisputeEmail);

                                                            //SendEmail.Email(2,lstNoDisputeEmails);
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
                                                            //SendEmail.Email(0,lstDisputeEmail);

                                                            //SendEmail.Email(3,lstNoDisputeEmails);
                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0,lstDisputeEmail);

                                                            //SendEmail.Email(4,lstNoDisputeEmails);
                                                        }
                                                    }
                                                    else if (3 < Convert.ToInt32(lstAVGR[i]) && Convert.ToInt32(lstAVGR[i]) < 6)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0,lstDisputeEmail);

                                                            //SendEmail.Email(1,lstNoDisputeEmails);
                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0,lstDisputeEmail);

                                                            //SendEmail.Email(2,lstNoDisputeEmails);
                                                        }
                                                    }
                                                    else if (Convert.ToInt32(lstAVGR[i]) > 6)
                                                    {
                                                        //SendEmail.Email(0,lstDisputeEmail);
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
                                                            //SendEmail.Email(0,lstDisputeEmail);

                                                            //SendEmail.Email(1,lstNoDisputeEmails);
                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0,lstDisputeEmail);

                                                            //SendEmail.Email(2,lstNoDisputeEmails);
                                                        }
                                                    }
                                                    else if (3 < Convert.ToInt32(lstAVGR[i]) && Convert.ToInt32(lstAVGR[i]) < 6)
                                                    {
                                                        //SendEmail.Email(0,lstDisputeEmail);
                                                    }
                                                    else if (Convert.ToInt32(lstAVGR[i]) > 6)
                                                    {
                                                        //SendEmail.Email(0,lstDisputeEmail);
                                                    }
                                                }
                                                else if (lstCustomerType[i].ToUpper().Equals("BUSINESS") ||
                                                    lstCustomerType[i].ToUpper().Equals("SERVICE") || lstCustomerType[i].ToUpper().Equals("EMER"))
                                                {
                                                    //SendEmail.Email(0,lstDisputeEmail);
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
                                                            //SendEmail.Email(0,lstDisputeEmail);

                                                            //SendEmail.Email(3,lstNoDisputeEmails);
                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0,lstDisputeEmail);

                                                            //SendEmail.Email(4,lstNoDisputeEmails);
                                                        }
                                                    }
                                                    else if (3 < Convert.ToInt32(lstAVGR[i]) && Convert.ToInt32(lstAVGR[i]) < 6)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0,lstDisputeEmail);

                                                            //SendEmail.Email(1,lstNoDisputeEmails);
                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0,lstDisputeEmail);

                                                            //SendEmail.Email(2,lstNoDisputeEmails);
                                                        }
                                                    }
                                                    else if (Convert.ToInt32(lstAVGR[i]) > 6)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0,lstDisputeEmail);

                                                            //SendEmail.Email(1,lstNoDisputeEmails);
                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0,lstDisputeEmail);

                                                            //SendEmail.Email(2,lstNoDisputeEmails);
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
                                                            //SendEmail.Email(0,lstDisputeEmail);

                                                            //SendEmail.Email(1,lstNoDisputeEmails);
                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0,lstDisputeEmail);

                                                            //SendEmail.Email(2,lstNoDisputeEmails);
                                                        }
                                                    }
                                                    else if (3 < Convert.ToInt32(lstAVGR[i]) && Convert.ToInt32(lstAVGR[i]) < 6)
                                                    {
                                                        //SendEmail.Email(0,lstDisputeEmail);
                                                    }
                                                    else if (Convert.ToInt32(lstAVGR[i]) > 6)
                                                    {
                                                        //SendEmail.Email(0,lstDisputeEmail);
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
                                                            //SendEmail.Email(0,lstDisputeEmail);

                                                            //SendEmail.Email(0,lstNoDisputeEmails);
                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0,lstDisputeEmail);

                                                            //SendEmail.Email(1,lstNoDisputeEmails);
                                                        }
                                                    }
                                                    else if (3 < Convert.ToInt32(lstAVGR[i]) && Convert.ToInt32(lstAVGR[i]) < 6)
                                                    {
                                                        //SendEmail.Email(0,lstDisputeEmail);
                                                    }
                                                    else if (Convert.ToInt32(lstAVGR[i]) > 6)
                                                    {
                                                        //SendEmail.Email(0,lstDisputeEmail);
                                                    }
                                                }
                                                else if (lstCustomerType[i].ToUpper().Equals("BUSINESS") ||
                                                    lstCustomerType[i].ToUpper().Equals("SERVICE") || lstCustomerType[i].ToUpper().Equals("EMER"))
                                                {
                                                    //SendEmail.Email(0,lstDisputeEmail);
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
                                                            //SendEmail.Email(0,lstDisputeEmail);

                                                            //SendEmail.Email(5,lstNoDisputeEmails);
                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0,lstDisputeEmail);

                                                            //SendEmail.Email(6,lstNoDisputeEmails);
                                                        }
                                                    }
                                                    else if (3 < Convert.ToInt32(lstAVGR[i]) && Convert.ToInt32(lstAVGR[i]) < 6)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0,lstDisputeEmail);

                                                            //SendEmail.Email(3,lstNoDisputeEmails);
                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0,lstDisputeEmail);

                                                            //SendEmail.Email(4,lstNoDisputeEmails);
                                                        }
                                                    }
                                                    else if (Convert.ToInt32(lstAVGR[i]) > 6)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0,lstDisputeEmail);

                                                            //SendEmail.Email(3,lstNoDisputeEmails);
                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0,lstDisputeEmail);

                                                            //SendEmail.Email(4,lstNoDisputeEmails);
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
                                                            //SendEmail.Email(0,lstDisputeEmail);

                                                            //SendEmail.Email(5,lstNoDisputeEmails);
                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0,lstDisputeEmail);

                                                            //SendEmail.Email(6,lstNoDisputeEmails);
                                                        }
                                                    }
                                                    else if (3 < Convert.ToInt32(lstAVGR[i]) && Convert.ToInt32(lstAVGR[i]) < 6)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0,lstDisputeEmail);

                                                            //SendEmail.Email(3,lstNoDisputeEmails);
                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0,lstDisputeEmail);

                                                            //SendEmail.Email(4,lstNoDisputeEmails);
                                                        }
                                                    }
                                                    else if (Convert.ToInt32(lstAVGR[i]) > 6)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0,lstDisputeEmail);

                                                            //SendEmail.Email(1,lstNoDisputeEmails);
                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0,lstDisputeEmail);

                                                            //SendEmail.Email(2,lstNoDisputeEmails);
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
                                                            //SendEmail.Email(0,lstDisputeEmail);

                                                            //SendEmail.Email(5,lstNoDisputeEmails);
                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0,lstDisputeEmail);

                                                            //SendEmail.Email(6,lstNoDisputeEmails);
                                                        }
                                                    }
                                                    else if (3 < Convert.ToInt32(lstAVGR[i]) && Convert.ToInt32(lstAVGR[i]) < 6)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0,lstDisputeEmail);

                                                            //SendEmail.Email(3,lstNoDisputeEmails);
                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0,lstDisputeEmail);

                                                            //SendEmail.Email(4,lstNoDisputeEmails);
                                                        }
                                                    }
                                                    else if (Convert.ToInt32(lstAVGR[i]) > 6)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0,lstDisputeEmail);

                                                            //SendEmail.Email(3,lstNoDisputeEmails);
                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0,lstDisputeEmail);

                                                            //SendEmail.Email(4,lstNoDisputeEmails);
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
                                                            //SendEmail.Email(0,lstDisputeEmail);

                                                            //SendEmail.Email(5,lstNoDisputeEmails);
                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0,lstDisputeEmail);

                                                            //SendEmail.Email(6,lstNoDisputeEmails);
                                                        }
                                                    }
                                                    else if (3 < Convert.ToInt32(lstAVGR[i]) && Convert.ToInt32(lstAVGR[i]) < 6)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0,lstDisputeEmail);

                                                            //SendEmail.Email(3,lstNoDisputeEmails);
                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0,lstDisputeEmail);

                                                            //SendEmail.Email(4,lstNoDisputeEmails);
                                                        }
                                                    }
                                                    else if (Convert.ToInt32(lstAVGR[i]) > 6)
                                                    {
                                                        //loyt
                                                        if (Convert.ToInt32(lstLOYT[i]) > 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0,lstDisputeEmail);

                                                            //SendEmail.Email(1,lstNoDisputeEmails);
                                                        }
                                                        else if (Convert.ToInt32(lstLOYT[i]) == 1)
                                                        {
                                                            //dispute
                                                            //SendEmail.Email(0,lstDisputeEmail);

                                                            //SendEmail.Email(2,lstNoDisputeEmails);
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
    }
}
