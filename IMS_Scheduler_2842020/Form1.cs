using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Net.Mail;
using System.Net;

namespace IMS_Scheduler_2842020
{
    public partial class Form1 : Form
    {
        OleDbConnection con;
        OleDbCommand cmd;
        OleDbDataAdapter adp;
        public string strConnection = NDS.con();
        string text = string.Empty, sqlQuery = string.Empty;
        string Empname = string.Empty;
        string Empid = string.Empty;
        public Form1()
        {
            InitializeComponent();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            SendMail();
            Application.Exit();
        }
        private void SendMail()
        {
            using (con = new OleDbConnection(strConnection))
            {
                DataTable dt = new DataTable();
                sqlQuery = "select distinct EMPID ISSUE_CARD_NO, EMAIL from itinv.IMS_EMAIL_MASTER   WHERE COMPANY='BRPL' AND EMAIL_FLAG='N' AND  trunc(ENTRY_DATE) between '29-Jan-2021' AND '13-Feb-2021'";

                //sqlQuery = " select unique ISSUE_CARD_NO from  itinv.ims_issue_info a, itinv.ims_issue_mst b  where a.issue_no = b.issue_no";
                // sqlQuery += " and ISSUE_FLAG = 'Y'  AND EMAIL_FLAG = 'N' and b.ISSUE_CARD_NO='41006019'";//'40133161'";//41003893 reema
                cmd = new OleDbCommand(sqlQuery, con);
                if (con.State == ConnectionState.Closed)
                {
                    con.Open();
                }
                adp = new OleDbDataAdapter(cmd);
                adp.Fill(dt);
                {
                    for (int j = 0; j < dt.Rows.Count; j++)
                    {
                        sqlQuery = " SELECT distinct UPPER(SUBSTR(m.ISSUE_DATE, 1, 11)) AS idate, m.ISSUE_CARD_NO as EMPID,INITCAP(m.ISSUE_TO) AS Name,dp.DEPT_NAME DEPT,";
                        sqlQuery += " d.SDIV_NM ||'-'|| l.LOCN_OFF_NAME AS Location, st.SUB_CTG_DESC AS subctg, i.ITEM_NAME AS Item,(md.MRN_ITEM_MAKE || '/' || md.MRN_ITEM_MODEL) make,";
                        sqlQuery += " f.ISSUE_SERIAL_NO AS Serial,ct.EMAIL_DETAILS FROM itinv.ims_issue_mst m, itinv.ims_dept_mst dp, itinv.ims_issue_info f, itinv.ims_loc_mst l,";
                        sqlQuery += " itinv.division d, itinv.IMS_ITEM_MASTER I, itinv.ims_ctg_mst ct,itinv.ims_subctg_mst st, itinv.ims_mrn_details md,";
                        sqlQuery += " itinv.ims_mrn_info_status ms WHERE m.ISSUE_DEPT_CODE = dp.DEPT_CODE(+) AND st.sub_CTG_CODE = SUBSTR(f.ITEM_CODE, 1, 5)";
                        sqlQuery += " AND ct.ctg_code = SUBSTR(f.item_code, 1, 3) AND m.ISSUE_NO = f.ISSUE_NO AND m.LOCN_CODE = l.LOCN_CODE";
                        sqlQuery += " AND f.ITEM_CODE = i.ITEM_CODE AND SUBSTR(l.LOCN_CODE,1,4)= d.sdo_CD AND f.ISSUE_FLAG = 'Y' AND ms.SERIAL_NO = f.ISSUE_SERIAL_NO";
                        sqlQuery += " AND md.MRN_NO = ms.mrn_no AND md.MRN_ITEM_CODE = f.item_code AND md.MRN_ITEM_CODE = i.item_code AND f.ISSUE_SERIAL_NO not";
                        sqlQuery += " in(SELECT SERIAL_NO FROM itinv.IMS_MRN_INFO_STATUS WHERE STATUS = 'M')  and ISSUE_COMP_NAME='BRPL'  AND EMAIL_FLAG='N'";//AND DEPT_NAME='HR'
                        sqlQuery += " and m.ISSUE_CARD_NO = '" + dt.Rows[j]["ISSUE_CARD_NO"].ToString() + "' order by 1";//41003893 

                        Empid = dt.Rows[j]["ISSUE_CARD_NO"].ToString();
                        DataTable _dtMailDT = new DataTable();
                        cmd = new OleDbCommand(sqlQuery, con);
                        if (con.State == ConnectionState.Closed)
                        {
                            con.Open();
                        }
                        adp = new OleDbDataAdapter(cmd);
                        adp.Fill(_dtMailDT);
                        if (_dtMailDT.Rows.Count > 0)
                        {
                            text = "Dear " + _dtMailDT.Rows[0]["NAME"].ToString() + "";
                            text += "<br /><br /><br />";
                            text += "Team IT is excited to have come up with another initiative to help serve to better";
                            text += "<br /><br />";
                            text += "We are sharing with you the details of the IT Assets issued to you by our Team";
                            text += "<br /><br />";
                            text += "<table width=100% cellspacing='0' border='1'>";
                            text += " <tr align='center' valign='Middle' bgcolor='yellow'>";
                            text += "</tr>";
                            text += " <tr align='center' valign='Middle' bgcolor='lightskyblue'>";
                            text += "<td colspan='6'><font face='Arial' color='MidnightBlue' size='2'> <strong>IT Asset detail for " + _dtMailDT.Rows[0]["NAME"].ToString() + "," + _dtMailDT.Rows[0]["EMPID"].ToString() + "</strong></font></td>";
                            text += "</tr>";
                            text += " <tr align='center' valign='Middle' bgcolor='lightskyblue'>";
                            text += "<td><font face='Arial' color='MidnightBlue' size='2'> <strong>S.No.</strong></font></td>";
                            text += "<td><font face='Arial' color='MidnightBlue' size='2'> <strong>Type of asset</strong></font></td>";
                            text += "<td><font face='Arial' color='MidnightBlue' size='2'> <strong>Serial Number</strong></font></td>";
                            text += "<td><font face='Arial' color='MidnightBlue' size='2'> <strong>Issue Date</strong></font></td>";
                            text += "<td><font face='Arial' color='MidnightBlue' size='2'> <strong>Division</strong></font></td>";
                            text += "<td><font face='Arial' color='MidnightBlue' size='2'> <strong>IT Asset Controller Details</strong></font></td>";
                            text += "</tr>";
                            if (_dtMailDT.Rows.Count > 0)
                            {
                                for (int i = 0; i < _dtMailDT.Rows.Count; i++)
                                {
                                    text += " <tr align='center' valign='Middle' bgcolor='LightCyan'>";
                                    text += "<td><font face='Arial' color='MidnightBlue' size='2'> " + (i + 1).ToString() + "</font></td>";
                                    text += "<td><font face='Arial' color='MidnightBlue' size='2'> " + _dtMailDT.Rows[i][6].ToString() + "</font></td>";
                                    text += "<td  align='left'><font face='Arial' color='MidnightBlue' size='2'> " + _dtMailDT.Rows[i][8].ToString() + "</font></td>";
                                    text += "<td><font face='Arial' color='MidnightBlue' size='2'> " + _dtMailDT.Rows[i][0].ToString() + "</font></td>";
                                    text += "<td><font face='Arial' color='MidnightBlue' size='2'> " + _dtMailDT.Rows[i][4].ToString() + "</font></td>";
                                    text += "<td><font face='Arial' color='MidnightBlue' size='2'> " + _dtMailDT.Rows[i][9].ToString() + "</font></td>";
                                    text += " </tr> ";
                                }
                            }
                            text += "</table>";
                            text += "<br /><br />";
                            text += "<br /><br />";
                            text += "Should you have any further queries or require assistance please reach us through E-mail or Phone at the IT Asset Controller's Details mentioned above";
                            text += "<br /><br />";
                            text += "Thanks & Regards";
                            text += "<br />";
                            text += "Team IT";
                            if (dt.Rows[j]["EMAIL"].ToString().Contains("@"))
                            {
                                SendEmail_smtp_In_SingleMail(dt.Rows[j]["EMAIL"].ToString().Replace(",", ""), "babalu.kumar@relianceada.com", "piyush.ratta@relianceada.com", "IMS Details", text);
                            }
                            // SendEmail_smtp_In_SingleMail("piyush.ratta@relianceada.com", "Sanjeev.Ranjan@relianceada.com", "babalu.kumar@relianceada.com", "IMS Details", text);
                        }
                        UpdateFail_Case();
                        // SendEmail_smtp_In_SingleMail("babalu.kumar@relianceada.com", "babalu.kumar@relianceada.com", "babalu.kumar@relianceada.com", "IMS Details", text);
                    }
                }
            }
        }
        public string SendEmail_smtp_In_SingleMail(string toAddress, string BCCAddress1, string BCCAddress, string subject, string Mailbody)
        {
            string result = "Y";
            try
            {
                MailMessage mail = new MailMessage();
                mail.From = new MailAddress("reldelhi.itsupport@relianceada.com");
                mail.To.Add(toAddress);
               // mail.CC.Add(CCAddress);
                mail.Bcc.Add(BCCAddress1);
                mail.Bcc.Add(BCCAddress);
                mail.Subject = subject;
                mail.Body = Mailbody;
                mail.IsBodyHtml = true;
                SmtpClient smtp = new SmtpClient("10.8.61.84");
                smtp.Port = 25;
                smtp.Send(mail);
                if (result == "Y")
                {
                    sqlQuery = "UPDATE itinv.ims_issue_mst set EMAIL_FLAG='Y' WHERE EMAIL_FLAG='N' AND ISSUE_CARD_NO='" + Empid + "'";
                    if (con.State == ConnectionState.Closed)
                    {
                        con.Open();
                    }
                    cmd = new OleDbCommand(sqlQuery, con);
                    adp = new OleDbDataAdapter(cmd);
                    adp.SelectCommand.CommandText = sqlQuery;
                    int j = cmd.ExecuteNonQuery();
                    if (j > 0)
                    {
                        sqlQuery = "UPDATE  itinv.IMS_EMAIL_MASTER SET EMAIL_FLAG='Y',SEND_MAIL_DATE=SYSDATE where EMAIL_FLAG='N' AND COMPANY='BRPL'  AND EMPID='" + Empid + "'";
                        if (con.State == ConnectionState.Closed)
                        {
                            con.Open();
                        }
                        cmd = new OleDbCommand(sqlQuery, con);
                        adp = new OleDbDataAdapter(cmd);
                        adp.SelectCommand.CommandText = sqlQuery;
                        cmd.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception ex)
            {
                if (con.State == ConnectionState.Open)
                {
                    con.Close();
                    con.Dispose();
                }
                result = "N";
            }
            return result;
        }

        public void UpdateFail_Case()
        {
            try
            {
                sqlQuery = "UPDATE  itinv.IMS_EMAIL_MASTER SET EMAIL_FLAG='F',SEND_MAIL_DATE=SYSDATE where EMAIL_FLAG='N' AND COMPANY='BRPL'  AND EMPID='" + Empid + "'";
                if (con.State == ConnectionState.Closed)
                {
                    con.Open();
                }
                cmd = new OleDbCommand(sqlQuery, con);
                adp = new OleDbDataAdapter(cmd);
                adp.SelectCommand.CommandText = sqlQuery;
                cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                if (con.State == ConnectionState.Open)
                {
                    con.Close();
                    con.Dispose();
                }
            }
        }
    }
}
