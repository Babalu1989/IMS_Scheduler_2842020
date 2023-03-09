using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Collections;
using System.ComponentModel;
using System.Data;
//using System.Web.UI;
using System.Data.OleDb;
using System.Text;

/// <summary>
/// Summary description for NDS
/// </summary>
public class NDS
{
    public NDS()
    {

    }
    public static string con()
    {
        string Strcon = string.Empty;
        string Database = string.Empty;
        string Userid = string.Empty;
        string Pass = string.Empty;
        string vs = string.Empty;
        try
        {
            //Cryptograph crp = new Cryptograph();
            //vs = AppDomain.CurrentDomain.BaseDirectory + "bses.ini";
            //string PW_KEY = "o8??^am(*)";
            //Userid = crp.Decrypt(NDSINI.GetINI(vs, "RECOVERY_LIVE", crp.Encrypt("dbuserid", PW_KEY), "?"), PW_KEY);
            //Pass = crp.Decrypt(NDSINI.GetINI(vs, "RECOVERY_LIVE", crp.Encrypt("dbuserpwd", PW_KEY), "?"), PW_KEY);
            //Database = crp.Decrypt(NDSINI.GetINI(vs, "RECOVERY_LIVE", crp.Encrypt("dbconn", PW_KEY), "?"), PW_KEY);
            //Strcon = "Provider=MSDAORA.1; User ID=" + Userid + "; Password=" + Pass + "; Data Source=" + Database + ";";
            //Strcon = "provider=oraoledb.oracle; User ID=itinv; Password=itinv; Data Source=ebstestold;";
            //Strcon = "provider=oraoledb.oracle; User ID=piyush; Password=piyush; Data Source=ebsdbstd;";
            Strcon = "Provider=MSDAORA.1; User ID=itinv; Password=zw39hi; Data Source=ebsdb;";
        }
        catch (Exception ex)
        {
            throw ex;
        }
        return Strcon;
    }
}