using System;
using System.Data;
using System.Configuration;
using System.Web;
//using System.Web.Security;
//using System.Web.UI;
//using System.Web.UI.WebControls;
//using System.Web.UI.WebControls.WebParts;
//using System.Web.UI.HtmlControls;
using System.Data.OleDb;

/// <summary>
/// Summary description for DbFunction
/// </summary>
/// 
namespace USEPMS
{
    public class DbFunction
    {
        static OleDbConnection ocon = new OleDbConnection(NDS.con());
        OleDbTransaction dbtrans;
        OleDbCommand dbcommand;
        OleDbDataAdapter da;
       // DataSet ds;
        DataTable dt;

        public DbFunction()
        {

        }
        public void dbbegintrans()
        {
            if (ocon.State == ConnectionState.Closed)
            {
                ocon.Open();
            }
            dbtrans = ocon.BeginTransaction();
        }
        public void dbcommittrans()
        {
            try
            {
                if (ocon.State == ConnectionState.Closed)
                {
                    ocon.Open();
                }
                dbtrans.Commit();
            }
            catch (Exception ex)
            {
                string str = ex.Message;
            }
        }
        public DataTable dmlgetquery(string sql)
        {
            try
            {
                ocon.Close();
                if (ocon.State == ConnectionState.Closed)
                {
                    ocon.Open();
                }
                dbcommand = new OleDbCommand();
                dbcommand.Connection = ocon;
                da = new OleDbDataAdapter();
                da.SelectCommand = dbcommand;
                dt = null;
                dt = new DataTable();
                dbcommand.CommandType = CommandType.Text;
                dbcommand.CommandText = sql;
                dbcommand.Transaction = dbtrans;
                da.Fill(dt);
                return dt;
            }
            catch(Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (ocon.State == ConnectionState.Open)
                {
                    ocon.Close();
                }
            }
        }
        public int dmlsinglequery(string sql)
        {
            int Result=0;
            try
            {
                if (ocon.State == ConnectionState.Closed)
                {
                    ocon.Open();
                }
                dbcommand = new OleDbCommand(sql, ocon);
                dbcommand.Transaction = dbtrans;
                Result = dbcommand.ExecuteNonQuery();
                return Result;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (ocon.State == ConnectionState.Open)
                {
                    ocon.Close();
                }
            }
        }
    }
}
