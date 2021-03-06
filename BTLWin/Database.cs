using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Data
{
    class Database
    {
        private const string StringConnection = @"Data Source=LAPTOP-HTD0S059;Initial Catalog=QLDSV;Integrated Security=True";
        private SqlConnection conn;
        private SqlCommand cmd;
        private DataTable dt;
        
        public Database()
        {
            conn = new SqlConnection(StringConnection);
            if(conn == null)
            {
               MessageBox.Show("Không thể kết nối được với SQLServer", "SqlConnection Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            conn.Close();
        }
        public DataTable SelectData(string query)
        {
            try
            {
                conn.Open();
                cmd = new SqlCommand(query, conn);
                dt = new DataTable();
                dt.Load(cmd.ExecuteReader());
                return dt;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi : " + ex.Message, "Excute Reader Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
            finally
            {
                conn.Close();
            }
        }
        public List<string> ExecCmd(string query)
        {
            List<string> result = new List<string>(2);
            try
            {               
                conn.Open();
                cmd = new SqlCommand(query, conn);
                result.Add(cmd.ExecuteNonQuery().ToString());
                return result; 
            }catch(Exception ex)
            {
                result.Add("0");
                result.Add(ex.Message);
                return result;
            }
            finally
            {
                conn.Close();
            }
        }

    }
}
