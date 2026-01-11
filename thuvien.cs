using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace projectQLTV
{
    internal class thuvien
    {
        public static SqlConnection con = new SqlConnection("Data Source=DUC;Initial Catalog=dbqltv;Integrated Security=True");
        public static void Hienthi(DataGridView dgv, SqlCommand cmd)
        {
            if (thuvien.con.State == ConnectionState.Closed)
            {
                thuvien.con.Open();
            }

            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataTable tb = new DataTable();
            da.Fill(tb);
            cmd.Dispose();
            con.Close();
            dgv.DataSource = tb;
            dgv.Refresh();
        }
        public static void hienthicbo(ComboBox cbo, string tenbang, string ma, string ten)
        {
            if (thuvien.con.State == ConnectionState.Closed)
                thuvien.con.Open();
            string sql = "Select * from " + tenbang;
            SqlCommand cmd = new SqlCommand(sql, con);
            SqlDataAdapter da = new SqlDataAdapter();
            da.SelectCommand = cmd;
            System.Data.DataTable tb = new System.Data.DataTable();
            da.Fill(tb);
            cmd.Dispose();
            con.Close();
            //Thêm 1 dòng vào vị trí đầu tiên của bảng 
            DataRow r = tb.NewRow();
            r[ma] = "";
            r[ten] = "--Chọn giá trị--";
            tb.Rows.InsertAt(r, 0);
            //B5: Đổ dữ liệu vào cbo
            cbo.DataSource = tb;
            cbo.DisplayMember = ten;
            cbo.ValueMember = ma;
        }
    }
}

